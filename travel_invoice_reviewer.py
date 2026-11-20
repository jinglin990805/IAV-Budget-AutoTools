"""差旅发票识别与差旅报销表写入。

这个模块只把上传文件当作业务数据来源：文档中的文字不会被当作程序指令执行。
当前支持可提取文本的 PDF 与 DOCX。扫描型 PDF 会使用本机 Tesseract OCR，识别结果
仍会保留人工复核提示，避免把低置信度的 OCR 结果静默写入报销表。
"""

from __future__ import annotations

import copy
import hashlib
import json
import math
import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
import uuid
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from importlib.metadata import PackageNotFoundError, version as package_version
from pathlib import Path
from typing import Iterable, Optional

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

from document_processing import (
    TesseractProvider,
    TravelInvoice,
    TravelPolicyEngine,
    TravelExpenseTemplateAdapter,
    build_tesseract_router,
    extract_travel_invoice,
    extract_travel_invoices,
    summarise_travel_validation,
    validate_travel_invoice,
)
from document_processing.travel import _invoice_page_fragment
from document_processing.reimbursement_claim import ClaimWorkbookError, load_claim_index
from document_processing.travel_claim_reconciliation import (
    add_reconciliation_sheets,
    reconcile_trips,
)


DEFAULT_TRAVEL_OUTPUT_DIR = r"C:\Users\LinJing(CN-TV-F)\Desktop\Jing\00_IAV_TV\02_Budget Report\CN03-22052\05_Acceptance\2026\TF3_BT"
DEFAULT_TRAVEL_TEMPLATE_NAME = "Vendor_Weekly_Business Trip Expense Claim_IAV_Q3_260824.xlsx"
TRAVEL_POLICY_FILE = Path(__file__).resolve().parent / "data" / "travel_policies.json"
TRAVEL_REOCR_CACHE_DIR = Path(__file__).resolve().parent / "output" / ".travel_ocr_cache"
_PDFIUM_LOCK = threading.RLock()
TRAVEL_REOCR_CACHE_SCHEMA = 4
TRAVEL_REOCR_CACHE_MAX_ENTRIES = 200
TRAVEL_REOCR_CACHE_MAX_BYTES = 256 * 1024 * 1024
TRAVEL_REOCR_CACHE_MAX_AGE_SECONDS = 90 * 24 * 60 * 60
_TRAVEL_REOCR_CACHE_LOCK = threading.RLock()
TRAVEL_PADDLE_REOCR_CACHE_DIR = Path(__file__).resolve().parent / "output" / ".travel_paddle_ocr_cache"
TRAVEL_PADDLE_REOCR_CACHE_SCHEMA = 2
TRAVEL_PADDLE_REOCR_CACHE_MAX_ENTRIES = 200
TRAVEL_PADDLE_REOCR_CACHE_MAX_BYTES = 256 * 1024 * 1024
TRAVEL_PADDLE_REOCR_CACHE_MAX_AGE_SECONDS = 90 * 24 * 60 * 60


DATE_RE = re.compile(r"(?<!\d)(\d{1,2})\s*[./-]\s*(\d{1,2})(?!\d)")
FULL_DATE_RE = re.compile(r"(?<!\d)(20\d{2})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(\d{1,2})(?!\d)")
LOCAL_FULL_DATE_RE = re.compile(r"(?<!\d)(\d{1,2})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(\d{2,4})(?!\d)")
CHINESE_DATE_RE = re.compile(r"(?<!\d)(\d{1,4})\s*月\s*(\d{1,2})\s*[日号]?")
COMPACT_SPAN_RE = re.compile(r"(?<!\d)(\d{2})(\d{2})\s*[-~至]\s*(\d{2})(\d{2})(?!\d)")
COMPACT_DATE_RE = re.compile(r"(?<!\d)(\d{2})(\d{2})(?!\d)")
TRAVEL_TEMPLATE_DATE_RE = re.compile(r"(?P<date>\d{6})(?:_v(?P<version>\d+))?\.(?:xlsx|xlsm)$", re.IGNORECASE)
NUMBER_RE = re.compile(r"(?<![\d.])\d{1,9}(?:,\d{3})*(?:\.\d{1,2})?(?![\d.])")
TIME_RE = re.compile(r"(?<!\d)(?P<hour>[01]?\d|2[0-3])\s*[:：]\s*(?P<minute>[0-5]\d)(?!\d)")
MAX_REASONABLE_EXPENSE = 1_000_000.0
CONFIRMED_SERVICE_DATE_SOURCES = frozenset({
    "SOURCE_TEXT",
    "PAGE_DATE_HEADING",
    "HOTEL_FOLIO",
    "HOTEL_FOLIO_RATE",
    "DIDI_ITINERARY",
    "MANUAL_SOURCE_REVIEW",
})


def find_latest_travel_template(output_dir: Optional[str] = None) -> Optional[str]:
    """Find the newest dated travel workbook in the configured output folder.

    The date in the filename is authoritative. If several versions share the
    same date, ``_v2``/``_v3`` and then the modification time determine which
    workbook becomes the next template.
    """
    folder = Path(output_dir or DEFAULT_TRAVEL_OUTPUT_DIR)
    if not folder.is_dir():
        return None

    candidates = []
    for path in folder.iterdir():
        if not path.is_file() or path.name.startswith("~$"):
            continue
        match = TRAVEL_TEMPLATE_DATE_RE.search(path.name)
        if not match:
            continue
        try:
            file_date = datetime.strptime(match.group("date"), "%y%m%d")
        except ValueError:
            continue
        version = int(match.group("version") or 1)
        try:
            modified = path.stat().st_mtime
        except OSError:
            modified = 0
        candidates.append((file_date, version, modified, path))

    if not candidates:
        return None
    return str(max(candidates, key=lambda item: item[:3])[3])


CITY_MAP = {
    "北京": "Beijing", "北京市": "Beijing", "Beijing": "Beijing",
    "上海": "Shanghai", "上海市": "Shanghai", "Shanghai": "Shanghai",
    "深圳": "Shenzhen", "Shenzhen": "Shenzhen",
    "广州": "Guangzhou", "Guangzhou": "Guangzhou",
    "成都": "Chengdu", "Chengdu": "Chengdu",
    "天津": "Tianjin", "天津市": "Tianjin", "Tianjin": "Tianjin",
    "重庆": "Chongqing", "重庆市": "Chongqing", "Chongqing": "Chongqing",
    "苏州": "Suzhou", "Suzhou": "Suzhou",
    "长沙": "Changsha", "Changsha": "Changsha",
    "武汉": "Wuhan", "Wuhan": "Wuhan",
    "西安": "Xian", "西安市": "Xian", "Xian": "Xian",
    "杭州": "Hangzhou", "Hangzhou": "Hangzhou",
    "南京": "Nanjing", "Nanjing": "Nanjing",
    "珠海": "Zhuhai", "珠海市": "Zhuhai", "Zhuhai": "Zhuhai",
    "南宁": "Nanning", "南宁市": "Nanning", "Nanning": "Nanning",
    "桂林": "Guilin", "桂林市": "Guilin", "Guilin": "Guilin",
    "贵阳": "Guiyang", "贵阳市": "Guiyang", "Guiyang": "Guiyang",
    "郑州": "Zhengzhou", "郑州市": "Zhengzhou", "Zhengzhou": "Zhengzhou",
    "东莞": "Dongguan", "东莞市": "Dongguan", "Dongguan": "Dongguan",
    "青岛": "Qingdao", "青岛市": "Qingdao", "Qingdao": "Qingdao",
    "沈阳": "Shenyang", "沈阳市": "Shenyang", "Shenyang": "Shenyang",
    "宁波": "Ningbo", "宁波市": "Ningbo", "Ningbo": "Ningbo",
    "昆明": "Kunming", "昆明市": "Kunming", "Kunming": "Kunming",
    "无锡": "Wuxi", "无锡市": "Wuxi", "Wuxi": "Wuxi",
    "佛山": "Foshan", "佛山市": "Foshan", "Foshan": "Foshan",
    "合肥": "Hefei", "合肥市": "Hefei", "Hefei": "Hefei",
    "大连": "Dalian", "大连市": "Dalian", "Dalian": "Dalian",
    "福州": "Fuzhou", "福州市": "Fuzhou", "Fuzhou": "Fuzhou",
    "厦门": "Xiamen", "厦门市": "Xiamen", "Xiamen": "Xiamen",
    "哈尔滨": "Harbin", "哈尔滨市": "Harbin", "Harbin": "Harbin",
    "济南": "Jinan", "济南市": "Jinan", "Jinan": "Jinan",
    "温州": "Wenzhou", "温州市": "Wenzhou", "Wenzhou": "Wenzhou",
    "长春": "Changchun", "长春市": "Changchun", "Changchun": "Changchun",
    "泉州": "Quanzhou", "泉州市": "Quanzhou", "Quanzhou": "Quanzhou",
    "石家庄": "Shijiazhuang", "石家庄市": "Shijiazhuang", "Shijiazhuang": "Shijiazhuang",
    "南昌": "Nanchang", "南昌市": "Nanchang", "Nanchang": "Nanchang",
    "金华": "Jinhua", "金华市": "Jinhua", "Jinhua": "Jinhua",
    "常州": "Changzhou", "常州市": "Changzhou", "Changzhou": "Changzhou",
    "南通": "Nantong", "南通市": "Nantong", "Nantong": "Nantong",
    "嘉兴": "Jiaxing", "嘉兴市": "Jiaxing", "Jiaxing": "Jiaxing",
    "太原": "Taiyuan", "太原市": "Taiyuan", "Taiyuan": "Taiyuan",
    "徐州": "Xuzhou", "徐州市": "Xuzhou", "Xuzhou": "Xuzhou",
    "惠州": "Huizhou", "惠州市": "Huizhou", "Huizhou": "Huizhou",
    "中山": "Zhongshan", "中山市": "Zhongshan", "Zhongshan": "Zhongshan", "Zhongshn": "Zhongshan",
    "台州": "Taizhou", "台州市": "Taizhou", "泰州": "Taizhou", "Taizhou": "Taizhou",
    "烟台": "Yantai", "烟台市": "Yantai", "Yantai": "Yantai",
    "兰州": "Lanzhou", "兰州市": "Lanzhou", "Lanzhou": "Lanzhou",
    "绍兴": "Shaoxing", "绍兴市": "Shaoxing", "Shaoxing": "Shaoxing",
    "海口": "Haikou", "海口市": "Haikou", "Haikou": "Haikou",
    "扬州": "Yangzhou", "扬州市": "Yangzhou", "Yangzhou": "Yangzhou",
    "香港": "Hong Kong", "香港特别行政区": "Hong Kong", "Hong Kong": "Hong Kong",
    "澳门": "Macau", "澳门特别行政区": "Macau", "Macau": "Macau", "Macao": "Macau",
    "台湾": "Taiwan", "台湾省": "Taiwan", "Taiwan": "Taiwan",
    "北京燕郊": "Beijing Yanjiao", "燕郊": "Beijing Yanjiao", "Beijing Yanjiao": "Beijing Yanjiao",
    "北京赛摩": "Beijing Saimo", "Beijing Saimo": "Beijing Saimo",
}

# Common IATA codes make ticket-route parsing work when OCR recovers airport
# codes but misses the printed city names. Codes shared by multiple cities are
# deliberately omitted so they cannot assign a confident but wrong city.
AIRPORT_CODE_CITY_MAP = {
    "PEK": "Beijing", "PKX": "Beijing", "SHA": "Shanghai", "PVG": "Shanghai",
    "SZX": "Shenzhen", "CAN": "Guangzhou", "CTU": "Chengdu", "TFU": "Chengdu",
    "HGH": "Hangzhou", "CKG": "Chongqing", "WUH": "Wuhan", "XIY": "Xian",
    "TSN": "Tianjin", "NKG": "Nanjing", "CSX": "Changsha", "CGO": "Zhengzhou",
    "TAO": "Qingdao", "SHE": "Shenyang", "NGB": "Ningbo", "KMG": "Kunming",
    "WUX": "Wuxi", "HFE": "Hefei", "DLC": "Dalian", "FOC": "Fuzhou",
    "XMN": "Xiamen", "HRB": "Harbin", "TNA": "Jinan", "WNZ": "Wenzhou",
    "NNG": "Nanning", "CGQ": "Changchun", "JJN": "Quanzhou", "SJW": "Shijiazhuang",
    "KWE": "Guiyang", "KHN": "Nanchang", "YIW": "Jinhua", "CZX": "Changzhou",
    "NTG": "Nantong", "TYN": "Taiyuan", "XUZ": "Xuzhou", "HUZ": "Huizhou",
    "ZUH": "Zhuhai", "KWL": "Guilin", "HYN": "Taizhou", "YNT": "Yantai",
    "LHW": "Lanzhou", "HAK": "Haikou", "YTY": "Yangzhou", "HKG": "Hong Kong",
    "MFM": "Macau", "TPE": "Taiwan", "TSA": "Taiwan",
}

# Customer travel policy. Any city not present in these lists defaults to
# Layer 2; the workbook lookup table is only used for the visible formulas.
LAYER_1_CITIES = {"Beijing", "Shanghai", "Shenzhen", "Guangzhou"}
LAYER_15_CITIES = {
    "Chengdu", "Hangzhou", "Chongqing", "Wuhan", "Xian", "Suzhou",
    "Tianjin", "Nanjing", "Changsha", "Zhengzhou", "Dongguan", "Qingdao",
    "Shenyang", "Ningbo", "Kunming",
}
LAYER_2_CITIES = {
    "Wuxi", "Foshan", "Hefei", "Dalian", "Fuzhou", "Xiamen", "Harbin",
    "Jinan", "Wenzhou", "Nanning", "Changchun", "Quanzhou", "Shijiazhuang",
    "Guiyang", "Nanchang", "Jinhua", "Changzhou", "Nantong", "Jiaxing",
    "Taiyuan", "Xuzhou", "Huizhou", "Zhuhai", "Zhongshan", "Zhongshn",
    "Taizhou", "Yantai", "Lanzhou", "Shaoxing", "Haikou", "Yangzhou",
}

DEFAULT_TRAVEL_POLICY = {
    "name": "VGC差标",
    "levels": {
        "1": {"label": "Layer 1", "cities": sorted(LAYER_1_CITIES), "accommodation": 600, "allowance": 100},
        "1.5": {"label": "Layer 1.5", "cities": sorted(LAYER_15_CITIES), "accommodation": 400, "allowance": 80},
        "2": {"label": "Layer 2", "cities": sorted(LAYER_2_CITIES), "accommodation": 300, "allowance": 50},
        "international": {"label": "International", "cities": ["Hong Kong", "Macau", "Taiwan"], "accommodation": 700, "allowance": 200},
    },
    "unknown_level": "2",
}


def load_travel_policies() -> dict:
    """Load editable customer travel policies, falling back to built-in VGC."""
    default = {"active_policy": "VGC", "policies": {"VGC": DEFAULT_TRAVEL_POLICY}}
    try:
        with TRAVEL_POLICY_FILE.open("r", encoding="utf-8") as handle:
            value = json.load(handle)
        if isinstance(value, dict) and isinstance(value.get("policies"), dict) and value["policies"]:
            default.update({key: value[key] for key in ("active_policy", "policies") if key in value})
    except (OSError, ValueError, TypeError):
        pass
    if default["active_policy"] not in default["policies"]:
        default["active_policy"] = next(iter(default["policies"]))
    return default


def save_travel_policies(value: dict) -> dict:
    """Validate and persist the editable travel policy configuration."""
    if not isinstance(value, dict) or not isinstance(value.get("policies"), dict) or not value["policies"]:
        raise ValueError("差标配置至少需要包含一个差标")
    policies = value["policies"]
    active = str(value.get("active_policy") or next(iter(policies)))
    if active not in policies:
        raise ValueError("当前差标不存在")
    for policy_name, policy in policies.items():
        if not isinstance(policy, dict):
            raise ValueError(f"差标 {policy_name} 配置无效")
        levels = policy.get("levels")
        if not isinstance(levels, dict):
            raise ValueError(f"差标 {policy_name} 缺少城市等级")
        unknown_level = str(policy.get("unknown_level", "2"))
        if unknown_level not in levels:
            raise ValueError(f"差标 {policy_name} 的未知城市默认等级无效")
        seen_cities: dict[str, str] = {}
        for level in ("1", "1.5", "2", "international"):
            definition = levels.get(level)
            if not isinstance(definition, dict):
                raise ValueError(f"差标 {policy_name} 缺少 Layer {level}")
            try:
                definition["accommodation"] = float(definition.get("accommodation", 0))
                definition["allowance"] = float(definition.get("allowance", 0))
            except (TypeError, ValueError) as exc:
                raise ValueError(f"差标 {policy_name} 的金额必须是数字") from exc
            if (
                not math.isfinite(definition["accommodation"])
                or not math.isfinite(definition["allowance"])
                or definition["accommodation"] < 0
                or definition["allowance"] < 0
            ):
                raise ValueError(f"差标 {policy_name} 的金额必须是非负有限数字")
            cities = definition.get("cities", [])
            if not isinstance(cities, list):
                raise ValueError(f"差标 {policy_name} 的城市必须是列表")
            definition["cities"] = [str(city).strip() for city in cities if str(city).strip()]
            for city in definition["cities"]:
                city_key = city.casefold()
                if city_key in seen_cities:
                    raise ValueError(
                        f"差标 {policy_name} 的城市 {city} 同时出现在 {seen_cities[city_key]} 和 {level}"
                    )
                seen_cities[city_key] = level
    TRAVEL_POLICY_FILE.parent.mkdir(parents=True, exist_ok=True)
    payload = {"active_policy": active, "policies": policies}
    with TRAVEL_POLICY_FILE.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    return payload

NAME_MAP = {
    "邱栋": "Qiu, Dong", "Qiu Dong": "Qiu, Dong", "QiuDong": "Qiu, Dong",
    "张艳伟": "Zhang, Yanwei", "张燕伟": "Zhang, Yanwei",
    "Zhang Yanwei": "Zhang, Yanwei", "ZhangYanwei": "Zhang, Yanwei",
    "孙亚涛": "Sun, Yatao", "Sun Yatao": "Sun, Yatao", "SunYatao": "Sun, Yatao",
    "张然": "Zhang, Ran", "Zhang Ran": "Zhang, Ran", "Zhang_Ran": "Zhang, Ran",
}

SPECIAL_KEYWORDS = (
    "加油", "燃油", "高速", "通行费", "过路费", "停车", "停车费",
    "toll", "fuel", "gasoline", "parking", "highway fee",
)
HOTEL_KEYWORDS = (
    "酒店", "住宿", "客房", "住宿费", "hotel", "accommodation", "lodging",
)
TICKET_KEYWORDS = (
    "机票", "飞机", "航班", "航空", "火车", "高铁", "动车", "车票",
    "铁路", "客运", "flight", "train", "railway", "airline",
)
TAXI_KEYWORDS = (
    "打车", "出租", "网约车", "滴滴", "出行", "的士", "地铁", "单车",
    "自行车", "骑行", "共享单车", "单车租赁", "自行车租赁", "共享电单车",
    "taxi", "uber", "metro", "subway", "bike", "ride",
)


def _normalise_expense_category(category: Optional[str]) -> Optional[str]:
    """Keep travel output to the three business categories used by the claim."""
    if category is None:
        return None
    value = str(category).strip().casefold()
    if not value:
        return None
    if value in {"ticket", "airfare", "flight", "train", "rail", "railway"}:
        return "ticket"
    if value in {"accommodation", "hotel", "lodging", "hotel_expense"}:
        return "accommodation"
    # Toll, parking, fuel, bike rental and legacy/unknown categories belong to
    # the local transport bucket in the three-category travel claim.
    return "taxi"


@dataclass
class Expense:
    category: str
    amount: float
    start: Optional[date]
    end: Optional[date]
    description: str
    source_file: str
    aggregate: bool = False
    accommodation_nights: Optional[int] = None
    accommodation_folio_total: Optional[float] = None
    source_page: Optional[int] = None
    allocation_group: Optional[str] = None
    allocation_anchor: Optional[date] = None
    allocation_splits: list[dict] = field(default_factory=list)
    document_id: str = ""
    correction_records: list[dict] = field(default_factory=list)
    service_date_source: str = "UNCONFIRMED"
    service_date_evidence_page: Optional[int] = None
    service_date_evidence_text: str = ""
    accommodation_nights_source: str = "UNCONFIRMED"
    allocation_group_total: Optional[float] = None
    allocation_group_uses_written_total: bool = False
    inferred_from_document_summary: bool = False

    def __post_init__(self) -> None:
        self.category = _normalise_expense_category(self.category) or "taxi"

    def as_dict(self) -> dict:
        return {
            "document_id": self.document_id,
            "category": self.category,
            "amount": round(float(self.amount), 2),
            "start": self.start.isoformat() if self.start else None,
            "end": self.end.isoformat() if self.end else None,
            "description": self.description,
            "source_file": self.source_file,
            "source_page": self.source_page,
            "aggregate": self.aggregate,
            "accommodation_nights": self.accommodation_nights,
            "accommodation_folio_total": self.accommodation_folio_total,
            "allocation_group": self.allocation_group,
            "allocation_group_total": self.allocation_group_total,
            "allocation_group_uses_written_total": self.allocation_group_uses_written_total,
            "inferred_from_document_summary": self.inferred_from_document_summary,
            "allocation_anchor": self.allocation_anchor.isoformat() if self.allocation_anchor else None,
            "allocation_splits": copy.deepcopy(self.allocation_splits),
            "service_date_source": self.service_date_source,
            "service_date_evidence_page": self.service_date_evidence_page,
            "service_date_evidence_text": self.service_date_evidence_text,
            "accommodation_nights_source": self.accommodation_nights_source,
            "correction_records": copy.deepcopy(self.correction_records),
        }


@dataclass
class AccommodationFolioEvidence:
    """Non-payable hotel folio data used to validate a lodging invoice."""

    start: Optional[date]
    end: Optional[date]
    total: Optional[float]
    nights: Optional[int]
    unit_rate: Optional[float]
    source_page: Optional[int] = None
    raw_text: str = ""
    nights_source: str = "UNCONFIRMED"
    city: str = ""


@dataclass
class DidiTripTableEvidence:
    """A Didi trip-table summary used to date/validate nearby taxi invoices."""

    start: Optional[date]
    end: Optional[date]
    total: Optional[float]
    source_page: Optional[int] = None
    annotated_total: Optional[float] = None
    raw_text: str = ""
    has_separate_taxi_receipt: bool = False


@dataclass
class Trip:
    start: date
    end: date
    name: str
    city: str
    source_file: str
    expenses: list[Expense] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    policy_name: str = "VGC"
    structured_invoices: list[TravelInvoice] = field(default_factory=list)
    policy_trace: list[dict] = field(default_factory=list)
    policy_exceptions: list[dict] = field(default_factory=list)
    policy_summary: dict = field(default_factory=dict)
    correction_records: list[dict] = field(default_factory=list)
    page_coverage: list[dict] = field(default_factory=list)
    allocation_trace: list[dict] = field(default_factory=list)
    allocation_audit: dict = field(default_factory=dict)
    city_by_day: dict[date, str] = field(default_factory=dict)

    def city_for(self, day: date) -> str:
        return self.city_by_day.get(day, self.city)

    @property
    def days(self) -> list[date]:
        return [self.start + timedelta(days=i) for i in range((self.end - self.start).days + 1)]

    @property
    def itinerary_cities(self) -> list[str]:
        values = [self.city_for(day) for day in self.days]
        return list(dict.fromkeys(value for value in values if value))

    @property
    def totals(self) -> dict[str, float]:
        totals: dict[str, float] = defaultdict(float)
        for expense in self.expenses:
            category = _normalise_expense_category(expense.category) or "taxi"
            if category == "taxi" and expense.category == "taxi":
                continue
            totals[category] += expense.amount
        totals["taxi"] += _recognized_taxi_total(self.expenses)
        if "accommodation" in totals:
            totals["accommodation"] = min(
                round(totals["accommodation"], 2),
                _accommodation_limit(self),
            )
        return {k: round(v, 2) for k, v in totals.items()}


def _normalise_text(value: str) -> str:
    value = (
        value.replace("：", ":").replace("，", ",").replace("￥", "¥")
        .replace("－", "-").replace("—", "-").replace("–", "-")
        .replace("／", "/").replace("．", ".").strip()
    )
    # Tesseract commonly separates the decimal point from cents on scanned
    # receipts (for example ``¥2370. 00``).  Joining only around punctuation
    # preserves ordinary word spacing and makes amount parsing deterministic.
    value = re.sub(r"(?<=\d)\s+(?=[.,]\s*\d)", "", value)
    value = re.sub(r"(?<=[.,])\s+(?=\d)", "", value)
    return value


def _date_tokens(text: str, default_year: int) -> list[date]:
    """Extract dates without interpreting invoice numbers as dates."""
    result: list[date] = []
    occupied: list[tuple[int, int]] = []
    for match in FULL_DATE_RE.finditer(text):
        try:
            result.append(date(int(match.group(1)), int(match.group(2)), int(match.group(3))))
            occupied.append(match.span())
        except ValueError:
            continue

    for match in LOCAL_FULL_DATE_RE.finditer(text):
        if any(start <= match.start() < end for start, end in occupied):
            continue
        parsed = _parse_folio_date(match.group(0), default_year)
        if parsed is not None:
            result.append(parsed)
            occupied.append(match.span())

    def is_occupied(pos: int) -> bool:
        return any(start <= pos < end for start, end in occupied)

    for match in COMPACT_SPAN_RE.finditer(text):
        try:
            result.append(date(default_year, int(match.group(1)), int(match.group(2))))
            result.append(date(default_year, int(match.group(3)), int(match.group(4))))
            occupied.append(match.span())
        except ValueError:
            continue

    for match in DATE_RE.finditer(text):
        if is_occupied(match.start()):
            continue
        try:
            first, second = int(match.group(1)), int(match.group(2))
            if first > 12 and second <= 12:
                result.append(date(default_year, second, first))
            else:
                result.append(date(default_year, first, second))
        except ValueError:
            continue
    for match in CHINESE_DATE_RE.finditer(text):
        try:
            # Chinese short dates are normally “8月3日”; a four-digit first
            # group is treated as a year only for the uncommon “2026月8日” form.
            if len(match.group(1)) == 4:
                result.append(date(int(match.group(1)), int(match.group(2)), 1))
            else:
                result.append(date(default_year, int(match.group(1)), int(match.group(2))))
        except ValueError:
            continue
    return result


def _date_span(text: str, default_year: int) -> tuple[Optional[date], Optional[date]]:
    tokens = _date_tokens(text, default_year)
    if not tokens:
        return None, None
    return tokens[0], tokens[1] if len(tokens) > 1 else tokens[0]


def _document_year(text: str, default_year: Optional[int]) -> int:
    match = re.search(r"\b(20\d{2})\b", text)
    if match:
        return int(match.group(1))
    return int(default_year or datetime.now().year)


def _ocr_executable() -> Optional[Path]:
    """Resolve the configured local Tesseract executable.

    The workbench is commonly launched with a different Python installation
    from the one used for the data tools, so OCR must not depend on a Python
    wrapper such as ``pytesseract``.  Calling the local executable also keeps
    uploaded travel documents offline.
    """
    try:
        import config
    except ImportError:  # pragma: no cover - config is present in production
        config = None

    configured = getattr(config, "TESSERACT_EXE", "") if config else ""
    if configured:
        candidate = Path(configured)
        if candidate.is_file():
            return candidate
    discovered = shutil.which("tesseract")
    return Path(discovered) if discovered else None


def _ocr_languages(executable: Path) -> str:
    """Return configured languages that are actually installed in Tesseract."""
    try:
        import config
        configured = str(getattr(config, "OCR_LANGUAGE", "eng") or "eng")
    except ImportError:  # pragma: no cover - config is present in production
        configured = "eng"

    requested = [item for item in re.split(r"[+,;\s]+", configured) if item]
    available = set(TesseractProvider.list_languages(executable, timeout=15))

    selected = [language for language in requested if language in available]
    if selected:
        return "+".join(selected)
    if "eng" in available:
        return "eng"
    return requested[0] if requested else "eng"


def _run_tesseract_image(
    executable: Path,
    image,
    languages: str,
    psm: str,
) -> tuple[str, str]:
    """Run the provider-neutral OCR router while preserving the old API."""
    result = build_tesseract_router(
        executable, languages=languages, timeout=90,
    ).recognize(image, languages=languages, psm=psm)
    return result.text, result.error


def _auto_orient_image(executable: Path, image):
    """Rotate a scanned page using Tesseract's script/orientation detector."""
    output, _error = _run_tesseract_image(executable, image, "osd", psm="0")
    match = re.search(r"Rotate:\s*(0|90|180|270)", output)
    if not match:
        return image
    clockwise_rotation = int(match.group(1))
    if clockwise_rotation == 0:
        return image
    # PIL uses counter-clockwise angles; Tesseract reports the clockwise
    # correction.  A reported 270 therefore becomes PIL +90.
    return image.rotate((360 - clockwise_rotation) % 360, expand=True)


def _normalise_scan_background(image):
    """Flatten a grey/uneven scan so Tesseract can see faint text.

    Phone/scanner PDFs sometimes contain a nearly full-page grey photograph.
    A direct OCR pass can return no text even though the receipt is readable to
    a person. Subtracting a blurred copy removes the slowly varying paper
    background while retaining small, dark strokes.
    """
    from PIL import ImageChops, ImageFilter, ImageOps

    grey = ImageOps.grayscale(image)
    background = None
    flattened = None
    enhanced = None
    try:
        radius = max(12, min(30, min(grey.size) // 80))
        background = grey.filter(ImageFilter.GaussianBlur(radius=radius))
        flattened = ImageChops.subtract(grey, background, scale=1, offset=255)
        enhanced = ImageOps.autocontrast(flattened, cutoff=1)
        return enhanced.convert("RGB")
    finally:
        grey.close()
        if background is not None:
            background.close()
        if flattened is not None:
            flattened.close()
        if enhanced is not None:
            enhanced.close()


def _didi_trip_table_totals(page_text: str) -> list[float]:
    """Return totals printed in Didi trip-table headers, not row noise."""
    normalised = _normalise_text(page_text)
    upper = normalised.upper()
    if not (re.search(r"DIDI\s*TRAVEL", upper) and "TRIP TABLE" in upper):
        return []
    matches = re.findall(
        r"(?:行程\s*)?合\s*计\s*(?:[:：])?\s*(?:¥|RMB|CNY)?\s*"
        r"([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)\s*元?",
        normalised,
        flags=re.IGNORECASE,
    )
    return [amount for match in matches if (amount := _plausible_amount(match)) is not None]


def _embedded_taxi_receipt_cities(page_text: str) -> set[str]:
    """Identify city names printed beside a taxi-receipt label on a page.

    A scanned pack can place an unrelated taxi receipt beside a Didi trip
    table. Only names in the receipt's local OCR context count here; the city
    printed in the Didi row is not evidence about the separate receipt.
    """
    normalised = _normalise_text(page_text)
    normalised = re.sub(r"BE[\(\[]JING", "BEIJING", normalised, flags=re.IGNORECASE)
    receipt_cities: set[str] = set()
    receipt_cities.update(
        _normalise_city(match.group(1).replace("_", " "))
        for match in re.finditer(
            r"IAV_EMBEDDED_TAXI_RECEIPT_CITY\s+([A-Za-z]+(?:_[A-Za-z]+)*)",
            normalised,
            flags=re.IGNORECASE,
        )
    )
    taxi_context = re.compile(
        r"出租车|出租汽车|出租|\bTAXI\b|\bCAB\b|SPECIAL\s+INVOICE",
        flags=re.IGNORECASE,
    )
    aliases = sorted(CITY_MAP.items(), key=lambda item: len(item[0]), reverse=True)
    for alias, canonical_city in aliases:
        flags = re.IGNORECASE if alias.isascii() else 0
        for match in re.finditer(re.escape(alias), normalised, flags=flags):
            context = normalised[max(0, match.start() - 55):match.end() + 75]
            if taxi_context.search(context):
                receipt_cities.add(_normalise_city(canonical_city))
    return receipt_cities


def _embedded_taxi_receipt_city(
    executable: Path,
    page_text: str,
    pdf_page,
) -> Optional[str]:
    """Read a local taxi-invoice heading for a city cross-check.

    Full-page OCR often reads the amount but misses the small issuer heading.
    This narrow red-channel pass reads the separate receipt area and only
    returns a city when one unambiguous mapped name is adjacent to taxi/invoice
    wording.
    """
    upper = page_text.upper()
    if not (re.search(r"DIDI\s*TRAVEL", upper) and "TRIP TABLE" in upper):
        return None
    try:
        from PIL import ImageOps

        detail_image = pdf_page.render(
            scale=600 / 72, prefer_bgrx=True,
        ).to_pil().convert("RGB")
    except Exception:
        return None
    crop = None
    red_channel = None
    contrasted = None
    try:
        width, height = detail_image.size
        crop = detail_image.crop((
            int(width * 0.22), int(height * 0.48),
            int(width * 0.57), int(height * 0.76),
        ))
        red_channel = crop.getchannel("R")
        contrasted = ImageOps.autocontrast(red_channel, cutoff=1)
        recognised: list[str] = []
        for psm in ("6", "11"):
            output, _error = _run_tesseract_image(
                executable, contrasted, "chi_sim+eng", psm=psm,
            )
            if output:
                recognised.append(output)
        cities = set()
        for output in recognised:
            cities.update(_embedded_taxi_receipt_cities(output))
        return next(iter(cities)) if len(cities) == 1 else None
    except Exception:
        return None
    finally:
        if contrasted is not None:
            contrasted.close()
        if red_channel is not None:
            red_channel.close()
        if crop is not None:
            crop.close()
        detail_image.close()


def _out_of_trip_taxi_receipt_city(page_text: str, trip_city: str) -> Optional[str]:
    """Return a receipt city only when one unambiguous city conflicts with the trip."""
    if not trip_city:
        return None
    receipt_cities = _embedded_taxi_receipt_cities(page_text)
    if len(receipt_cities) != 1:
        return None
    receipt_city = next(iter(receipt_cities))
    if receipt_city.casefold() == _normalise_city(trip_city).casefold():
        return None
    return receipt_city


def _didi_annotation_additions(page_text: str, table_totals: list[float]) -> list[float]:
    """Recover a reviewer-marked addition when OCR loses the plus operands.

    Scanned Didi pages are sometimes annotated with an equation such as
    ``invoice + taxi receipt = reviewed total``.  The handwritten operands can
    be unreadable, but the final total and the printed trip-table total are
    often clear.  Their positive difference is a data-derived supplemental
    receipt amount; it is accepted only when the annotation appears before
    the trip-table header and a taxi receipt is present on the page.
    """
    if not table_totals:
        return []
    normalised = _normalise_text(page_text)
    upper = normalised.upper()
    table_marker = re.search(r"DIDI\s*TRAVEL", upper)
    if not table_marker:
        return []
    annotation_area = normalised[:table_marker.start()]
    annotation_matches = re.findall(
        r"=\s*(?:¥|RMB|CNY)?\s*([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)",
        annotation_area,
        flags=re.IGNORECASE,
    )
    annotated_totals: list[float] = []
    for match in annotation_matches:
        compact = match.replace(",", "").replace(" ", "")
        candidates = []
        if (amount := _plausible_amount(compact)) is not None:
            candidates.append(amount)
        # Handwritten decimal points are frequently dropped by OCR (267.5 ->
        # 2675). Try decimal placements and retain the smallest value that is
        # still above the printed table total below; this is data-derived and
        # does not assume any particular invoice amount.
        if "." not in compact and 2 <= len(compact) <= 8:
            for position in range(1, len(compact)):
                variant = _plausible_amount(
                    compact[:position] + "." + compact[position:]
                )
                if variant is not None:
                    candidates.append(variant)
        annotated_totals.extend(candidates)
    if not annotated_totals or not (
        any(marker in normalised for marker in ("出租车", "机打发票"))
        or any(marker in normalised.upper() for marker in ("TAXI", "SPECIAL INVOICE"))
    ):
        return []
    additions: list[float] = []
    for table_total in table_totals:
        possible_differences = [
            round(annotated_total - table_total, 2)
            for annotated_total in annotated_totals
            if annotated_total > table_total + 0.01
        ]
        if not possible_differences:
            continue
        difference = min(possible_differences)
        if difference <= MAX_REASONABLE_EXPENSE and not any(
            abs(difference - existing) <= 0.01 for existing in additions
        ):
            additions.append(difference)
    return additions


def _embedded_taxi_receipt_amounts(
    executable: Path,
    image,
    page_text: str,
    pdf_page=None,
) -> list[float]:
    """Read side-by-side machine taxi receipts below a Didi trip table.

    These narrow blue-on-yellow receipts are often covered by a red invoice
    stamp. Full-page greyscale OCR can mistake footer serials for money. The
    red colour channel suppresses the stamp; fixed left/right value columns are
    then read at 2x scale and only standalone decimal values are considered.
    """
    upper = page_text.upper()
    if not (re.search(r"DIDI\s*TRAVEL", upper) and "TRIP TABLE" in upper):
        return []
    if not any(marker in page_text for marker in ("出租车", "出租汽车", "机打发")) and not (
        "TAXI" in upper or "SPECIAL INVOICE" in upper
    ):
        return []
    trip_table_totals = _didi_trip_table_totals(page_text)
    # A handwritten reviewed total can identify one otherwise unreadable
    # receipt. This path is independent of the fixed two-receipt crop layout.
    annotated_amounts = _didi_annotation_additions(page_text, trip_table_totals)
    repeated_receipt_fields = (
        page_text.count("公司电话"),
        page_text.count("监督电话"),
        page_text.count("卡号"),
    )
    has_repeated_receipts = not any(count < 2 for count in repeated_receipt_fields)

    from PIL import Image, ImageOps

    detail_image = None
    source_image = image
    if pdf_page is not None:
        # The amount digits are only a few pixels wide in a normal 300-DPI
        # full-page render. Re-render this exceptional page at 600 DPI instead
        # of trying to recover digits after an up-scale.
        try:
            detail_image = pdf_page.render(scale=600 / 72, prefer_bgrx=True).to_pil().convert("RGB")
            source_image = detail_image
        except Exception:
            detail_image = None
            source_image = image

    if detail_image is not None and trip_table_totals:
        # The reviewer annotation is often too small for the full-page OCR
        # layout, but remains readable in a high-resolution top-strip pass.
        annotation_crop = None
        annotation_image = None
        try:
            annotation_crop = detail_image.crop(
                (0, 0, detail_image.width, int(detail_image.height * 0.42))
            )
            annotation_image = ImageOps.autocontrast(
                annotation_crop.getchannel("R"), cutoff=1,
            )
            annotation_texts = []
            for psm in ("6", "11"):
                output, _error = _run_tesseract_image(
                    executable, annotation_image, "eng", psm=psm,
                )
                if output:
                    annotation_texts.append(output)
            if annotation_texts:
                annotated_amounts = _didi_annotation_additions(
                    page_text + "\n" + "\n".join(annotation_texts),
                    trip_table_totals,
                )
        finally:
            if annotation_image is not None:
                annotation_image.close()
            if annotation_crop is not None:
                annotation_crop.close()

    width, height = source_image.size
    crop_specs = (
        ((0.19, 0.70, 0.34, 0.88), -3),
        ((0.43, 0.66, 0.58, 0.75), -5),
    )
    amounts: list[float] = []
    for relative_box, angle in crop_specs:
        box = tuple(int(value * size) for value, size in zip(
            relative_box, (width, height, width, height),
        ))
        crop = source_image.crop(box)
        rotated = None
        enlarged = None
        red_channel = None
        contrasted = None
        binary = None
        try:
            rotated = crop.rotate(angle, expand=True, fillcolor="white")
            scale_factor = 1 if detail_image is not None else 2
            enlarged = rotated.resize(
                (rotated.width * scale_factor, rotated.height * scale_factor),
                resample=Image.Resampling.LANCZOS,
            )
            red_channel = enlarged.getchannel("R")
            contrasted = ImageOps.autocontrast(red_channel, cutoff=1)
            binary = contrasted.point(lambda pixel: 255 if pixel > 140 else 0).convert("RGB")
            output, _error = _run_tesseract_image(executable, binary, "eng", psm="6")
            candidates: list[float] = []
            for line in output.splitlines():
                compact = line.strip().replace(" ", "").replace(",", ".")
                match = re.fullmatch(r"[¥￥YF#]?([0-9]{1,4}\.[0-9]{1,2})", compact, re.IGNORECASE)
                if match:
                    amount = _plausible_amount(match.group(1))
                    if amount is not None:
                        candidates.append(amount)
            if candidates:
                amounts.append(max(candidates))
        finally:
            crop.close()
            for temporary in (rotated, enlarged, red_channel, contrasted, binary):
                if temporary is not None:
                    temporary.close()
    if detail_image is not None:
        detail_image.close()
    amounts = [
        amount for amount in amounts
        if not any(abs(amount - total) <= 0.01 for total in trip_table_totals)
    ]
    # This fallback is intentionally conservative: its target layout contains
    # two extra, side-by-side paper receipts. A single crop value is commonly
    # just the Didi trip total or a mileage/unit-price false positive.
    crop_amounts = amounts if has_repeated_receipts and len(amounts) >= 2 else []
    combined = annotated_amounts + crop_amounts
    result: list[float] = []
    for amount in combined:
        if not any(abs(amount - existing) <= 0.01 for existing in result):
            result.append(amount)
    return result


def _toll_invoice_amounts(
    executable: Path,
    page_text: str,
    pdf_page=None,
) -> list[float]:
    """Recover the payable amount from a sparse physical highway invoice.

    These receipts may OCR the visible ``131元`` as ``131z`` or merge nearby
    small figures into a misleading number at normal resolution. A sparse,
    high-resolution pass prioritises an amount followed by the yuan suffix;
    currency-marked multi-digit values are the fallback. No invoice amount or
    invoice number is used as a special case.
    """
    normalised = _normalise_text(page_text)
    if not (
        "发票代码" in normalised
        and "发票号码" in normalised
        and any(
            keyword in normalised.lower()
            for keyword in ("高速", "通行费", "过路费", "收费公路", "highway fee")
        )
    ):
        return []
    if pdf_page is None:
        return []

    from PIL import Image

    detail_image = None
    oriented_image = None
    try:
        detail_image = pdf_page.render(
            scale=600 / 72,
            prefer_bgrx=True,
        ).to_pil().convert("RGB")
        oriented_image = _auto_orient_image(executable, detail_image)
        languages = _ocr_languages(executable)
        for psm in ("11", "6", "12"):
            output, _error = _run_tesseract_image(
                executable, oriented_image, languages, psm=psm,
            )
            if not output:
                continue
            output = _normalise_text(output)
            suffix_matches = re.findall(
                r"([0-9]{1,7}(?:\.\s*[0-9]{1,2})?)\s*元\b",
                output,
                flags=re.IGNORECASE,
            )
            suffix_amounts = [
                amount for match in suffix_matches
                if (amount := _plausible_amount(match)) is not None
            ]
            if suffix_amounts:
                return [max(suffix_amounts)]

            currency_matches = re.findall(
                r"(?:¥|￥|RMB|CNY|(?<![A-Z])Y)\s*:?\s*"
                r"([0-9]{2,}[0-9,]*(?:\.\s*[0-9]{1,2})?)",
                output.upper(),
            )
            currency_amounts = [
                amount for match in currency_matches
                if (amount := _plausible_amount(match)) is not None
            ]
            if currency_amounts:
                return [max(currency_amounts)]
    finally:
        if oriented_image is not None and oriented_image is not detail_image:
            oriented_image.close()
        if detail_image is not None:
            detail_image.close()
    return []


def _ocr_quality(text: str) -> tuple[int, int, int]:
    """Rank OCR layouts by travel-relevant signals, not raw character count."""
    normalised = _normalise_text(text)
    currency_decimal_amounts = re.findall(
        r"(?:¥|￥|RMB|CNY|yuan|元)\s*[0-9][0-9,]*\.\s*[0-9]{1,2}",
        normalised,
        flags=re.IGNORECASE,
    )
    currency_amounts = re.findall(
        r"(?:¥|￥|RMB|CNY|yuan|元)\s*[0-9][0-9,]*(?:\.\s*[0-9]{1,2})?",
        normalised,
        flags=re.IGNORECASE,
    )
    decimal_amounts = re.findall(r"(?<!\d)\d{1,7}\.\s*\d{1,2}(?!\d)", normalised)
    date_signals = re.findall(r"20\d{2}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{1,2}", normalised)
    keywords = sum(
        normalised.lower().count(keyword.lower())
        for keyword in SPECIAL_KEYWORDS + HOTEL_KEYWORDS + TICKET_KEYWORDS + TAXI_KEYWORDS
    )
    invoice_cues = sum(
        normalised.lower().count(keyword)
        for keyword in (
            "发票", "invoice", "receipt", "价税合计", "grand total",
            "小写", "出租汽车", "铁路电子客票", "酒店管理",
        )
    )
    invoice_records, supporting_only = _ocr_page_invoice_records(normalised)
    payable_candidates = len(invoice_records) if not supporting_only else 0
    folio_points = 0
    folio = _hotel_folio_evidence(normalised, _document_year(normalised, None))
    if folio is not None:
        # Prefer OCR candidates that recover the dates needed to calculate
        # nights over longer but poorly oriented folio transcripts.
        if folio.start is not None and folio.end is not None and folio.end > folio.start:
            folio_points += 100
        if folio.nights is not None:
            folio_points += 25
        if folio.total is not None:
            folio_points += 15
    return (
        len(currency_decimal_amounts) * 35 + len(currency_amounts) * 8
        + len(decimal_amounts) * 5 + len(date_signals) * 8 + keywords * 12
        + invoice_cues * 18 + payable_candidates * 40 + folio_points,
        len(normalised),
        sum(character.isdigit() for character in normalised),
    )


def _ocr_page_needs_orientation_retry(text: str) -> bool:
    """Identify nonempty OCR text that still lacks invoice-level evidence."""
    normalised = _normalise_text(text)
    has_hotel_stay_signal = bool(re.search(
        r"\u5165\u4f4f|\u79bb\u5e97|\u9000\u623f|\u6c34\u5355|\u7ed3\u7b97\u5355|"
        r"check[\s-]*in|check[\s-]*out|arrival|departure|folio",
        normalised,
        flags=re.IGNORECASE,
    ))
    if _is_hotel_folio(normalised) or has_hotel_stay_signal:
        folio = _hotel_folio_evidence(normalised, _document_year(normalised, None))
        # Folios are supporting pages, but their dates are required business
        # evidence. Supporting-only classification must not suppress a retry
        # when a check-in/check-out date is missing or unreadable. The broader
        # stay signal also catches sideways OCR that retained one date label
        # and therefore did not yet qualify as a hotel folio.
        return not bool(
            folio
            and folio.start is not None
            and folio.end is not None
            and folio.end > folio.start
            and folio.nights is not None
        )
    records, supporting_only = _ocr_page_invoice_records(normalised)
    if supporting_only:
        return False
    if any(category is None for _amount, category in records):
        return True
    if records:
        return False
    has_document_signal = bool(re.search(
        r"发票|invoice|receipt|出租|打车|的士|网约车|酒店|住宿|火车|客票|机票|"
        r"高速|通行费|价税合计|小\s*写|(?:¥|￥|RMB|CNY)\s*\d",
        normalised,
        flags=re.IGNORECASE,
    ))
    has_supporting_signal = bool(re.search(
        r"行程单|TRIP\s*TABLE|订单详情|入住水单|酒店水单|check[ -]*in|check[ -]*out",
        normalised,
        flags=re.IGNORECASE,
    ))
    # Retry short/noisy or table-like pages, but leave text-rich itinerary and
    # policy pages on the regular route instead of spending extra OCR passes.
    table_date_count = len(_date_tokens(normalised, _document_year(normalised, None)))
    table_amount_count = len(re.findall(
        r"(?<!\d)\d{1,7}\.\s*\d{1,2}(?!\d)", normalised,
    ))
    has_unclassified_table_signal = table_date_count >= 2 or table_amount_count >= 3
    has_no_travel_evidence = _ocr_quality(normalised)[0] == 0
    return not has_supporting_signal and (
        has_document_signal or len(normalised) < 500
        or has_unclassified_table_signal or has_no_travel_evidence
    )


def _looks_like_taxi_invoice(text: str) -> bool:
    """Recognise clean and partially OCR'd taxi-invoice titles."""
    normalised = _normalise_text(text)
    compact_upper = re.sub(r"\s+", "", normalised.upper())
    return bool(
        re.search(
            r"出租车|出租汽车|出租发票|打车|网约车|的士|TAXI|RIDE\s*HAIL",
            normalised,
            flags=re.IGNORECASE,
        )
        or re.search(r"TAX.{0,12}SPECI.{0,30}NVOI", compact_upper)
    )


def _handwritten_taxi_note_amount_candidates(executable: Path, pdf_page) -> list[float]:
    """Read possible handwritten amounts from the upper-right note area.

    A number ending in a percent-like glyph is returned only as a review
    candidate: OCR often mistakes the Chinese yuan glyph for ``%`` in pen
    annotations. Callers must keep this separate from verified invoice totals.
    """
    from PIL import Image, ImageEnhance, ImageOps

    source_image = None
    note_crop = None
    candidates: list[float] = []
    try:
        source_image = pdf_page.render(
            scale=600 / 72,
            prefer_bgrx=True,
        ).to_pil().convert("RGB")
        note_crop = source_image.crop((
            int(source_image.width * 0.80),
            0,
            source_image.width,
            int(source_image.height * 0.23),
        ))
        for angle in (90, 270):
            rotated = note_crop.rotate(angle, expand=True, fillcolor="white")
            grey = ImageOps.grayscale(rotated)
            contrasted = ImageOps.autocontrast(grey, cutoff=1)
            enlarged = contrasted.resize(
                (contrasted.width * 3, contrasted.height * 3),
                resample=Image.Resampling.LANCZOS,
            )
            try:
                output, _error = _run_tesseract_image(
                    executable, enlarged, _ocr_languages(executable), psm="6",
                )
                normalised = _normalise_text(output)
                patterns = (
                    r"(?:¥|￥|RMB|CNY)\s*([0-9]{1,4}(?:\.\s*[0-9]{1,2})?)",
                    r"([0-9]{1,4}(?:\.\s*[0-9]{1,2})?)\s*(?:元|YUAN)",
                    r"([0-9]{2,3})\s*%",
                )
                for pattern in patterns:
                    for match in re.findall(pattern, normalised, flags=re.IGNORECASE):
                        amount = _plausible_amount(match)
                        if amount is not None and not any(
                            abs(amount - existing) <= 0.01 for existing in candidates
                        ):
                            candidates.append(amount)
                # Faint pen strokes can make the yuan glyph look like a slash.
                # Keep a single split two-digit token as an unverified review
                # candidate; never treat it as an invoice total automatically.
                for left, right in re.findall(r"(?<!\d)(\d)\s*/\s*(\d)(?!\d)", normalised):
                    amount = _plausible_amount(f"{left}{right}")
                    if amount is not None and not any(
                        abs(amount - existing) <= 0.01 for existing in candidates
                    ):
                        candidates.append(amount)
            finally:
                enlarged.close()
                contrasted.close()
                grey.close()
                rotated.close()
    finally:
        if note_crop is not None:
            note_crop.close()
        if source_image is not None:
            source_image.close()
    return candidates


def _ocr_pdf_pages(path: str, page_indexes: list[int]) -> tuple[dict[int, str], list[str]]:
    """Serialize PDFium calls; the native library is not safe across threads."""
    with _PDFIUM_LOCK:
        return _ocr_pdf_pages_locked(path, page_indexes)


def _ocr_pdf_pages_locked(path: str, page_indexes: list[int]) -> tuple[dict[int, str], list[str]]:
    """Render selected PDF pages and OCR them locally.

    Only pages with no useful text are rendered.  This keeps normal text PDFs
    fast and also handles mixed PDFs containing both native and scanned pages.
    """
    if not page_indexes:
        return {}, []
    executable = _ocr_executable()
    if executable is None:
        return {}, ["未找到 Tesseract OCR，请安装 Tesseract 或检查 config.py 的 TESSERACT_EXE"]

    try:
        import config
        dpi = int(getattr(config, "OCR_DPI", 300) or 300)
    except (ImportError, TypeError, ValueError):  # pragma: no cover - config is present in production
        dpi = 300
    dpi = max(150, min(dpi, 400))
    languages = _ocr_languages(executable)

    try:
        import pypdfium2 as pdfium
    except ImportError as exc:  # pragma: no cover - environment-specific
        return {}, [f"未安装 pypdfium2，无法把 PDF 页面转换为图片: {exc}"]

    results: dict[int, str] = {}
    errors: list[str] = []
    document = None
    try:
        document = pdfium.PdfDocument(str(path))
        page_count = len(document)
        for page_index in page_indexes:
            if page_index < 0 or page_index >= page_count:
                continue
            page = None
            image = None
            ocr_image = None
            normalised_image = None
            try:
                page = document.get_page(page_index)
                image = page.render(scale=dpi / 72, prefer_bgrx=True).to_pil().convert("RGB")
                ocr_image = _auto_orient_image(executable, image)
                candidates: list[tuple[tuple[int, int, int], str]] = []
                page_errors: list[str] = []
                # PSM 6 works well for form-like pages; PSM 12 is better for
                # sparse receipts and often preserves currency amounts.
                for psm in ("6", "12"):
                    output, error = _run_tesseract_image(
                        executable, ocr_image, languages, psm=psm,
                    )
                    if output:
                        candidates.append((_ocr_quality(output), output))
                    if error:
                        page_errors.append(error)
                # Low-contrast hotel folios and photographed receipts can be
                # perfectly legible while a direct Tesseract pass returns an
                # empty string. Retry only those failed pages after flattening
                # the grey background, so normal pages do not pay this cost.
                # Low-contrast pages with no usable first-pass text get the
                # existing background-normalisation retry.  Non-empty pages
                # remain on the proven production route until a field-level
                # Travel regression corpus justifies a broader retry gate.
                best_first_pass = max(candidates, key=lambda item: item[0])[1] if candidates else ""
                # A single OSD rotation can make a sideways receipt unreadable
                # when it shares a page with an upright note. Retry a few
                # alternate directions only while invoice-level evidence is
                # still incomplete.
                if not candidates or _ocr_page_needs_orientation_retry(best_first_pass):
                    normalised_image = _normalise_scan_background(ocr_image)
                    retry_images = [normalised_image]
                    rotated_images = []
                    for angle in (90, 180, 270):
                        rotated = image.rotate(angle, expand=True, fillcolor="white")
                        rotated_images.append(rotated)
                        retry_images.append(rotated)
                        # The folio pages that reach this retry path are often
                        # photographed on uneven grey paper. Flatten each
                        # alternate orientation too; flattening only the OSD
                        # orientation leaves the correctly rotated page noisy.
                        normalised_rotated = _normalise_scan_background(rotated)
                        rotated_images.append(normalised_rotated)
                        retry_images.append(normalised_rotated)
                    try:
                        for retry_image in retry_images:
                            # PSM 6 keeps the line layout of upright hotel
                            # folios; without it the background-flattened
                            # candidate was never produced for low-contrast
                            # water bills. Alternate orientations stay on the
                            # sparse-text modes to bound OCR cost.
                            retry_psms = (
                                ("6", "11", "12")
                                if retry_image is normalised_image
                                else ("11", "12")
                            )
                            for psm in retry_psms:
                                output, error = _run_tesseract_image(
                                    executable, retry_image, languages, psm=psm,
                                )
                                if output:
                                    candidates.append((_ocr_quality(output), output))
                                if error:
                                    page_errors.append(error)
                    finally:
                        for rotated in rotated_images:
                            rotated.close()
                if candidates:
                    best_text = max(candidates, key=lambda item: item[0])[1]
                    toll_amounts: list[float] = []
                    if not _ocr_gross_total_matches(best_text):
                        try:
                            toll_amounts = _toll_invoice_amounts(
                                executable, best_text, pdf_page=page,
                            )
                        except Exception:
                            # The high-resolution toll pass is supplemental;
                            # it must never discard the full-page OCR result.
                            toll_amounts = []
                    try:
                        embedded_amounts = _embedded_taxi_receipt_amounts(
                            executable, ocr_image, best_text, pdf_page=page,
                        )
                    except Exception:
                        # Supplemental receipt OCR must never discard a valid
                        # full-page OCR result.
                        embedded_amounts = []
                    embedded_receipt_city = None
                    if embedded_amounts:
                        try:
                            embedded_receipt_city = _embedded_taxi_receipt_city(
                                executable, best_text, page,
                            )
                        except Exception:
                            # City evidence is a scope check, not a prerequisite
                            # for retaining the page's base OCR result.
                            embedded_receipt_city = None
                    supplemental_lines = [
                        *(f"IAV_EMBEDDED_TAXI_RECEIPT ¥{amount:.2f}"
                          for amount in embedded_amounts),
                        *([f"IAV_EMBEDDED_TAXI_RECEIPT_CITY {embedded_receipt_city.replace(' ', '_')}" ]
                          if embedded_receipt_city else []),
                        *(f"IAV_TOLL_INVOICE_AMOUNT ¥{amount:.2f}"
                          for amount in toll_amounts),
                    ]
                    page_invoice_records, page_is_supporting = _ocr_page_invoice_records(best_text)
                    if (
                        not page_invoice_records
                        and not page_is_supporting
                        and _looks_like_taxi_invoice(best_text)
                    ):
                        for amount in _handwritten_taxi_note_amount_candidates(executable, page):
                            supplemental_lines.append(
                                f"IAV_TAXI_NOTE_AMOUNT_REVIEW ¥{amount:.2f}"
                            )
                    if supplemental_lines:
                        supplemental = "\n".join(supplemental_lines)
                        if embedded_amounts:
                            table_totals = _didi_trip_table_totals(best_text)
                            annotation_additions = _didi_annotation_additions(
                                best_text, table_totals,
                            )
                            if annotation_additions and table_totals:
                                reviewed_total = round(
                                    table_totals[-1] + sum(annotation_additions), 2,
                                )
                                supplemental += (
                                    f"\nIAV_DIDI_REVIEWED_TOTAL ¥{reviewed_total:.2f}"
                                )
                        best_text = f"{best_text}\n{supplemental}"
                    results[page_index] = best_text
                elif page_errors:
                    errors.append(f"第 {page_index + 1} 页: {page_errors[-1]}")
                else:
                    errors.append(f"第 {page_index + 1} 页: Tesseract 未返回文字")
            except Exception as exc:  # OCR failure should remain a review warning
                errors.append(f"第 {page_index + 1} 页: {type(exc).__name__}: {exc}")
            finally:
                if normalised_image is not None:
                    normalised_image.close()
                if ocr_image is not None and ocr_image is not image:
                    ocr_image.close()
                if image is not None:
                    image.close()
                if page is not None and hasattr(page, "close"):
                    page.close()
    except Exception as exc:
        errors.append(f"PDF OCR 初始化失败: {type(exc).__name__}: {exc}")
    finally:
        if document is not None and hasattr(document, "close"):
            document.close()
    return results, errors


def _travel_reocr_cache_config(executable: Path, languages: str, dpi: int) -> dict:
    """Return every OCR/render input that can change a cached page result."""
    try:
        stat = executable.stat()
        executable_fingerprint = {
            "path": str(executable.resolve()).casefold(),
            "size": stat.st_size,
            "mtime_ns": stat.st_mtime_ns,
        }
    except OSError:
        executable_fingerprint = {"path": str(executable).casefold()}
    try:
        result = subprocess.run(
            [str(executable), "--version"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=5,
            check=False,
            creationflags=(getattr(subprocess, "CREATE_NO_WINDOW", 0)
                           if os.name == "nt" else 0),
        )
        tesseract_version = (result.stdout or result.stderr or "").splitlines()[0][:160]
    except (OSError, subprocess.SubprocessError, IndexError):
        tesseract_version = "unknown"

    traineddata = []
    tessdata_prefix = os.environ.get("TESSDATA_PREFIX", "")
    data_roots = [
        Path(tessdata_prefix) if tessdata_prefix else None,
        executable.parent / "tessdata",
        executable.parent / "share" / "tessdata",
        executable.parent.parent / "share" / "tesseract-ocr" / "5" / "tessdata",
    ]
    for language in languages.split("+"):
        model_path = None
        for root in data_roots:
            if root is None:
                continue
            for candidate in (root / f"{language}.traineddata", root / "tessdata" / f"{language}.traineddata"):
                if candidate.is_file():
                    model_path = candidate
                    break
            if model_path is not None:
                break
        if model_path is None:
            traineddata.append({"language": language, "status": "unresolved"})
            continue
        try:
            stat = model_path.stat()
            traineddata.append({
                "language": language,
                "path": str(model_path.resolve()).casefold(),
                "size": stat.st_size,
                "mtime_ns": stat.st_mtime_ns,
            })
        except OSError:
            traineddata.append({"language": language, "status": "unavailable"})

    def installed_version(name: str) -> str:
        try:
            return package_version(name)
        except PackageNotFoundError:
            return "unknown"

    return {
        "cache_schema": TRAVEL_REOCR_CACHE_SCHEMA,
        "candidate_parser": "travel-invoice-fields-v3-lower-page-region",
        "renderer": "pypdfium2",
        "renderer_version": installed_version("pypdfium2"),
        "image_library": "Pillow",
        "image_library_version": installed_version("Pillow"),
        "tesseract": executable_fingerprint,
        "tesseract_version": tesseract_version,
        "traineddata": traineddata,
        "tessdata_prefix": tessdata_prefix,
        "languages": languages,
        "dpi": dpi,
        "page_modes": ["6", "11", "12"],
        "output_mode": "tsv",
        "orientation": "tesseract-osd-hotel-folio-retry-v2",
    }


def _travel_reocr_cache_path(source_sha256: str, page: int, config: dict) -> Path:
    identity = {
        "source_sha256": source_sha256,
        "page": page,
        "config": config,
    }
    encoded = json.dumps(identity, sort_keys=True, separators=(",", ":")).encode("utf-8")
    cache_key = hashlib.sha256(encoded).hexdigest()
    return TRAVEL_REOCR_CACHE_DIR / f"{cache_key}.json"


def _prune_travel_reocr_cache(now: Optional[float] = None) -> int:
    """Bound cache growth by age and entry count; return removed entry count."""
    now = time.time() if now is None else now
    removed = 0
    with _TRAVEL_REOCR_CACHE_LOCK:
        try:
            TRAVEL_REOCR_CACHE_DIR.mkdir(parents=True, exist_ok=True)
            entries = []
            for path in TRAVEL_REOCR_CACHE_DIR.glob("*.json"):
                try:
                    stat = path.stat()
                except OSError:
                    continue
                if now - stat.st_mtime > TRAVEL_REOCR_CACHE_MAX_AGE_SECONDS:
                    try:
                        path.unlink()
                        removed += 1
                    except OSError:
                        pass
                    continue
                entries.append((stat.st_mtime, path, stat.st_size))
            for path in TRAVEL_REOCR_CACHE_DIR.glob("*.tmp"):
                try:
                    if now - path.stat().st_mtime > 24 * 60 * 60:
                        path.unlink()
                        removed += 1
                except OSError:
                    pass
            entries.sort(reverse=True, key=lambda item: item[0])
            for _modified, path, _size in entries[TRAVEL_REOCR_CACHE_MAX_ENTRIES:]:
                try:
                    path.unlink()
                    removed += 1
                except OSError:
                    pass
            retained = entries[:TRAVEL_REOCR_CACHE_MAX_ENTRIES]
            retained_bytes = sum(item[2] for item in retained)
            for _modified, path, size in reversed(retained):
                if retained_bytes <= TRAVEL_REOCR_CACHE_MAX_BYTES:
                    break
                try:
                    path.unlink()
                    retained_bytes -= size
                    removed += 1
                except OSError:
                    continue
        except OSError:
            return removed
    return removed


def _load_travel_reocr_cache(
    path: Path,
    source_sha256: str,
    page_number: int,
    config: dict,
) -> tuple[Optional[dict], Optional[str]]:
    """Read a validated cached page; corrupt entries are discarded for retry."""
    with _TRAVEL_REOCR_CACHE_LOCK:
        try:
            with path.open("r", encoding="utf-8") as cache_handle:
                entry = json.load(cache_handle)
        except FileNotFoundError:
            return None, None
        except (OSError, UnicodeError, json.JSONDecodeError):
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass
            return None, "缓存文件损坏，已清除并重新识别。"

        if (
            not isinstance(entry, dict)
            or entry.get("schema") != TRAVEL_REOCR_CACHE_SCHEMA
            or entry.get("source_sha256") != source_sha256
            or entry.get("page") != page_number
            or entry.get("config") != config
            or not isinstance(entry.get("page_result"), dict)
            or entry["page_result"].get("page") != page_number
            or entry["page_result"].get("error")
        ):
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass
            return None, "缓存记录与当前文件或识别配置不一致，已清除并重新识别。"

        try:
            os.utime(path, None)
        except OSError:
            pass
        return copy.deepcopy(entry["page_result"]), None


def _store_travel_reocr_cache(
    path: Path,
    source_sha256: str,
    page_number: int,
    config: dict,
    page_result: dict,
) -> Optional[str]:
    """Atomically persist successful OCR evidence without storing page images."""
    try:
        if not any(
            isinstance(candidate, dict) and str(candidate.get("text", "")).strip()
            for candidate in page_result.get("candidates", [])
        ):
            return ""
        entry = {
            "schema": TRAVEL_REOCR_CACHE_SCHEMA,
            "source_sha256": source_sha256,
            "page": page_number,
            "config": config,
            "cached_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "page_result": page_result,
        }
        with _TRAVEL_REOCR_CACHE_LOCK:
            TRAVEL_REOCR_CACHE_DIR.mkdir(parents=True, exist_ok=True)
            temporary_path = None
            try:
                with tempfile.NamedTemporaryFile(
                    mode="w",
                    encoding="utf-8",
                    newline="\n",
                    dir=TRAVEL_REOCR_CACHE_DIR,
                    prefix=f".{path.stem}-",
                    suffix=".tmp",
                    delete=False,
                ) as cache_handle:
                    temporary_path = Path(cache_handle.name)
                    json.dump(entry, cache_handle, ensure_ascii=False, separators=(",", ":"))
                    cache_handle.flush()
                    os.fsync(cache_handle.fileno())
                os.replace(temporary_path, path)
            finally:
                if temporary_path is not None:
                    try:
                        temporary_path.unlink(missing_ok=True)
                    except OSError:
                        pass
        _prune_travel_reocr_cache()
        if not path.exists():
            return "识别结果超过缓存容量上限，未持久化；当前结果仍可查看。"
        return None
    except (OSError, TypeError, ValueError) as exc:
        return f"OCR 缓存未能写入；本次结果仍可查看，之后会重新识别（{type(exc).__name__}）。"


def reocr_pdf_pages_for_review(
    path: str,
    page_numbers: list[int],
    dpi: int = 600,
) -> dict:
    """Serialize review OCR with all other PDFium calls in this process."""
    with _PDFIUM_LOCK:
        return _reocr_pdf_pages_for_review_locked(path, page_numbers, dpi)


def _reocr_pdf_pages_for_review_locked(
    path: str,
    page_numbers: list[int],
    dpi: int = 600,
) -> dict:
    """Re-read a small set of PDF pages and return review-only OCR evidence.

    This path is deliberately additive: it does not modify the parsed Trip or
    create billable expenses. Successful page evidence is cached by source
    SHA-256, page, and OCR/render configuration. Cached values contain text and
    geometry only; rendered page images are never persisted.
    """
    pages = list(dict.fromkeys(int(value) for value in page_numbers))
    if not pages or len(pages) > 3:
        raise ValueError("一次局部重识别请选择 1 到 3 页。")
    if any(value < 1 for value in pages):
        raise ValueError("页码必须从 1 开始。")
    dpi = max(300, min(int(dpi), 600))

    executable = _ocr_executable()
    if executable is None:
        raise RuntimeError("未找到 Tesseract OCR，请检查 config.py 中的 TESSERACT_EXE。")
    languages = _ocr_languages(executable)
    router = build_tesseract_router(executable, languages=languages, timeout=90)
    try:
        import pypdfium2 as pdfium
    except ImportError as exc:  # pragma: no cover - environment-specific
        raise RuntimeError(f"未安装 pypdfium2，无法定向读取 PDF 页面: {exc}") from exc

    digest = hashlib.sha256()
    with open(path, "rb") as source_handle:
        for chunk in iter(lambda: source_handle.read(1024 * 1024), b""):
            digest.update(chunk)
    source_sha256 = digest.hexdigest()
    cache_config = _travel_reocr_cache_config(executable, languages, dpi)
    cache_config_fingerprint = hashlib.sha256(
        json.dumps(cache_config, sort_keys=True, separators=(",", ":")).encode("utf-8")
    ).hexdigest()
    pruned_cache_entries = _prune_travel_reocr_cache()

    started = time.perf_counter()
    document = None
    page_results = []
    cache_hits = 0
    cache_misses = 0
    cache_stored = 0
    cache_failures = 0
    cache_warnings = []
    try:
        document = pdfium.PdfDocument(str(path))
        page_count = len(document)
        if any(value > page_count for value in pages):
            raise ValueError(f"页码超出 PDF 范围；文件共 {page_count} 页。")

        for page_number in pages:
            page_started = time.perf_counter()
            cache_path = _travel_reocr_cache_path(source_sha256, page_number, cache_config)
            cache_started = time.perf_counter()
            cached_page, cache_warning = _load_travel_reocr_cache(
                cache_path, source_sha256, page_number, cache_config,
            )
            cache_lookup_ms = round((time.perf_counter() - cache_started) * 1000, 1)
            if cache_warning:
                cache_warnings.append(f"第 {page_number} 页：{cache_warning}")
            if cached_page is not None:
                cached_page["ocr_elapsed_ms"] = cached_page.get("elapsed_ms")
                cached_page["elapsed_ms"] = cache_lookup_ms
                cached_page["cache_status"] = "HIT"
                cached_page["cache_elapsed_ms"] = cache_lookup_ms
                page_results.append(cached_page)
                cache_hits += 1
                continue
            cache_misses += 1
            page = None
            image = None
            oriented_image = None
            try:
                page = document.get_page(page_number - 1)
                image = page.render(
                    scale=dpi / 72,
                    prefer_bgrx=True,
                ).to_pil().convert("RGB")
                oriented_image = _auto_orient_image(executable, image)
                width, height = oriented_image.size
                candidates = []
                warnings = []

                def build_review_candidate(result, method: str, region_bbox=None) -> dict:
                    region_left, region_top = (region_bbox[:2] if region_bbox else (0, 0))
                    candidate_warnings = list(result.warnings)
                    warnings.extend(candidate_warnings)
                    line_groups: dict[tuple[int, ...], list[str]] = {}
                    positioned_blocks = []
                    for block in result.blocks:
                        line_id = tuple(block.line_id or ())
                        line_groups.setdefault(line_id, []).append(block.text)
                        left, top, right, bottom = block.bbox or (0, 0, 0, 0)
                        left += region_left
                        right += region_left
                        top += region_top
                        bottom += region_top
                        positioned_blocks.append({
                            "text": block.text,
                            "confidence": round(block.confidence / 100, 4)
                            if block.confidence is not None else None,
                            "bbox_pixels": [left, top, right, bottom],
                            "bbox_normalized": [
                                round(left / width, 5),
                                round(top / height, 5),
                                round(right / width, 5),
                                round(bottom / height, 5),
                            ],
                        })
                    text = "\n".join(
                        " ".join(words)
                        for words in line_groups.values()
                    ).strip()
                    quality = _ocr_quality(text)
                    invoice_candidates = []
                    if text:
                        try:
                            page_records, _supporting = _ocr_page_invoice_records(text)
                            if not page_records:
                                page_records = [(None, _ocr_page_category(text, None))]
                            invoice_blocks = [
                                {
                                    "text": block["text"],
                                    "bbox": block["bbox_pixels"],
                                    "confidence": block["confidence"],
                                    "page": page_number,
                                }
                                for block in positioned_blocks
                            ]
                            for invoice_index, (amount, category) in enumerate(page_records, 1):
                                fragment = (
                                    _invoice_page_fragment(text, amount, invoice_index - 1)
                                    if amount is not None else text
                                )
                                invoice = extract_travel_invoice(
                                    fragment,
                                    source_file=path,
                                    page=page_number,
                                    category=category or "",
                                    total_amount=amount,
                                    default_year=_document_year(text, None),
                                    blocks=invoice_blocks,
                                    ocr_engine="tesseract",
                                    ocr_confidence=result.confidence,
                                )
                                fields = []
                                for field_name in (
                                    "total_amount", "net_amount", "tax_amount", "currency",
                                    "expense_type", "invoice_date", "travel_date_from",
                                    "travel_date_to", "vendor", "invoice_number",
                                ):
                                    value = getattr(invoice, field_name)
                                    if isinstance(value, (date, datetime)):
                                        value = value.isoformat()
                                    evidence = invoice.evidence.get(field_name, {})
                                    bbox = evidence.get("bbox")
                                    normalised_bbox = (
                                        [
                                            round(bbox[0] / width, 5),
                                            round(bbox[1] / height, 5),
                                            round(bbox[2] / width, 5),
                                            round(bbox[3] / height, 5),
                                        ]
                                        if bbox and len(bbox) == 4 else None
                                    )
                                    if value not in (None, "") or evidence:
                                        fields.append({
                                            "field": field_name,
                                            "value": value,
                                            "source_text": evidence.get("raw_text", ""),
                                            "bbox_pixels": bbox,
                                            "bbox_normalized": normalised_bbox,
                                            "ocr_confidence": evidence.get("ocr_confidence"),
                                            "validation_status": evidence.get("validation_status", "CANDIDATE"),
                                            "travel_date_source": (
                                                invoice.travel_date_source
                                                if field_name.startswith("travel_date_") else None
                                            ),
                                        })
                                invoice_candidates.append({
                                    "candidate_index": invoice_index,
                                    "category": category,
                                    "total_amount": amount,
                                    "document_type": invoice.document_type,
                                    "validation": invoice.validation,
                                    "fields": fields,
                                })
                        except Exception as exc:
                            warning_text = (
                                f"字段候选关联未完成（{method}）：{type(exc).__name__}: {exc}"
                            )
                            warnings.append(warning_text)
                            candidate_warnings.append(warning_text)
                    candidate = {
                        "method": method,
                        "text": text,
                        "quality": list(quality),
                        "confidence": result.confidence,
                        "elapsed_ms": round(result.elapsed_ms or 0, 1),
                        "block_count": len(positioned_blocks),
                        "blocks_truncated": len(positioned_blocks) > 300,
                        "blocks": positioned_blocks[:300],
                        "invoice_candidates": invoice_candidates,
                        "warnings": candidate_warnings,
                    }
                    if region_bbox:
                        region_right, region_bottom = region_bbox[2:]
                        candidate.update({
                            "region_name": "lower_page_receipts",
                            "region_bbox_pixels": list(region_bbox),
                            "region_bbox_normalized": [
                                round(region_bbox[0] / width, 5),
                                round(region_bbox[1] / height, 5),
                                round(region_right / width, 5),
                                round(region_bottom / height, 5),
                            ],
                            "review_only": True,
                        })
                    return candidate

                for psm in ("6", "11", "12"):
                    result = router.recognize(
                        oriented_image,
                        languages=languages,
                        psm=psm,
                        page=page_number,
                        output_mode="tsv",
                    )
                    candidates.append(build_review_candidate(
                        result,
                        f"{dpi}dpi_psm{psm}_tsv",
                    ))

                # A lower-page pass helps separate a photographed taxi receipt
                # from the Didi itinerary printed above it. Trigger only when
                # full-page OCR shows both the trip sheet and receipt cues.
                all_page_text = "\n".join(candidate["text"] for candidate in candidates)
                compact_page_text = re.sub(r"\s+", "", all_page_text).upper()
                has_didi_sheet = bool(re.search(
                    r"DIDITRAVEL|TRIPTABLE|滴滴.*(?:行程|出行)",
                    compact_page_text,
                ))
                has_receipt_cues = bool(re.search(
                    r"TAXI|出租|国家税务总局|发票专用章|国家税务局",
                    compact_page_text,
                ))
                if has_didi_sheet and has_receipt_cues:
                    crop_top = int(height * 0.40)
                    receipt_region = oriented_image.crop((0, crop_top, width, height))
                    try:
                        for psm in ("6", "11", "12"):
                            result = router.recognize(
                                receipt_region,
                                languages=languages,
                                psm=psm,
                                page=page_number,
                                output_mode="tsv",
                            )
                            candidates.append(build_review_candidate(
                                result,
                                f"{dpi}dpi_lower_page_psm{psm}_tsv",
                                (0, crop_top, width, height),
                            ))
                    finally:
                        receipt_region.close()
                best = max(
                    (candidate for candidate in candidates if not candidate.get("region_name")),
                    key=lambda item: (
                        item["quality"][0], item["quality"][1],
                        item["confidence"] if item["confidence"] is not None else -1,
                    ),
                ) if candidates else None
                page_result = {
                    "page": page_number,
                    "render_size": [width, height],
                    "coordinate_space": "auto_oriented_rendered_page_pixels",
                    "elapsed_ms": round((time.perf_counter() - page_started) * 1000, 1),
                    "best_method": best["method"] if best else None,
                    "best_text": best["text"] if best else "",
                    "candidates": candidates,
                    "warnings": list(dict.fromkeys(warnings)),
                }
                cache_warning = _store_travel_reocr_cache(
                    cache_path, source_sha256, page_number, cache_config, page_result,
                )
                if cache_warning is None:
                    page_result["cache_status"] = "MISS_STORED"
                    cache_stored += 1
                elif cache_warning:
                    page_result["cache_status"] = "MISS_WRITE_FAILED"
                    page_result["cache_warning"] = cache_warning
                    cache_warnings.append(f"第 {page_number} 页：{cache_warning}")
                else:
                    page_result["cache_status"] = "MISS_NOT_CACHEABLE"
                page_results.append(page_result)
            except Exception as exc:
                page_results.append({
                    "page": page_number,
                    "elapsed_ms": round((time.perf_counter() - page_started) * 1000, 1),
                    "error": f"{type(exc).__name__}: {exc}",
                    "cache_status": "MISS_FAILED_NOT_CACHED",
                })
                cache_failures += 1
            finally:
                if oriented_image is not None and oriented_image is not image:
                    oriented_image.close()
                if image is not None:
                    image.close()
                if page is not None and hasattr(page, "close"):
                    page.close()
    finally:
        if document is not None and hasattr(document, "close"):
            document.close()

    return {
        "source_sha256": source_sha256,
        "engine": "tesseract",
        "languages": languages,
        "dpi": dpi,
        "page_count": page_count,
        "requested_pages": pages,
        "elapsed_ms": round((time.perf_counter() - started) * 1000, 1),
        "pages": page_results,
        "cache": {
            "persistent": True,
            "hits": cache_hits,
            "misses": cache_misses,
            "stored": cache_stored,
            "failed_pages_not_cached": cache_failures,
            "pruned_entries": pruned_cache_entries,
            "max_entries": TRAVEL_REOCR_CACHE_MAX_ENTRIES,
            "max_bytes": TRAVEL_REOCR_CACHE_MAX_BYTES,
            "max_age_days": TRAVEL_REOCR_CACHE_MAX_AGE_SECONDS // (24 * 60 * 60),
            "config_fingerprint": cache_config_fingerprint,
        },
        "cache_warnings": list(dict.fromkeys(cache_warnings)),
        "review_only": True,
        "auto_applied": False,
    }


def _travel_paddle_model_fingerprint(model_root: Path) -> list[dict]:
    files = []
    for model_name in ("PP-OCRv6_medium_det", "PP-OCRv6_medium_rec"):
        model_dir = model_root / model_name
        if not model_dir.is_dir():
            return [{"status": "unavailable", "model": model_name}]
        for name in ("inference.json", "inference.pdiparams", "inference.yml"):
            path = model_dir / name
            try:
                stat = path.stat()
            except OSError:
                files.append({"model": model_name, "file": name, "status": "unavailable"})
                continue
            files.append({
                "model": model_name,
                "file": name,
                "size": stat.st_size,
                "mtime_ns": stat.st_mtime_ns,
            })
    return files


def _travel_paddle_runtime_versions(python_path: Path) -> dict[str, str]:
    versions = {"paddleocr": "unknown", "paddlepaddle": "unknown"}
    site_packages = python_path.parent.parent / "Lib" / "site-packages"
    if not site_packages.is_dir():
        return versions
    for metadata in site_packages.glob("*.dist-info/METADATA"):
        package_name = metadata.parent.name.casefold()
        target = next((
            name for name in versions
            if package_name.startswith(name.replace("-", ""))
            or package_name.startswith(name + "-")
        ), None)
        if target is None:
            # Wheel distribution names use paddlepaddle and paddleocr without
            # normalising separators; account for both spellings.
            target = next((
                name for name in versions
                if package_name.startswith(name.replace("-", ""))
            ), None)
        if target is None:
            continue
        try:
            for line in metadata.read_text(encoding="utf-8", errors="replace").splitlines():
                if line.lower().startswith("version:"):
                    versions[target] = line.split(":", 1)[1].strip()
                    break
        except OSError:
            continue
    return versions


def _paddle_amount_field_evidence(blocks: list[dict]) -> list[dict]:
    """Keep OCR money tokens attached to nearby labels as review evidence.

    This deliberately reports multiple labeled values (for example fare vs.
    paid amount) instead of choosing one as an invoice total.
    """
    label_pattern = re.compile(
        r"实收\s*金额|价税合计|小写|税额|税款|AMOUNT|TOTAL|FARE|金额|合计",
        re.IGNORECASE,
    )
    money_pattern = re.compile(
        r"(?<!\d)(?:[¥￥]\s*)?(?P<number>\d{1,7}(?:,\d{3})*(?:\.\d{1,2})?)(?:\s*(?P<unit>元|CNY|RMB))?(?!\d)",
        re.IGNORECASE,
    )
    number_only_pattern = re.compile(r"^\d{1,7}(?:,\d{3})*\.\d{2}$")
    one_decimal_pattern = re.compile(r"^\d{1,7}(?:,\d{3})*\.\d$")
    number_only_pattern = re.compile(r"^\d{1,7}(?:,\d{3})*\.\d{2}$")
    labels = []
    for index, block in enumerate(blocks):
        text = str(block.get("text") or "")
        bbox = block.get("bbox_normalized")
        if not isinstance(bbox, list) or len(bbox) != 4:
            continue
        for match in label_pattern.finditer(text):
            label = match.group(0)
            normalised_label = re.sub(r"\s+", "", label).casefold()
            if any(value in normalised_label for value in ("税额", "税款")):
                field_type = "tax"
            elif any(value in normalised_label for value in ("实收金额", "价税合计", "小写", "total")):
                field_type = "total_or_paid"
            elif "fare" in normalised_label:
                field_type = "fare"
            else:
                field_type = "amount_label"
            labels.append({
                "block_index": index,
                "text": label,
                "field_type": field_type,
                "same_block_only": "小写" in normalised_label,
                "bbox_normalized": bbox,
            })

    def horizontal_gap(value_box: list, label_box: list) -> float:
        return max(
            0.0,
            float(label_box[0]) - float(value_box[2]),
            float(value_box[0]) - float(label_box[2]),
        )

    def same_row(value_box: list, label_box: list) -> bool:
        value_height = max(0.0, float(value_box[3]) - float(value_box[1]))
        label_height = max(0.0, float(label_box[3]) - float(label_box[1]))
        min_height = min(value_height, label_height)
        vertical_overlap = max(
            0.0,
            min(float(value_box[3]), float(label_box[3]))
            - max(float(value_box[1]), float(label_box[1])),
        )
        # On taxi receipts, OCR can place the label and amount in separate
        # blocks with other printed fields between them. Keep the row overlap
        # check strict, while allowing that wider horizontal spacing.
        return (
            min_height > 0
            and vertical_overlap / min_height >= 0.20
            and horizontal_gap(value_box, label_box) <= 0.12
        )

    money_block_indexes = set()
    for index, block in enumerate(blocks):
        text = str(block.get("text") or "")
        bbox = block.get("bbox_normalized")
        if not isinstance(bbox, list) or len(bbox) != 4:
            continue
        for match in money_pattern.finditer(text):
            number = match.group("number")
            token = match.group(0).strip()
            has_currency_marker = bool(
                re.search(r"[¥￥]|元|\bCNY\b|\bRMB\b", token, re.IGNORECASE)
            )
            normalized_number = number.replace(",", "")
            if (
                has_currency_marker
                or number_only_pattern.fullmatch(normalized_number)
                or one_decimal_pattern.fullmatch(normalized_number)
            ):
                money_block_indexes.add(index)

    # A page can contain side-by-side receipts whose rows overlap vertically.
    # Attach an external label only to its nearest plausible money block so a
    # total on one receipt cannot relabel a fare/date on its neighbor.
    labels_by_money_block = {}
    for label in labels:
        if label.get("same_block_only"):
            continue
        label_index = label["block_index"]
        label_box = label["bbox_normalized"]
        nearby_money_blocks = [
            index for index in money_block_indexes
            if index != label_index
            and isinstance(blocks[index].get("bbox_normalized"), list)
            and same_row(blocks[index]["bbox_normalized"], label_box)
        ]
        if not nearby_money_blocks:
            continue
        gaps = {
            index: horizontal_gap(blocks[index]["bbox_normalized"], label_box)
            for index in nearby_money_blocks
        }
        nearest_gap = min(gaps.values())
        for index, gap in gaps.items():
            if abs(gap - nearest_gap) <= 0.005:
                labels_by_money_block.setdefault(index, []).append(label)

    evidence = []
    matched_label_indexes = set()
    for block_index, block in enumerate(blocks):
        text = str(block.get("text") or "")
        bbox = block.get("bbox_normalized")
        if not isinstance(bbox, list) or len(bbox) != 4:
            continue
        block_labels = [item for item in labels if item["block_index"] == block_index]
        for match in money_pattern.finditer(text):
            number = match.group("number")
            token = match.group(0).strip()
            has_currency_marker = bool(
                re.search(r"[¥￥]|元|\bCNY\b|\bRMB\b", token, re.IGNORECASE)
            )
            is_two_decimal_number = bool(number_only_pattern.fullmatch(number.replace(",", "")))
            is_one_decimal_number = bool(one_decimal_pattern.fullmatch(number.replace(",", "")))
            nearby_labels = [
                item for item in labels
                if item["block_index"] != block_index
                and not item.get("same_block_only")
                and item in labels_by_money_block.get(block_index, [])
            ]
            linked_labels = block_labels + nearby_labels
            # A one-decimal amount such as ``66.0`` is valid evidence when it
            # is tied to an amount label on the same row, even if OCR omitted
            # the currency marker. Do not promote integers (including date
            # fragments on adjacent side-by-side receipts) into money.
            if not has_currency_marker and not (
                is_two_decimal_number
                or (is_one_decimal_number and linked_labels)
            ):
                continue
            if not linked_labels and not has_currency_marker:
                continue
            for label in linked_labels:
                matched_label_indexes.add(label["block_index"])
            types = {item["field_type"] for item in linked_labels}
            if "total_or_paid" in types:
                field_type = "total_or_paid"
            elif "tax" in types:
                field_type = "tax"
            elif "fare" in types:
                field_type = "fare"
            elif "amount_label" in types:
                field_type = "amount_label"
            else:
                field_type = "unlabelled_amount"
            evidence.append({
                "field_type": field_type,
                "value_text": number.replace(",", ""),
                "source_text": text,
                "labels": list(dict.fromkeys(item["text"] for item in linked_labels)),
                "association": (
                    "same_block" if block_labels
                    else "same_row" if nearby_labels
                    else "currency_token_only" if has_currency_marker
                    else "unlabelled_decimal"
                ),
                "confidence": block.get("confidence"),
                "bbox_normalized": bbox,
            })

    for label in labels:
        if label["block_index"] in matched_label_indexes:
            continue
        evidence.append({
            "field_type": label["field_type"],
            "value_text": None,
            "source_text": str(blocks[label["block_index"]].get("text") or ""),
            "labels": [label["text"]],
            "association": "label_without_nearby_value",
            "confidence": blocks[label["block_index"]].get("confidence"),
            "bbox_normalized": label["bbox_normalized"],
        })
    return evidence[:60]


def _travel_paddle_cache_config(dpi: int, python_path: Path, model_root: Path) -> dict:
    try:
        python_stat = python_path.stat()
        python_identity = {
            "path": str(python_path.resolve()).casefold(),
            "size": python_stat.st_size,
            "mtime_ns": python_stat.st_mtime_ns,
        }
    except OSError:
        python_identity = {"path": str(python_path).casefold(), "status": "unavailable"}
    return {
        "cache_schema": TRAVEL_PADDLE_REOCR_CACHE_SCHEMA,
        "engine": "PaddleOCR",
        "packages": _travel_paddle_runtime_versions(python_path),
        "python": python_identity,
        "models": _travel_paddle_model_fingerprint(model_root),
        "dpi": dpi,
        "crop": "lower_page_band_x0.20-0.85_y0.45-1.00_v1",
        "preprocessing": "rgb-plus-red-channel-1.5x-lanczos-v1",
        "provider_adapter": "document_processing.PaddleOCRProvider_v1",
    }


def _travel_paddle_cache_path(source_sha256: str, page: int, config: dict) -> Path:
    identity = {"source_sha256": source_sha256, "page": page, "config": config}
    encoded = json.dumps(identity, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return TRAVEL_PADDLE_REOCR_CACHE_DIR / f"{hashlib.sha256(encoded).hexdigest()}.json"


def _build_paddle_review_candidate(
    raw: dict,
    *,
    page_number: int,
    width: int,
    height: int,
    region: list,
    channel: str,
) -> dict:
    scale_factor = max(0.01, float(raw.get("scale_factor", 1.0) or 1.0))
    region_left, region_top = region[:2]
    positioned_blocks = []
    for block in raw.get("blocks", []):
        if not isinstance(block, dict):
            continue
        bbox = block.get("bbox_crop_pixels")
        absolute_bbox = None
        normalized_bbox = None
        if isinstance(bbox, list) and len(bbox) == 4 and width and height:
            absolute_bbox = [
                float(bbox[0]) / scale_factor + region_left,
                float(bbox[1]) / scale_factor + region_top,
                float(bbox[2]) / scale_factor + region_left,
                float(bbox[3]) / scale_factor + region_top,
            ]
            normalized_bbox = [
                round(absolute_bbox[0] / width, 5),
                round(absolute_bbox[1] / height, 5),
                round(absolute_bbox[2] / width, 5),
                round(absolute_bbox[3] / height, 5),
            ]
        positioned_blocks.append({
            "text": str(block.get("text") or ""),
            "confidence": block.get("confidence"),
            "bbox_pixels": absolute_bbox,
            "bbox_normalized": normalized_bbox,
        })
    is_red_channel = channel == "red"
    method = "PaddleOCR_PP-OCRv6_medium_red_channel" if is_red_channel else "PaddleOCR_PP-OCRv6_medium"
    return {
        "method": method,
        "engine": "paddleocr",
        "channel": "R" if is_red_channel else "RGB",
        "preprocessing": str(raw.get("preprocessing") or ("R-channel, 1.5x LANCZOS" if is_red_channel else "Original RGB crop")),
        "text": str(raw.get("text") or ""),
        "quality": list(_ocr_quality(str(raw.get("text") or ""))),
        "confidence": raw.get("confidence"),
        "elapsed_ms": raw.get("elapsed_ms"),
        "block_count": len(positioned_blocks),
        "blocks_truncated": len(positioned_blocks) > 300,
        "blocks": positioned_blocks[:300],
        "invoice_candidates": [],
        "amount_evidence": _paddle_amount_field_evidence(positioned_blocks),
        "amount_evidence_schema": 4,
        "warnings": list(raw.get("warnings") or []),
        "region_name": "paddle_lower_page_receipts_red_channel" if is_red_channel else "paddle_lower_page_receipts",
        "region_bbox_pixels": region,
        "region_bbox_normalized": [
            round(float(value), 5)
            for value in raw.get("region_bbox_normalized", [])
        ],
        "review_only": True,
        "page": page_number,
    }


def _prune_travel_paddle_reocr_cache(now: Optional[float] = None) -> int:
    now = time.time() if now is None else now
    removed = 0
    try:
        TRAVEL_PADDLE_REOCR_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        entries = []
        for path in TRAVEL_PADDLE_REOCR_CACHE_DIR.glob("*.json"):
            try:
                stat = path.stat()
            except OSError:
                continue
            if now - stat.st_mtime > TRAVEL_PADDLE_REOCR_CACHE_MAX_AGE_SECONDS:
                try:
                    path.unlink()
                    removed += 1
                except OSError:
                    pass
                continue
            entries.append((stat.st_mtime, path, stat.st_size))
        for path in TRAVEL_PADDLE_REOCR_CACHE_DIR.glob("*.tmp"):
            try:
                if now - path.stat().st_mtime > 24 * 60 * 60:
                    path.unlink()
                    removed += 1
            except OSError:
                pass
        entries.sort(reverse=True, key=lambda item: item[0])
        for _modified, path, _size in entries[TRAVEL_PADDLE_REOCR_CACHE_MAX_ENTRIES:]:
            try:
                path.unlink()
                removed += 1
            except OSError:
                pass
        retained = entries[:TRAVEL_PADDLE_REOCR_CACHE_MAX_ENTRIES]
        retained_bytes = sum(item[2] for item in retained)
        for _modified, path, size in reversed(retained):
            if retained_bytes <= TRAVEL_PADDLE_REOCR_CACHE_MAX_BYTES:
                break
            try:
                path.unlink()
                retained_bytes -= size
                removed += 1
            except OSError:
                pass
    except OSError:
        return removed
    return removed


def _load_travel_paddle_cache(path: Path, source_sha256: str, page: int, config: dict):
    try:
        with path.open("r", encoding="utf-8") as handle:
            entry = json.load(handle)
    except FileNotFoundError:
        return None, None
    except (OSError, UnicodeError, json.JSONDecodeError):
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass
        return None, "缓存文件损坏，已清除并准备重识别。"
    if (
        not isinstance(entry, dict)
        or entry.get("schema") != TRAVEL_PADDLE_REOCR_CACHE_SCHEMA
        or entry.get("source_sha256") != source_sha256
        or entry.get("page") != page
        or entry.get("config") != config
        or not isinstance(entry.get("candidate"), dict)
    ):
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass
        return None, "缓存记录与当前 PDF、页码或 Paddle 模型配置不一致，已清除并准备重识别。"
    candidate = entry["candidate"]
    if candidate.get("error"):
        try:
            failed_at = datetime.fromisoformat(str(entry.get("cached_at", ""))).timestamp()
        except (TypeError, ValueError, OverflowError):
            failed_at = 0
        if not failed_at or time.time() - failed_at > 24 * 60 * 60:
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass
            return None, "此前 Paddle 复读超时记录已过期，将自动重试。"
    try:
        os.utime(path, None)
    except OSError:
        pass
    return copy.deepcopy(candidate), None


def _store_travel_paddle_cache(
    path: Path,
    source_sha256: str,
    page: int,
    config: dict,
    candidate: dict,
) -> Optional[str]:
    if not str(candidate.get("text", "")).strip() and not candidate.get("error"):
        return ""
    entry = {
        "schema": TRAVEL_PADDLE_REOCR_CACHE_SCHEMA,
        "source_sha256": source_sha256,
        "page": page,
        "config": config,
        "cached_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "candidate": candidate,
    }
    temporary_path = None
    try:
        TRAVEL_PADDLE_REOCR_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", newline="\n",
            dir=TRAVEL_PADDLE_REOCR_CACHE_DIR,
            prefix=f".{path.stem}-", suffix=".tmp", delete=False,
        ) as handle:
            temporary_path = Path(handle.name)
            json.dump(entry, handle, ensure_ascii=False, separators=(",", ":"))
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_path, path)
        _prune_travel_paddle_reocr_cache()
        return None if path.exists() else "识别结果超过缓存容量上限，当前可查看但未持久化。"
    except (OSError, TypeError, ValueError) as exc:
        return f"PaddleOCR 缓存未能写入：{type(exc).__name__}；本次结果仍可查看。"
    finally:
        if temporary_path is not None:
            try:
                temporary_path.unlink(missing_ok=True)
            except OSError:
                pass


def reocr_pdf_pages_with_paddle_for_review(
    path: str,
    page_numbers: list[int],
    dpi: int = 600,
) -> dict:
    """Run optional isolated PaddleOCR over the lower receipt band of pages.

    Only text blocks and their mapped full-page coordinates are returned.
    Page images are temporary and Paddle results never update expenses.
    """
    pages = list(dict.fromkeys(int(value) for value in page_numbers))
    if not pages or len(pages) > 3 or any(value < 1 for value in pages):
        raise ValueError("PaddleOCR 区域复核一次需要选择 1 到 3 个有效页码。")
    dpi = max(300, min(int(dpi), 600))
    root = Path(__file__).resolve().parent
    try:
        import config
        python_path = Path(config.PADDLE_PYTHON)
        model_root = Path(config.PADDLE_PDX_CACHE_HOME) / "official_models"
        temp_dir = Path(config.PADDLE_TEMP_DIR)
    except (ImportError, AttributeError):
        python_path = root / ".paddle-venv" / "Scripts" / "python.exe"
        model_root = root / ".paddle-cache" / "paddlex" / "official_models"
        temp_dir = root / ".paddle-tmp"
    if not python_path.is_file():
        raise RuntimeError("未找到隔离 PaddleOCR 环境 .paddle-venv；主程序 OCR 环境未受影响。")
    if any(not (model_root / name).is_dir() for name in (
        "PP-OCRv6_medium_det", "PP-OCRv6_medium_rec",
    )):
        raise RuntimeError("本机 PaddleOCR 模型缓存不完整，未启动识别，也未尝试下载模型。")
    temp_dir.mkdir(parents=True, exist_ok=True)

    digest = hashlib.sha256()
    with open(path, "rb") as source_handle:
        for chunk in iter(lambda: source_handle.read(1024 * 1024), b""):
            digest.update(chunk)
    source_sha256 = digest.hexdigest()
    cache_config = _travel_paddle_cache_config(dpi, python_path, model_root)
    config_fingerprint = hashlib.sha256(
        json.dumps(cache_config, sort_keys=True, separators=(",", ":")).encode("utf-8")
    ).hexdigest()
    pruned = _prune_travel_paddle_reocr_cache()
    page_candidates = {}
    cache_hits = 0
    cache_misses = 0
    cache_stored = 0
    cache_failures = 0
    cache_warnings = []
    missing_pages = []
    cache_paths = {}
    for page_number in pages:
        cache_path = _travel_paddle_cache_path(source_sha256, page_number, cache_config)
        cache_paths[page_number] = cache_path
        lookup_started = time.perf_counter()
        cached, warning = _load_travel_paddle_cache(
            cache_path, source_sha256, page_number, cache_config,
        )
        lookup_ms = round((time.perf_counter() - lookup_started) * 1000, 1)
        if warning:
            cache_warnings.append(f"第 {page_number} 页：{warning}")
        if cached is not None:
            cached_candidates = [cached, *(
                cached.get("alternate_candidates", [])
                if isinstance(cached.get("alternate_candidates"), list) else []
            )]
            migration_needed = False
            for cached_candidate in cached_candidates:
                if not isinstance(cached_candidate, dict):
                    continue
                if cached_candidate.get("amount_evidence_schema") != 4:
                    cached_candidate["amount_evidence"] = _paddle_amount_field_evidence(
                        cached_candidate.get("blocks", []),
                    )
                    cached_candidate["amount_evidence_schema"] = 4
                    migration_needed = True
            if migration_needed:
                migration_warning = _store_travel_paddle_cache(
                    cache_path, source_sha256, page_number, cache_config, cached,
                )
                if migration_warning:
                    cache_warnings.append(f"第 {page_number} 页：{migration_warning}")
            cached["ocr_elapsed_ms"] = cached.get("elapsed_ms")
            cached["elapsed_ms"] = lookup_ms
            cached["cache_status"] = "HIT"
            for alternate in cached_candidates[1:]:
                if isinstance(alternate, dict):
                    alternate["cache_status"] = "HIT"
            page_candidates[page_number] = cached
            cache_hits += 1
        else:
            cache_misses += 1
            missing_pages.append(page_number)

    started = time.perf_counter()
    if missing_pages:
        output_fd, output_name = tempfile.mkstemp(
            prefix="travel-paddle-worker-", suffix=".json", dir=temp_dir,
        )
        os.close(output_fd)
        output_path = Path(output_name)
        try:
            output_path.unlink(missing_ok=True)
            worker = root / "document_processing" / "paddle_review_worker.py"
            command = [
                str(python_path), "-u", str(worker),
                "--pdf", str(Path(path).resolve()),
                "--pages", ",".join(str(value) for value in missing_pages),
                "--dpi", str(dpi),
                "--output", str(output_path),
            ]
            worker_process = subprocess.Popen(
                command,
                cwd=str(root),
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) if os.name == "nt" else 0,
            )
            worker_timed_out = False
            try:
                worker_process.wait(timeout=360)
            except subprocess.TimeoutExpired:
                worker_timed_out = True
                timeout_message = "PaddleOCR 单次推理超过 6 分钟，已终止工作进程；24 小时后自动重试。"
                if os.name == "nt":
                    try:
                        subprocess.run(
                            ["taskkill", "/PID", str(worker_process.pid), "/T", "/F"],
                            capture_output=True,
                            text=True,
                            encoding="utf-8",
                            errors="replace",
                            timeout=15,
                            check=False,
                        )
                    except (OSError, subprocess.SubprocessError):
                        pass
                if worker_process.poll() is None:
                    try:
                        worker_process.kill()
                    except OSError:
                        pass
                try:
                    worker_process.wait(timeout=15)
                except subprocess.TimeoutExpired:
                    if os.name == "nt":
                        try:
                            subprocess.run(
                                ["taskkill", "/PID", str(worker_process.pid), "/T", "/F"],
                                capture_output=True,
                                text=True,
                                encoding="utf-8",
                                errors="replace",
                                timeout=10,
                                check=False,
                            )
                        except (OSError, subprocess.SubprocessError):
                            pass
                    try:
                        worker_process.kill()
                    except OSError:
                        pass
                    try:
                        worker_process.wait(timeout=10)
                    except subprocess.TimeoutExpired:
                        pass
                for page_number in missing_pages:
                    timeout_candidate = {
                        "page": page_number,
                        "text": "",
                        "error": timeout_message,
                        "cache_status": "TIMEOUT_CACHED",
                        "amount_evidence": [],
                    }
                    _store_travel_paddle_cache(
                        cache_paths[page_number],
                        source_sha256,
                        page_number,
                        cache_config,
                        timeout_candidate,
                    )
                    page_candidates[page_number] = timeout_candidate
                    cache_failures += 1
            if not worker_timed_out:
                completed = subprocess.CompletedProcess(
                    command, worker_process.returncode, "", "",
                )
                if completed.returncode != 0:
                    detail = (completed.stderr or completed.stdout or "").strip()
                    detail = detail[-1600:] if detail else "无工作进程错误详情。"
                    raise RuntimeError(f"隔离 PaddleOCR 工作进程失败：{detail}")
                try:
                    with output_path.open("r", encoding="utf-8") as handle:
                        worker_result = json.load(handle)
                except (OSError, UnicodeError, json.JSONDecodeError) as exc:
                    raise RuntimeError(f"PaddleOCR 工作进程未生成有效结果：{type(exc).__name__}") from exc
                raw_pages = {
                    int(item.get("page")): item
                    for item in worker_result.get("pages", [])
                    if isinstance(item, dict) and item.get("page") is not None
                }
                for page_number in missing_pages:
                    raw = raw_pages.get(page_number)
                    if raw is None:
                        page_candidates[page_number] = {
                            "page": page_number,
                            "error": "PaddleOCR 工作进程未返回该页结果。",
                            "cache_status": "MISS_FAILED_NOT_CACHED",
                        }
                        cache_failures += 1
                        continue
                    width, height = raw.get("render_size", [0, 0])
                    region = raw.get("region_bbox_pixels", [0, 0, 0, 0])
                    candidate = _build_paddle_review_candidate(
                        raw,
                        page_number=page_number,
                        width=width,
                        height=height,
                        region=region,
                        channel="rgb",
                    )
                    candidate["cache_status"] = "MISS"
                    red_raw = raw.get("red_channel")
                    if isinstance(red_raw, dict):
                        alternate = _build_paddle_review_candidate(
                            {
                                **red_raw,
                                "region_bbox_normalized": raw.get("region_bbox_normalized", []),
                            },
                            page_number=page_number,
                            width=width,
                            height=height,
                            region=region,
                            channel="red",
                        )
                        alternate["cache_status"] = "MISS"
                        candidate["alternate_candidates"] = [alternate]
                    cache_warning = _store_travel_paddle_cache(
                        cache_paths[page_number], source_sha256, page_number,
                        cache_config, candidate,
                    )
                    if cache_warning is None:
                        candidate["cache_status"] = "MISS_STORED"
                        cache_stored += 1
                    elif cache_warning:
                        candidate["cache_status"] = "MISS_WRITE_FAILED"
                        candidate["cache_warning"] = cache_warning
                        cache_warnings.append(f"第 {page_number} 页：{cache_warning}")
                    else:
                        candidate["cache_status"] = "MISS_NOT_CACHEABLE"
                    for alternate in candidate.get("alternate_candidates", []):
                        alternate["cache_status"] = candidate["cache_status"]
                    page_candidates[page_number] = candidate
        finally:
            try:
                output_path.unlink(missing_ok=True)
            except OSError:
                pass

    return {
        "engine": "PaddleOCR",
        "model": "PP-OCRv6 medium det+rec",
        "dpi": dpi,
        "source_sha256": source_sha256,
        "requested_pages": pages,
        "elapsed_ms": round((time.perf_counter() - started) * 1000, 1),
        "pages": [page_candidates[number] for number in pages],
        "cache": {
            "persistent": True,
            "hits": cache_hits,
            "misses": cache_misses,
            "stored": cache_stored,
            "failed_pages_not_cached": cache_failures,
            "pruned_entries": pruned,
            "max_entries": TRAVEL_PADDLE_REOCR_CACHE_MAX_ENTRIES,
            "max_bytes": TRAVEL_PADDLE_REOCR_CACHE_MAX_BYTES,
            "max_age_days": TRAVEL_PADDLE_REOCR_CACHE_MAX_AGE_SECONDS // (24 * 60 * 60),
            "config_fingerprint": config_fingerprint,
        },
        "cache_warnings": list(dict.fromkeys(cache_warnings)),
        "review_only": True,
        "auto_applied": False,
    }


def _pdf_page_needs_ocr(text: str) -> bool:
    """Detect blank, nearly blank, or unusable hidden OCR text layers."""
    compact = re.sub(r"\s+", "", text or "")
    if not compact or len(compact) < 80:
        return True
    # A long native text page is already usable, even when it is only a
    # heading or an approval page.  Restricting OCR to short/empty layers is
    # important for 40+ page payment-request PDFs.
    return False


def _read_pdf(path: str) -> tuple[str, list[str]]:
    try:
        import pdfplumber
    except ImportError as exc:  # pragma: no cover - environment-specific
        raise RuntimeError("未安装 pdfplumber，无法读取 PDF") from exc

    native_pages: list[str] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            native_pages.append(page.extract_text() or "")

    ocr_targets = [index for index, page_text in enumerate(native_pages) if _pdf_page_needs_ocr(page_text)]
    ocr_pages, ocr_errors = _ocr_pdf_pages(path, ocr_targets)
    pages = [ocr_pages.get(index, page_text) or page_text for index, page_text in enumerate(native_pages)]
    text = "\n\f\n".join(pages)
    warnings: list[str] = []
    if ocr_targets:
        if ocr_pages:
            scope = "PDF" if len(ocr_targets) == len(native_pages) else "PDF 部分页面"
            warnings.append(
                f"{scope} 没有可提取文本，已执行本地 OCR"
                f"（成功识别 {len(ocr_pages)}/{len(ocr_targets)} 页）；金额结果仍需人工复核。"
            )
        else:
            warnings.append("PDF 没有可提取文本，尝试本地 OCR 失败，请人工复核。")
        failed_page_numbers = [index + 1 for index in ocr_targets if index not in ocr_pages]
        if failed_page_numbers:
            warnings.append(
                "OCR 未成功页码：" + ", ".join(map(str, failed_page_numbers))
            )
        if ocr_errors:
            warnings.append(f"OCR 诊断：{'; '.join(ocr_errors[:3])}")
    if not text.strip():
        warnings.append("OCR 未提取到有效文字，请人工复核原始 PDF。")
    return text, warnings


def _read_docx(path: str) -> tuple[str, list[str]]:
    try:
        from docx import Document
    except ImportError as exc:  # pragma: no cover - environment-specific
        raise RuntimeError("未安装 python-docx，无法读取 Word 文件") from exc

    document = Document(path)
    chunks: list[str] = [p.text for p in document.paragraphs if p.text.strip()]
    for table in document.tables:
        for row in table.rows:
            chunks.append(" ".join(cell.text for cell in row.cells if cell.text.strip()))
    text = "\n".join(chunks)
    warnings: list[str] = []
    if not text.strip():
        warnings.append("Word 文件没有可提取文本，请人工复核。")
    return text, warnings


def extract_document_text(path: str) -> tuple[str, list[str]]:
    ext = Path(path).suffix.lower()
    if ext == ".pdf":
        return _read_pdf(path)
    if ext == ".docx":
        return _read_docx(path)
    if ext == ".doc":
        raise RuntimeError("暂不支持旧版 .doc 二进制文件，请另存为 .docx 后上传。")
    raise RuntimeError(f"不支持的文件类型: {ext or '未知'}，请上传 PDF 或 DOCX。")


def _filename_hints(path: str) -> tuple[Optional[date], Optional[date], str, str]:
    stem = Path(path).stem
    year = datetime.now().year
    start, end = _date_span(stem, year)

    # Scanner exports often use YYYYMMDD, e.g. 20260816, while older files
    # use a leading MMDD or MMDD-MMDD range.  Do not interpret the second
    # compact token in names such as 20260816_0822 as a second travel date;
    # it is normally the scan/export time.
    compact_full_dates: list[tuple[date, int]] = []
    for match in re.finditer(r"(?<!\d)(20\d{2})(\d{2})(\d{2})(?!\d)", stem):
        try:
            compact_full_dates.append((date(
                int(match.group(1)), int(match.group(2)), int(match.group(3)),
            ), match.end()))
        except ValueError:
            continue
    if compact_full_dates:
        start, end = compact_full_dates[0][0], compact_full_dates[-1][0]
        # Scanner packs use YYYYMMDD_MMDD for the trip range. The second
        # token is not an export time in these business-trip filenames.
        if len(compact_full_dates) == 1:
            suffix = stem[compact_full_dates[0][1]:]
            short_end = re.match(r"[_\s-]+(\d{2})(\d{2})(?!\d)", suffix)
            if short_end:
                try:
                    candidate = date(start.year, int(short_end.group(1)), int(short_end.group(2)))
                    if candidate >= start:
                        end = candidate
                except ValueError:
                    pass
    elif not start:
        compact = re.search(r"(?<!\d)(\d{2})(\d{2})(?!\d)", stem)
        if compact:
            try:
                start = end = date(year, int(compact.group(1)), int(compact.group(2)))
            except ValueError:
                pass

    lower_stem = stem.casefold()
    known_name = next(
        (target for alias, target in sorted(NAME_MAP.items(), key=lambda item: len(item[0]), reverse=True)
         if alias.casefold() in lower_stem),
        "",
    )
    filename_cities = _city_mentions_in(stem)
    # A single explicit city is useful filename metadata. A filename listing
    # several stops is an itinerary; choosing whichever alias sorts first can
    # incorrectly turn one intermediate stop into the whole trip destination.
    known_city = filename_cities[0] if len(filename_cities) == 1 else ""

    parts = [p for p in re.split(r"[_\s]+", stem) if p]
    tail: list[str] = []
    for index, part in enumerate(parts):
        if part.casefold() == "iav":
            tail = parts[index + 1:]
            break
    ignored = {"v2", "logistic", "toll", "fee", "scannerfiles", "scanner", "files", "scan"}
    tail = [
        part for part in tail
        if part.casefold() not in ignored
        and not re.fullmatch(r"20\d{6}|\d{4}", part)
    ]
    name_hint = known_name
    city_hint = known_city
    if not name_hint:
        if len(tail) >= 2 and not city_hint:
            city_hint = tail[-1]
            name_hint = " ".join(tail[:-1])
        elif tail:
            name_hint = tail[0]
    if not name_hint:
        # Some supplier exports put the employee name directly before the
        # document type, e.g. ``Stephan机票发票_part2.pdf``.
        marker = re.search(
            r"(?:机票|火车票|发票|打车|费用|flight|train|ticket|invoice|hotel|taxi)",
            stem,
            flags=re.IGNORECASE,
        )
        if marker:
            prefix = stem[:marker.start()].strip(" _-")
            prefix = re.sub(r"^(?:20\d{6}|\d{4}[-_])", "", prefix).strip(" _-")
            if prefix:
                name_hint = prefix.split("_")[-1].strip()
    return start, end, name_hint, city_hint


def _normalise_city(value: str) -> str:
    value = re.sub(r"\s+", " ", value.strip(" _-"))
    if value in CITY_MAP:
        return CITY_MAP[value]
    for source, target in sorted(CITY_MAP.items(), key=lambda item: len(item[0]), reverse=True):
        if source.lower() in value.lower():
            return target
    return value


def _known_city_hint(value: str) -> str:
    """Return a normalized city only when the hint matches a known city.

    A filename such as ``0902_IAV_Qiu Dong_Tianjin.pdf`` is a stronger
    destination signal than a city embedded in a transport route later in
    the document.  Unknown filename suffixes are not treated as destinations
    because the filename parser also has to support files without a city
    suffix (for example ``..._Qiu_Dong.pdf``).
    """
    if not value:
        return ""
    normalized = _normalise_city(value)
    known_cities = set(CITY_MAP.values())
    return normalized if normalized in known_cities else ""


def _city_mentions_in(value: str) -> list[str]:
    """Return distinct known cities in text order, preferring longer aliases."""
    matches: list[tuple[int, int, int, str]] = []
    for alias, target in CITY_MAP.items():
        flags = re.IGNORECASE if alias.isascii() else 0
        for match in re.finditer(re.escape(alias), value, flags=flags):
            matches.append((match.start(), -len(alias), match.end(), target))
    matches.sort()
    cities: list[str] = []
    consumed_until = -1
    for start, _negative_length, end, target in matches:
        if start < consumed_until:
            continue
        consumed_until = end
        if target not in cities:
            cities.append(target)
    return cities


def _city_mentioned_in(value: str) -> str:
    """Return the first known city in a station, route, or note fragment."""
    cities = _city_mentions_in(value)
    return cities[0] if cities else ""


def _ticket_route_events(text: str, default_year: int) -> list[dict]:
    """Read dated ticket routes from route strings or labelled ticket fields."""
    origin_labels = re.compile(
        r"出发地|出发站|出发城市|起飞地|始发地|始发站|起点|\bFROM\b|\bORIGIN\b|(?<!\S)自(?=\s*[:：])",
        re.IGNORECASE,
    )
    destination_labels = re.compile(
        r"到达地|到达站|到达城市|目的地|目的站|终点|降落地|\bTO\b|\bDESTINATION\b|(?<!\S)至(?=\s*[:：])",
        re.IGNORECASE,
    )
    all_labels = re.compile(
        rf"(?:{origin_labels.pattern})|(?:{destination_labels.pattern})",
        re.IGNORECASE,
    )
    def label_city(page: str, labels: re.Pattern) -> str:
        for label in labels.finditer(page):
            next_label = all_labels.search(page, label.end())
            stop = next_label.start() if next_label else min(len(page), label.end() + 100)
            fragment = page[label.end():stop]
            city = _city_mentioned_in(fragment)
            if city:
                return city
            code_match = re.search(r"\b([A-Z]{3})\b", fragment.upper())
            if code_match and code_match.group(1) in AIRPORT_CODE_CITY_MAP:
                return AIRPORT_CODE_CITY_MAP[code_match.group(1)]
        return ""

    routes: list[dict] = []
    sequence = 0

    def add_route(origin: str, destination: str, context: str, page_text: str, page_number: int) -> None:
        nonlocal sequence
        if not origin or not destination or origin == destination:
            return
        dates = _date_tokens(context, default_year)
        if not dates:
            dates = _date_tokens(page_text, default_year)
        time_match = TIME_RE.search(context) or TIME_RE.search(page_text)
        minutes = (
            int(time_match.group("hour")) * 60 + int(time_match.group("minute"))
            if time_match else None
        )
        routes.append({
            "date": dates[0] if dates else None,
            "minutes": minutes,
            "sequence": sequence,
            "page": page_number,
            "origin": origin,
            "destination": destination,
        })
        sequence += 1

    pages = text.replace("\r\n", "\n").split("\f")
    ticket_page_marker = re.compile(
        r"机票|火车票|车票|高铁|动车|铁路|航班|航空|登机牌|登机凭证|电子客票|航空运输电子客票|boarding\s+pass|flight|airline|train|railway",
        re.IGNORECASE,
    )
    for page_number, raw_page in enumerate(pages, 1):
        page = _normalise_text(raw_page)
        if not ticket_page_marker.search(page):
            continue

        for raw_line in page.splitlines():
            line = _normalise_text(raw_line)
            if not line or _category_for_line(line) != "ticket" or not _looks_like_route(line):
                continue
            separator = re.search(r"[-—－]", line)
            if not separator:
                continue
            add_route(
                _city_mentioned_in(line[:separator.start()]),
                _city_mentioned_in(line[separator.end():]),
                line,
                page,
                page_number,
            )

        origin = label_city(page, origin_labels)
        destination = label_city(page, destination_labels)
        if origin and destination:
            add_route(origin, destination, page, page, page_number)

    # The same flight may appear as an e-ticket invoice and a boarding pass.
    # Keep both as evidence, but collapse duplicates before chronology sorting.
    unique_routes: list[dict] = []
    seen: set[tuple[Optional[date], str, str]] = set()
    for route in routes:
        key = (route["date"], route["origin"], route["destination"])
        if key in seen:
            continue
        seen.add(key)
        unique_routes.append(route)

    if any(route["minutes"] is not None for route in unique_routes):
        unique_routes.sort(key=lambda route: (
            route["date"] or date.max,
            route["minutes"] if route["minutes"] is not None else 24 * 60 + route["sequence"],
            route["sequence"],
        ))
    else:
        unique_routes.sort(key=lambda route: (
            route["date"] or date.max,
            route["sequence"],
        ))
    return unique_routes


def _ticket_route_city(text: str, default_year: int) -> str:
    """Infer the first arrival city from the chronological ticket route."""
    routes = _ticket_route_events(text, default_year)
    return routes[0]["destination"] if routes else ""


def _get_travel_policy(policy_name: str = "VGC") -> dict:
    policies = load_travel_policies().get("policies", {})
    return policies.get(policy_name) or policies.get("VGC") or DEFAULT_TRAVEL_POLICY


def _policy_level_key(city: str, policy_name: str = "VGC") -> str:
    """Return the same level key used by the deterministic policy engine."""
    return _travel_policy_engine(policy_name).resolve_city(city)["level"]


def _city_level(city: str, policy_name: str = "VGC") -> float:
    key = _policy_level_key(city, policy_name)
    try:
        return float(key)
    except ValueError:
        return 2.0


def _is_defined_policy_city(city: str, policy_name: str = "VGC") -> bool:
    return bool(_travel_policy_engine(policy_name).resolve_city(city)["known_city"])


def _accommodation_rate(city: str, policy_name: str = "VGC") -> float:
    return float(_travel_policy_engine(policy_name).resolve_city(city)["accommodation_per_night"])


def _allowance_rate(city: str, policy_name: str = "VGC") -> float:
    return float(_travel_policy_engine(policy_name).resolve_city(city)["allowance_per_day"])


def _travel_policy_engine(policy_name: str = "VGC") -> TravelPolicyEngine:
    return TravelPolicyEngine(load_travel_policies(), policy_name=policy_name)


def _normalise_name(value: str) -> str:
    value = re.sub(r"\s+", " ", value.strip(" _-"))
    if value in NAME_MAP:
        return NAME_MAP[value]
    for alias, target in NAME_MAP.items():
        if value.casefold() == alias.casefold():
            return target
    compact = value.replace(" ", "")
    if compact in NAME_MAP:
        return NAME_MAP[compact]
    for alias, target in NAME_MAP.items():
        if compact.casefold() == alias.replace(" ", "").casefold():
            return target
    if "," in value:
        left, right = [part.strip() for part in value.split(",", 1)]
        return f"{left}, {right}"
    parts = value.split()
    if len(parts) >= 2:
        return f"{parts[0]}, {' '.join(parts[1:])}"
    return value


def _infer_trip_identity(text: str, path: str, default_year: int) -> tuple[date, date, str, str]:
    filename_start, filename_end, filename_name, filename_city = _filename_hints(path)
    header_start: Optional[date] = None
    header_end: Optional[date] = None
    header_line = ""
    for line in text.replace("\f", "\n").splitlines()[:40]:
        line = _normalise_text(line)
        start, end = _date_span(line, default_year)
        if start and end and start != end:
            header_start, header_end, header_line = start, end, line
            break

    # A dated scanner export name is usually the only reliable metadata left
    # when the image contains no readable heading.  Prefer it when present;
    # text-document names have no date and continue using the document header.
    start = filename_start or header_start
    end = filename_end or header_end
    if not start:
        # A single-day invoice is still a valid trip group.
        start = end = _date_tokens(text[:3000], default_year)[0] if _date_tokens(text[:3000], default_year) else date(default_year, 1, 1)
    if not end:
        end = start
    if end < start:
        end = end.replace(year=end.year + 1) if end.month < start.month else start

    # Prefer an explicit, known city in the filename.  A full-text scan can
    # otherwise select the departure city from a route such as
    # “北京南-天津南” before it reaches the actual destination.
    city = _known_city_hint(filename_city)
    # A filename alias is more trustworthy than English OCR over a Chinese
    # name, which Tesseract may turn into unrelated Latin characters.
    name = _normalise_name(filename_name) if filename_name else ""
    identity_text = header_line or text[:3000]

    # When the filename has no usable city, use the earliest ticket route.
    # Prefer the route over a title/full-text scan because a title or a taxi
    # line may mention the departure city rather than the destination.
    if not city:
        city = _ticket_route_city(text, default_year)

    # If ticket routes are unavailable, inspect the leading identity lines.
    # Do not use route lines here: they contain both origin and destination.
    if not city:
        for raw_line in text.replace("\f", "\n").splitlines()[:20]:
            line = _normalise_text(raw_line)
            if not line or _looks_like_route(line):
                continue
            city = _city_mentioned_in(line)
            if city:
                break

    # Last resort for older documents whose city appears only in body text.
    if not city:
        for source, target in sorted(CITY_MAP.items(), key=lambda item: len(item[0]), reverse=True):
            if source.lower() in identity_text.lower():
                city = target
                break
    for source, target in sorted(NAME_MAP.items(), key=lambda item: len(item[0]), reverse=True):
        if source.lower() in identity_text.lower():
            name = target
            break
    city = city or _normalise_city(filename_city)
    name = name or _normalise_name(filename_name)
    if not city:
        city = "Unknown"
    if not name:
        name = "Unknown"
    return start, end, name, city


def _category_for_line(line: str) -> Optional[str]:
    lowered = line.lower()
    # Airline and railway invoices often print fuel surcharges, toll terms or
    # other fee wording on the same page. The ticket type is the stronger
    # business signal; those ancillary costs stay with the ticket invoice.
    if any(keyword.lower() in lowered for keyword in TICKET_KEYWORDS):
        return "ticket"
    # Taxi route lines can contain hotel names (e.g. “康莱德酒店-亚朵酒店”).
    # Treat a route as transport before looking for accommodation keywords.
    if _looks_like_route(line):
        return "taxi"
    if any(keyword.lower() in lowered for keyword in HOTEL_KEYWORDS):
        return "accommodation"
    if (
        any(keyword.lower() in lowered for keyword in TAXI_KEYWORDS)
        or any(keyword.lower() in lowered for keyword in SPECIAL_KEYWORDS)
    ):
        return "taxi"
    return None


def _category_from_filename(filename: str) -> Optional[str]:
    """Use explicit expense words in a filename, never route punctuation.

    Server-side upload paths contain a hyphenated UUID prefix. Feeding that
    temporary name into ``_category_for_line`` makes it look like a taxi route
    and incorrectly forces every page into the transport category.
    """
    lowered = Path(filename).stem.lower()
    if any(keyword.lower() in lowered for keyword in TICKET_KEYWORDS):
        return "ticket"
    if any(keyword.lower() in lowered for keyword in HOTEL_KEYWORDS):
        return "accommodation"
    if (
        any(keyword.lower() in lowered for keyword in TAXI_KEYWORDS)
        or any(keyword.lower() in lowered for keyword in SPECIAL_KEYWORDS)
    ):
        return "taxi"
    return None


def _amount_from_line(line: str) -> Optional[float]:
    """Prefer the amount label/currency over dates and invoice numbers."""
    line = _normalise_text(line)
    if not line:
        return None

    # Word summaries frequently use expressions such as
    # “140.45+65.74+50.1=256.19”. The value after '=' is the section total;
    # individual dated lines are parsed separately below.
    equals_matches = re.findall(r"=\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)", line)
    for match in reversed(equals_matches):
        amount = _plausible_amount(match)
        if amount is not None:
            return amount

    label_patterns = (
        r"(?:价税合计|小写|金额|总额|合计|票价|车票|住宿|酒店住宿|网约车|打车|费用|amount|total|fare|fee)"
        r"[^0-9]{0,24}([0-9][0-9,]*(?:\.[0-9]{1,2})?)",
    )
    for pattern in label_patterns:
        matches = re.findall(pattern, line, flags=re.IGNORECASE)
        for match in reversed(matches):
            amount = _plausible_amount(match)
            if amount is not None:
                return amount

    currency_matches = re.findall(r"(?:¥|RMB|CNY)\s*:?[\s]*([0-9][0-9,]*(?:\.[0-9]{1,2})?)", line, flags=re.IGNORECASE)
    for match in reversed(currency_matches):
        amount = _plausible_amount(match)
        if amount is not None:
            return amount

    yuan_matches = re.findall(r"([0-9][0-9,]*(?:\.[0-9]{1,2})?)\s*(?:元|yuan|RMB)\b", line, flags=re.IGNORECASE)
    for match in reversed(yuan_matches):
        amount = _plausible_amount(match)
        if amount is not None:
            return amount

    # Route-style taxi lines often omit the currency symbol, so use the final
    # standalone number only when it is not a date-only line.
    trailing = re.search(r"(?:^|\s|:)([0-9][0-9,]*(?:\.[0-9]{1,2})?)\s*$", line)
    if trailing:
        return _plausible_amount(trailing.group(1))
    return None


def _plausible_amount(value: object) -> Optional[float]:
    """Reject OCR artefacts such as invoice IDs and phone numbers as money."""
    try:
        amount = round(float(str(value).replace(",", "").replace(" ", "")), 2)
    except (TypeError, ValueError):
        return None
    if 0 < amount <= MAX_REASONABLE_EXPENSE:
        return amount
    return None


def _ocr_gross_total_matches(page: str) -> list[tuple[float, int, int]]:
    """Return every readable payable total, including standalone small-case totals.

    Chinese VAT invoices place ``小写 ¥...`` beside the total. OCR often drops
    or fragments ``价税合计`` while preserving that adjacent label, so it is an
    independent anchor. Position-based deduplication prevents the same printed
    amount being emitted once for each label while preserving two equal-valued
    invoices at different positions on one page.
    """
    normalised = _normalise_text(page)
    upper = normalised.upper()
    labels = list(re.finditer(
        r"价\s*税\s*合\s*计|GRAND\s+TOTAL|小\s*写",
        upper,
        flags=re.IGNORECASE,
    ))
    results: list[tuple[float, int, int, int]] = []
    amount_pattern = r"([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)"
    for index, label in enumerate(labels):
        is_small_case = bool(re.fullmatch(r"小\s*写", label.group(0), flags=re.IGNORECASE))
        next_gross_label = next((
            later.start() for later in labels[index + 1:]
            if not re.fullmatch(r"小\s*写", later.group(0), flags=re.IGNORECASE)
        ), len(upper))
        next_label = labels[index + 1].start() if index + 1 < len(labels) else len(upper)
        window_end = min(
            label.end() + (90 if is_small_case else 260),
            next_label if is_small_case else next_gross_label,
        )
        window = upper[label.end():window_end]
        # X/Y are common OCR substitutions for the yuan glyph. They are safe
        # here because matching is restricted to a labelled total area.
        currency_amount = re.search(
            rf"(?:¥|RMB|CNY|(?<![A-Z])Y|(?<![A-Z])X)\s*:?\s*{amount_pattern}",
            window,
        )
        amount_match = currency_amount
        if is_small_case and not amount_match:
            # OCR can mistake the amount next to ``小写`` and then pick up a
            # vehicle plate or invoice number from the following line. A
            # currency-marked amount may cross a line; an unmarked fallback
            # must stay on the label line and cannot be embedded in text.
            same_line = window.splitlines()[0][:40]
            amount_match = re.search(
                rf"(?<![0-9A-Z\u4e00-\u9fff]){amount_pattern}(?![0-9A-Z\u4e00-\u9fff])",
                same_line,
            )
        if not amount_match:
            continue
        amount = _plausible_amount(amount_match.group(1))
        if amount is not None:
            amount_start = label.end() + amount_match.start()
            if is_small_case and not currency_amount and amount_match.start() > 0:
                amount_start += max(0, amount_match.group(0).find(amount_match.group(1)))
            results.append((
                amount, label.start(), label.end() + amount_match.end(), amount_start,
            ))
    results.sort(key=lambda item: (item[3], item[0]))
    unique: list[tuple[float, int, int]] = []
    amount_positions: set[tuple[float, int]] = set()
    for amount, label_start, match_end, amount_start in results:
        key = (round(amount, 2), amount_start)
        if key in amount_positions:
            continue
        amount_positions.add(key)
        unique.append((amount, label_start, match_end))
    return unique


def _ocr_ticket_total_matches(
    page: str,
    gross_matches: list[tuple[float, int, int]],
) -> list[float]:
    """Recover currency-marked ticket totals that use ``合计`` without ``小写``.

    Flight e-tickets commonly print their total as ``合计 CNY 1560.00``. The
    same scanned page can also contain a separate VAT service-fee invoice, so
    only accept this label when nearby ticket-fare evidence identifies it.
    """
    normalised = _normalise_text(page)
    upper = normalised.upper()
    amount_pattern = r"([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)"
    ticket_markers = (
        "标价", "燃油附加费", "民航发展基金", "电子客票号码", "电子客票",
        "航班号", "承运人", "AIRLINE", "FLIGHT",
    )
    existing_amounts = {round(amount, 2) for amount, _start, _end in gross_matches}
    results: list[float] = []
    currency_pattern = re.compile(
        rf"(?:¥|RMB|CNY)\s*:?\s*{amount_pattern}",
        flags=re.IGNORECASE,
    )
    for label in re.finditer(r"合\s*计", upper):
        if re.search(r"价\s*税\s*$", upper[max(0, label.start() - 12):label.start()]):
            continue
        prefix = normalised[max(0, label.start() - 260):label.start()].casefold()
        if not any(marker.casefold() in prefix for marker in ticket_markers):
            continue

        # The fare table may be OCR'd in label-first order: one ``合计`` label
        # is followed by base fare, fuel surcharge, tax, fund, then the final
        # total. Stop before the e-ticket number or the next VAT invoice and
        # use the last currency amount within this ticket block.
        section_end = min(len(upper), label.end() + 1200)
        for marker in ("电子客票号码", "发票号码", "INVOICE NO", "INVOICE NUMBER"):
            marker_position = upper.find(marker, label.end(), section_end)
            if marker_position >= 0:
                section_end = marker_position
        amounts = [
            _plausible_amount(match.group(1))
            for match in currency_pattern.finditer(upper[label.end():section_end])
        ]
        amounts = [amount for amount in amounts if amount is not None]
        if not amounts:
            continue
        amount = amounts[-1]
        if amount is None or round(amount, 2) in existing_amounts:
            continue
        if not any(abs(amount - existing) <= 0.01 for existing in results):
            results.append(amount)
    return results


def _is_hotel_folio(page: str) -> bool:
    """Recognise Chinese water bills and English hotel folios."""
    normalised = _normalise_text(page)
    lowered = normalised.lower()
    if any(marker in normalised for marker in (
        "入住水单", "酒店水单", "结账水单", "退房明细", "酒店结算单",
        "宾客账单", "宾客帐单", "酒店账单", "住宿账单",
    )):
        return True
    has_english_dates = bool(
        re.search(r"\b(?:arrival|check[\s-]*in)\b", lowered)
        and re.search(r"\b(?:departure|check[\s-]*out)\b", lowered)
    )
    has_chinese_dates = bool(
        re.search(r"(?:入住|到店|抵店|住)\s*(?:日期|时间)?", normalised)
        and re.search(r"(?:离店|退店|退房|高店)\s*(?:日期|时间)?", normalised)
    )
    has_hotel_fields = (
        "酒店名称" in normalised
        and ("消费金额" in normalised or "房间数" in normalised)
    )
    folio_fields = sum(bool(re.search(pattern, normalised, flags=re.IGNORECASE)) for pattern in (
        r"房型|房号|房间|客房",
        r"入住(?:日期|时间)?|到店(?:日期|时间)?|抵店(?:日期|时间)?|check[ -]*in|arrival",
        r"离店|退房|check[ -]*out|departure",
        r"消费明细|账单明细|结算金额|住宿天数|房价|单价|合计消费|消费合计",
    ))
    has_hotel_context = bool(re.search(
        r"酒店|宾馆|客房|住宿|hotel|inn|plaza|folio",
        normalised,
        flags=re.IGNORECASE,
    ))
    has_folio_layout = has_hotel_context and folio_fields >= 2
    return (
        ("folio" in lowered and has_english_dates)
        or has_chinese_dates
        or has_hotel_fields
        or has_folio_layout
    )


def _ocr_unlabelled_invoice_records(
    page: str,
    gross_matches: list[tuple[float, int, int]],
    filename_category: Optional[str] = None,
) -> list[tuple[float, Optional[str]]]:
    """Recover a total when OCR drops the gross-total label on one of several invoices.

    This conservative fallback runs only when a physical page contains multiple
    invoice-number blocks. For a block with no labelled total, it selects the
    largest currency-marked decimal amount, which is the gross total on Chinese
    VAT invoices when the nearby total caption itself is unreadable.
    """
    upper = page.upper()
    starts = list(re.finditer(
        r"发[^\r\n]{0,2}号码|INVOICE\s*(?:NO\.?|NUMBER)",
        upper,
        flags=re.IGNORECASE,
    ))
    if len(starts) < 2:
        return []
    records = []
    for index, start_match in enumerate(starts):
        block_start = start_match.start()
        block_end = starts[index + 1].start() if index + 1 < len(starts) else len(page)
        if any(block_start <= label_start < block_end for _amount, label_start, _end in gross_matches):
            continue
        block = page[block_start:block_end]
        amounts = [
            amount
            for value in re.findall(
                r"(?:¥|RMB|CNY|(?<![A-Z])Y)\s*:?\s*"
                r"([0-9][0-9,]*(?:\.\s*[0-9]{1,2}))",
                block.upper(),
            )
            if (amount := _plausible_amount(value)) is not None
        ]
        if not amounts:
            continue
        category = _ocr_page_category(block, filename_category)
        if category is None:
            category = _ocr_page_category(page, filename_category)
        records.append((max(amounts), category))
    return records


def _parse_folio_date(value: str, default_year: int) -> Optional[date]:
    """Parse hotel dates in ISO, month-first, and day-first formats."""
    raw_parts = re.findall(r"\d+", value)
    parts = [int(part) for part in raw_parts]
    if len(parts) == 1 and len(raw_parts[0]) == 8:
        compact = raw_parts[0]
        if compact.startswith(("19", "20")):
            parts = [int(compact[:4]), int(compact[4:6]), int(compact[6:8])]
            raw_parts = [compact[:4], compact[4:6], compact[6:8]]
        else:
            # OCR often drops the slashes from a localized hotel date.
            first, second, year_value = int(compact[:2]), int(compact[2:4]), int(compact[4:])
            parts = [first, second, year_value]
            raw_parts = [compact[:2], compact[2:4], compact[4:]]
    if len(parts) == 2:
        year, month, day = default_year, parts[0], parts[1]
    elif len(parts) == 3 and len(raw_parts[0]) == 4:
        year, month, day = parts
    elif len(parts) == 3:
        first, second, short_year = parts
        if short_year < 100:
            century = (default_year // 100) * 100
            year = century + short_year
            if year - default_year > 50:
                year -= 100
            elif default_year - year > 50:
                year += 100
        else:
            year = short_year
        if first > 12 and second <= 12:
            day, month = first, second
        elif second > 12 and first <= 12:
            month, day = first, second
        else:
            # Ambiguous dates retain the established month-first convention.
            month, day = first, second
    else:
        return None
    try:
        return date(year, month, day)
    except ValueError:
        return None


def _hotel_folio_evidence(
    page: str,
    default_year: int,
    source_page: Optional[int] = None,
) -> Optional[AccommodationFolioEvidence]:
    """Extract validation-only dates and totals without booking the folio again."""
    normalised = _normalise_text(page)
    has_arrival_label = bool(re.search(
        r"\u5165\u4f4f(?:\u65e5\u671f|\u65f6\u95f4|\u65e5)?|"
        r"\u5230\u5e97(?:\u65e5\u671f|\u65f6\u95f4|\u65e5)?|"
        r"check[\s-]*in|arrival",
        normalised,
        flags=re.IGNORECASE,
    ))
    has_room_detail = bool(re.search(
        r"\u623f\u578b|\u623f\u53f7|\u623f\u95f4|\u5927\u5e8a|\u5ba2\u623f|"
        r"room\s*(?:type|no\.?)",
        normalised,
        flags=re.IGNORECASE,
    ))
    weak_hotel_folio = has_arrival_label and has_room_detail
    if not _is_hotel_folio(normalised) and not weak_hotel_folio:
        return None

    # Normalise Chinese dates before applying the same labelled-date parser.
    searchable = re.sub(
        r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]?",
        r"\1-\2-\3",
        normalised,
    )
    searchable = re.sub(
        r"(?<!\d)(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]?",
        r"\1-\2",
        searchable,
    )
    date_value = r"(\d{1,4}\s*[-/.]\s*\d{1,2}(?:\s*[-/.]\s*\d{1,4})?|\d{8})"

    def labelled_date(labels: str) -> Optional[date]:
        match = re.search(
            rf"(?:{labels})\s*[:,\uFF1A\uFF0C]?\s*{date_value}",
            searchable,
            flags=re.IGNORECASE,
        )
        return _parse_folio_date(match.group(1), default_year) if match else None

    start = labelled_date(r"arrival|check[\s-]*in|入住(?:日期|时间|日)?|到店(?:日期|时间|日)?|住(?:日期|时间|日)?")
    end = labelled_date(r"departure|check[\s-]*out|离店(?:日期|时间|日)?|退房(?:日期|时间|日)?|退店(?:日期|时间|日)?|高店(?:日期|时间|日)?")

    if start is not None and end is None and weak_hotel_folio:
        arrival_match = re.search(
            r"(?:\u5165\u4f4f(?:\u65e5\u671f|\u65f6\u95f4|\u65e5)?|"
            r"\u5230\u5e97(?:\u65e5\u671f|\u65f6\u95f4|\u65e5)?|"
            r"check[\s-]*in|arrival)",
            searchable,
            flags=re.IGNORECASE,
        )
        if arrival_match:
            nearby = searchable[arrival_match.end():arrival_match.end() + 180]
            for match in re.finditer(date_value, nearby):
                candidate = _parse_folio_date(match.group(1), default_year)
                if candidate is not None and candidate > start:
                    # Sideways OCR often drops the checkout label while
                    # preserving the adjacent arrival and departure values.
                    end = candidate
                    break

    def labelled_amount(labels: str) -> Optional[float]:
        candidates = []
        for match in re.finditer(
            rf"(?:{labels})\s*[:：，,]?\s*(?:¥|￥|RMB|CNY)?\s*"
            r"([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)",
            searchable,
            flags=re.IGNORECASE,
        ):
            amount = _plausible_amount(match.group(1))
            if amount is not None:
                candidates.append(amount)
        return max(candidates) if candidates else None

    total = labelled_amount(
        r"grand\s+total|total|账单金额|消费金额|应付金额|合计消费|消费合计|总金额|总计|应收金额|合计"
    )
    if total is None:
        # A folio can share its page with the matching VAT invoice. The
        # invoice's labelled gross total is then valid folio evidence even if
        # OCR breaks the account's own bottom total across columns.
        gross_totals = {
            round(amount, 2)
            for amount, _start, _end in _ocr_gross_total_matches(normalised)
        }
        if len(gross_totals) == 1:
            total = next(iter(gross_totals))
    unit_rate = labelled_amount(r"room\s+rate|daily\s+rate|房价|单价")
    if total is None and weak_hotel_folio:
        suffix_amounts = [
            amount
            for value in re.findall(
                r"([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)\s*(?:\u5143|yuan)",
                searchable,
                flags=re.IGNORECASE,
            )
            if (amount := _plausible_amount(value)) is not None
        ]
        if suffix_amounts:
            total = max(suffix_amounts)

    # Some low-contrast folios lose the checkout date while keeping both the
    # nightly rate and the folio total. 399 x 3 = 1197 is still decisive
    # evidence of three nights, and is safer than falling back to one night.
    rate_nights = None
    if total is not None and unit_rate and unit_rate > 0:
        candidate_nights = round(total / unit_rate)
        if (
            1 <= candidate_nights <= 60
            and abs(total - unit_rate * candidate_nights) <= max(1.0, abs(total) * 0.01)
        ):
            rate_nights = candidate_nights

    # A posted room charge for each service date is stronger than an OCR'd
    # checkout label when the folio totals those exact dated charges. Hotel
    # scans can misread the checkout day (or the printed folio itself can be
    # internally inconsistent), so require unique consecutive dates and a
    # sum that reconciles to the invoice/folio total before using line items.
    nightly_rows: list[tuple[date, float]] = []
    iso_date_pattern = re.compile(
        r"(?<!\d)(20\d{2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})(?!\d)"
    )
    for raw_line in searchable.splitlines():
        date_match = iso_date_pattern.search(raw_line)
        if date_match is None:
            continue
        try:
            charge_day = date(
                int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3)),
            )
        except ValueError:
            continue
        amount_line = iso_date_pattern.sub(" ", raw_line)
        amount_tokens = re.findall(
            r"(?<!\d)(?:\d{1,3}(?:,\d{3})+|\d+)[.,]\d{2}(?!\d)",
            amount_line,
        )
        if not amount_tokens:
            continue
        token = amount_tokens[-1].replace(" ", "")
        if "," in token and "." not in token and re.fullmatch(r"\d+,\d{2}", token):
            token = token.replace(",", ".")
        charge_amount = _plausible_amount(token)
        if charge_amount is not None:
            nightly_rows.append((charge_day, charge_amount))

    line_item_nights = None
    if total is not None and len(nightly_rows) >= 2:
        row_dates = [row[0] for row in nightly_rows]
        ordered_dates = sorted(set(row_dates))
        line_item_total = round(sum(amount for _day, amount in nightly_rows), 2)
        consecutive = all(
            (later - earlier).days == 1
            for earlier, later in zip(ordered_dates, ordered_dates[1:])
        )
        total_matches_rows = abs(line_item_total - total) <= max(1.0, total * 0.01)
        if len(ordered_dates) == len(nightly_rows) and consecutive and total_matches_rows:
            line_item_nights = len(nightly_rows)
            if unit_rate is None and len({round(amount, 2) for _day, amount in nightly_rows}) == 1:
                unit_rate = nightly_rows[0][1]

    nights_match = re.search(
        r"(?:nights?|晚数|间夜|住宿天数)\s*[:：]?\s*(\d{1,2})(?!\d)",
        searchable,
        flags=re.IGNORECASE,
    )
    nights = int(nights_match.group(1)) if nights_match else None
    nights_source = "LABEL" if nights is not None else "UNCONFIRMED"
    if line_item_nights is not None:
        nights = line_item_nights
        nights_source = "LINE_ITEMS"
        start = min(day for day, _amount in nightly_rows)
        end = max(day for day, _amount in nightly_rows) + timedelta(days=1)
    elif start and end and end > start:
        nights = (end - start).days
        nights_source = "DATES"
    elif rate_nights is not None:
        nights = rate_nights
        nights_source = "RATE"
    if nights is not None and not (1 <= nights <= 60):
        nights = None
        nights_source = "UNCONFIRMED"
    if nights is not None:
        # Complete a half-read stay span from the confirmed night count. This
        # keeps the cap calculation on the folio's real span even when OCR
        # drops one of the two date labels.
        if start is None and end is not None:
            start = end - timedelta(days=nights)
        elif end is None and start is not None:
            end = start + timedelta(days=nights)
        elif start is not None and end is not None and end <= start:
            end = start + timedelta(days=nights)
    return AccommodationFolioEvidence(
        start, end, total, nights, unit_rate, source_page, page[:500], nights_source,
    )


def _hotel_page_city(page: str) -> str:
    """Resolve the hotel city from the property name/address before other text."""
    hotel_markers = re.compile(
        r"酒店|宾馆|旅店|住宿|hotel|folio|marriott|sheraton|crowne\s+plaza|inn|resort",
        re.IGNORECASE,
    )
    mentions: list[tuple[int, str]] = []
    for marker in hotel_markers.finditer(page):
        context = page[max(0, marker.start() - 80):marker.end() + 140]
        city = _city_mentioned_in(context)
        if city:
            mentions.append((marker.start(), city))
    if mentions:
        mentions.sort()
        return mentions[0][1]
    return _city_mentioned_in(page)


def _hotel_invoice_city_span(page: str, default_year: int) -> Optional[tuple[date, date, str]]:
    """Read a hotel's city and stay dates from invoices or booking details."""
    normalised = _normalise_text(page)
    if not re.search(
        r"酒店名称|住宿日期|入住日期|hotel\s*name|check[\s-]*in|arrival|accommodation",
        normalised,
        flags=re.IGNORECASE,
    ):
        return None
    city = _hotel_page_city(normalised)
    if not city:
        return None

    # Use the dates adjacent to the property or stay label. The full page can
    # also contain invoice/print dates that must not be mistaken for the stay.
    markers = re.compile(
        r"酒店名称|住宿日期|入住日期|hotel\s*name|check[\s-]*in|arrival|accommodation",
        re.IGNORECASE,
    )
    for marker in markers.finditer(normalised):
        nearby = normalised[marker.end():marker.end() + 900]
        dates = sorted(set(_date_tokens(nearby, default_year)))
        if len(dates) >= 2:
            first, last = dates[0], dates[1]
            if last >= first:
                return first, last, city
    return None


def _invoice_note_city_span(page: str, default_year: int) -> Optional[tuple[date, date, str]]:
    """Read a dated city annotation from the top OCR lines of an invoice page."""
    lines = [_normalise_text(line) for line in page.splitlines() if line.strip()]
    header_lines = lines[:8]
    for index in range(len(header_lines)):
        window = " ".join(header_lines[index:index + 3])
        dates = _date_tokens(window, default_year)
        cities = _city_mentions_in(window)
        if len(dates) < 2 or len(cities) != 1:
            continue
        if not re.search(
            r"酒店|宾馆|住宿|发票|行程|出差|hotel|folio|invoice|travel",
            window,
            flags=re.IGNORECASE,
        ):
            continue
        start, end = dates[0], dates[1]
        if end < start:
            end = end.replace(year=end.year + 1)
        return start, end, cities[0]
    return None


def _infer_trip_city_timeline(
    pages: list[str],
    accommodation_folios: list[AccommodationFolioEvidence],
    start: date,
    end: date,
    default_year: int,
    fallback_city: str,
) -> dict[date, str]:
    """Assign one evidenced destination to each trip date.

    Hotel stays and dated invoice annotations are stronger than ticket endpoints.
    Flight arrivals fill gaps, while a return flight to the first origin keeps
    the departure city out of the trip's destination timeline.
    """
    priorities: dict[date, int] = {}
    city_by_day: dict[date, str] = {}

    def assign_span(first: date, last: date, city: str, priority: int) -> None:
        if not city:
            return
        if last < start or first > end:
            return
        first = max(first, start)
        last = min(last, end + timedelta(days=1))
        if last <= first:
            last = min(first + timedelta(days=1), end + timedelta(days=1))
        day = first
        while day < last:
            if priority >= priorities.get(day, -1):
                city_by_day[day] = city
                priorities[day] = priority
            day += timedelta(days=1)

    for folio in accommodation_folios:
        if folio.city and folio.start:
            assign_span(folio.start, folio.end or folio.start, folio.city, 4)

    for page in pages:
        hotel_span = _hotel_invoice_city_span(page, default_year)
        if hotel_span:
            assign_span(*hotel_span, priority=3)
        note_span = _invoice_note_city_span(page, default_year)
        if note_span:
            assign_span(*note_span, priority=2)

    route_events = _ticket_route_events("\f".join(pages), default_year)
    home_city = route_events[0]["origin"] if route_events else ""
    for index, route in enumerate(route_events):
        event_day = route["date"]
        if event_day is None:
            event_day = start if index == 0 else end if index == len(route_events) - 1 else None
        if event_day is None or event_day < start or event_day > end:
            continue
        city = route["destination"]
        if event_day == end and home_city and city == home_city:
            city = route["origin"]
        previous_priority = priorities.get(event_day, -1)
        if previous_priority <= 1:
            city_by_day[event_day] = city
            priorities[event_day] = 1

    days = [start + timedelta(days=offset) for offset in range((end - start).days + 1)]
    previous_city = ""
    for day in days:
        if city_by_day.get(day):
            previous_city = city_by_day[day]
        elif previous_city:
            city_by_day[day] = previous_city
    next_city = ""
    for day in reversed(days):
        if city_by_day.get(day):
            next_city = city_by_day[day]
        elif next_city:
            city_by_day[day] = next_city
    if not any(city_by_day.values()) and fallback_city:
        city_by_day = {day: fallback_city for day in days}
    elif fallback_city:
        first_city = next(iter(city_by_day.values()), "")
        if not first_city:
            city_by_day = {day: fallback_city for day in days}
    return city_by_day


def _didi_trip_table_evidence(
    page: str,
    default_year: int,
    source_page: Optional[int] = None,
) -> Optional[DidiTripTableEvidence]:
    """Extract the date and printed total from one Didi trip-table page."""
    normalised = _normalise_text(page)
    upper = normalised.upper()
    if not (re.search(r"DIDI\s*TRAVEL", upper) and "TRIP TABLE" in upper):
        return None

    date_tokens = _date_tokens(normalised, default_year)
    # The application date is normally first; the final two dates are the
    # service range. Prefer the explicit range when OCR keeps the separator.
    range_match = re.search(
        r"(?:行程起止日期|TRIP\s*(?:START|DATE)|DATE)\s*[:：]?\s*"
        r"(20\d{2}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{1,2})\s*"
        r"(?:至|到|[-~4])\s*"
        r"(20\d{2}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{1,2})",
        normalised,
        flags=re.IGNORECASE,
    )
    if range_match:
        start = _parse_folio_date(range_match.group(1), default_year)
        end = _parse_folio_date(range_match.group(2), default_year)
    elif len(date_tokens) >= 2:
        start, end = date_tokens[-2], date_tokens[-1]
    else:
        start = end = date_tokens[-1] if date_tokens else None
    totals = _didi_trip_table_totals(normalised)
    total = totals[-1] if totals else None
    if start is None or end is None or total is None:
        return None
    reviewed_matches = re.findall(
        r"IAV_DIDI_REVIEWED_TOTAL\s+¥\s*"
        r"([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)",
        normalised,
        flags=re.IGNORECASE,
    )
    reviewed_total = next(
        (amount for match in reviewed_matches if (amount := _plausible_amount(match)) is not None),
        None,
    )
    annotations = _didi_annotation_additions(normalised, totals)
    annotated_total = reviewed_total
    if annotated_total is None and annotations:
        annotated_total = round(total + sum(annotations), 2)
    return DidiTripTableEvidence(
        start=start,
        end=end,
        total=total,
        source_page=source_page,
        annotated_total=annotated_total,
        raw_text=page[:500],
        has_separate_taxi_receipt=bool(re.search(
            r"出租(?:车|汽车)|TAXI\s*(?:SPECIAL\s+)?INVOICE|SPECIAL\s+INVOICE|"
            r"IAV_EMBEDDED_TAXI_RECEIPT",
            normalised,
            flags=re.IGNORECASE,
        )),
    )


def _folio_amount_matches(invoice_amount: float, folio_total: Optional[float]) -> bool:
    """Allow minor tax/rounding differences, but never pair unrelated bills."""
    if folio_total is None:
        return False
    tolerance = max(1.0, min(invoice_amount, folio_total) * 0.01)
    return abs(invoice_amount - folio_total) <= tolerance


def _reconcile_accommodation_folios(
    expenses: list[Expense],
    folios: list[AccommodationFolioEvidence],
    trip_start: date,
    trip_end: date,
) -> None:
    """Attach folio dates and nights to matching invoices after amount validation.

    A scanned hotel package commonly separates the tax invoice and its folio
    onto adjacent pages. Amount matching remains the strongest signal; page
    proximity is the fallback when a low-contrast folio has a readable date
    range but no readable total. Invoice-side dates are only candidates until
    reconciled: an invoice can contain its issue date or an incomplete OCR date
    range, so already-populated dates must not prevent folio matching. Folios
    outside the trip range are retained too, allowing the actual dates to raise
    a review warning without losing their confirmed night count.
    """
    invoices = [
        expense for expense in expenses
        if expense.category == "accommodation"
    ]
    usable_folios = [
        folio for folio in folios
        if folio.start is not None
        and folio.end is not None
        and folio.nights is not None
        and folio.start < folio.end
    ]

    def page_distance(invoice: Expense, folio: AccommodationFolioEvidence) -> int:
        if invoice.source_page is None or folio.source_page is None:
            return 10**6
        return abs(invoice.source_page - folio.source_page)

    candidates = sorted(
        (
            abs(invoice.amount - folio.total),
            page_distance(invoice, folio),
            0 if trip_start <= folio.start and folio.end <= trip_end else 1,
            invoice_index,
            folio_index,
        )
        for invoice_index, invoice in enumerate(invoices)
        for folio_index, folio in enumerate(usable_folios)
        if folio.total is not None and _folio_amount_matches(invoice.amount, folio.total)
    )
    def nights_source(folio: AccommodationFolioEvidence) -> str:
        return (
            "HOTEL_FOLIO"
            if folio.nights_source in {"DATES", "LINE_ITEMS", "LABEL"}
            else "HOTEL_FOLIO_RATE"
        )

    used_invoices: set[int] = set()
    used_folios: set[int] = set()
    for _difference, _page_gap, _trip_scope, invoice_index, folio_index in candidates:
        if invoice_index in used_invoices or folio_index in used_folios:
            continue
        invoice = invoices[invoice_index]
        folio = usable_folios[folio_index]
        invoice.start = folio.start
        invoice.end = folio.end
        invoice.accommodation_nights = folio.nights
        invoice.accommodation_folio_total = folio.total
        invoice.service_date_source = "HOTEL_FOLIO"
        invoice.service_date_evidence_page = folio.source_page
        invoice.service_date_evidence_text = folio.raw_text
        invoice.accommodation_nights_source = nights_source(folio)
        used_invoices.add(invoice_index)
        used_folios.add(folio_index)

    proximity_candidates = sorted(
        (
            page_distance(invoice, folio),
            0 if trip_start <= folio.start and folio.end <= trip_end else 1,
            invoice_index,
            folio_index,
        )
        for invoice_index, invoice in enumerate(invoices)
        for folio_index, folio in enumerate(usable_folios)
        if invoice_index not in used_invoices
        and folio_index not in used_folios
        and invoice.source_page is not None
        and folio.source_page is not None
        and folio.total is None
        and page_distance(invoice, folio) <= 2
    )
    for _distance, _trip_scope, invoice_index, folio_index in proximity_candidates:
        if invoice_index in used_invoices or folio_index in used_folios:
            continue
        invoice = invoices[invoice_index]
        folio = usable_folios[folio_index]
        invoice.start = folio.start
        invoice.end = folio.end
        invoice.accommodation_nights = folio.nights
        invoice.accommodation_folio_total = folio.total
        invoice.service_date_source = "HOTEL_FOLIO"
        invoice.service_date_evidence_page = folio.source_page
        invoice.service_date_evidence_text = folio.raw_text
        invoice.accommodation_nights_source = nights_source(folio)
        used_invoices.add(invoice_index)
        used_folios.add(folio_index)


def _has_confirmed_service_date(expense: Expense) -> bool:
    return (
        expense.start is not None
        and expense.service_date_source in CONFIRMED_SERVICE_DATE_SOURCES
    )


def _reconcile_didi_trip_tables(
    expenses: list[Expense],
    tables: list[DidiTripTableEvidence],
    trip_start: date,
    trip_end: date,
) -> None:
    """Use Didi headers to date nearby invoices and recover missing summaries."""
    taxi_expenses = [expense for expense in expenses if expense.category == "taxi"]
    used_invoices: set[int] = set()
    for table in tables:
        if table.start is None or table.end is None or table.total is None:
            continue

        def date_compatible(expense: Expense) -> bool:
            if not _has_confirmed_service_date(expense):
                return True
            expense_end = expense.end or expense.start
            return expense.start <= table.end and expense_end >= table.start

        amount_matches = sorted(
            (
                abs(expense.source_page - table.source_page),
                index,
                expense,
            )
            for index, expense in enumerate(taxi_expenses)
            if index not in used_invoices
            and expense.source_page is not None
            and table.source_page is not None
            and abs(expense.amount - table.total) <= 0.01
        )
        # A nearby, date-compatible itinerary may fill a missing service date.
        # A unique amount several pages away is not enough to override the PDF
        # sequence or a date printed on the invoice.
        candidates = [
            item for item in amount_matches
            if item[0] <= 1 and date_compatible(item[2])
        ]
        if candidates:
            _distance, index, invoice = candidates[0]
            used_invoices.add(index)
            if not _has_confirmed_service_date(invoice):
                invoice.start = table.start
                invoice.end = table.end
                invoice.service_date_source = "DIDI_ITINERARY"
                invoice.service_date_evidence_page = table.source_page
                invoice.service_date_evidence_text = table.raw_text
            invoice.description = (
                f"{invoice.description}；Didi 行程单日期 {table.start:%Y-%m-%d}"
            )
        elif not table.has_separate_taxi_receipt and not amount_matches:
            # The printed header is the authoritative fallback when the
            # neighbouring electronic invoice page was not OCR-readable, but
            # a page with a separate taxi receipt remains supporting evidence
            # until that receipt's own amount has been extracted.
            if not any(
                abs(expense.amount - table.total) <= 0.01
                and expense.start == table.start
                and expense.end == table.end
                for expense in taxi_expenses
            ):
                recovered = Expense(
                    category="taxi",
                    amount=table.total,
                    start=table.start,
                    end=table.end,
                    description=(
                        f"Didi 行程单合计（第 {table.source_page} 页）: "
                        f"¥{table.total:.2f}"
                    ),
                    source_file="",
                    aggregate=True,
                    source_page=table.source_page,
                    service_date_source="DIDI_ITINERARY",
                    service_date_evidence_page=table.source_page,
                    service_date_evidence_text=table.raw_text,
                )
                expenses.append(recovered)
                taxi_expenses.append(recovered)

        page_embedded = [
            expense for expense in expenses
            if expense.category == "taxi"
            and expense.source_page == table.source_page
            and "embedded taxi receipt" in expense.description.lower()
        ]
        if not page_embedded:
            continue
        if table.annotated_total is not None:
            group = f"didi-annotated-{table.source_page}"
            group_members = [
                expense for expense in taxi_expenses
                if expense.source_page in {table.source_page, (table.source_page or 0) - 1}
                and (expense in page_embedded or abs(expense.amount - table.total) <= 0.01)
                and date_compatible(expense)
            ]
            confirmed_dates = {
                expense.start
                for expense in group_members
                if expense.start is not None
                and expense.service_date_source in CONFIRMED_SERVICE_DATE_SOURCES
            }
            if len(confirmed_dates) <= 1:
                group_total = round(sum(expense.amount for expense in group_members), 2)
                use_written_total = abs(group_total - table.annotated_total) <= 0.01
                for expense in group_members:
                    expense.allocation_group = group
                    if use_written_total:
                        expense.allocation_group_total = table.annotated_total
                    if not _has_confirmed_service_date(expense):
                        expense.allocation_anchor = table.start
        else:
            # Embedded taxi receipts are dated by the service date printed in
            # the Didi header; this prevents two same-day receipts from being
            # spread over later days merely because OCR missed their dates.
            for expense in page_embedded:
                if not _has_confirmed_service_date(expense):
                    expense.start = table.start
                    expense.end = table.end
                    expense.service_date_source = "DIDI_ITINERARY"
                    expense.service_date_evidence_page = table.source_page
                    expense.service_date_evidence_text = table.raw_text


def _infer_accommodation_spans(trip: Trip) -> None:
    """Fill missing hotel spans from neighbouring folios and trip boundaries."""
    accommodations = [
        expense for expense in trip.expenses
        if expense.category == "accommodation"
    ]
    if not accommodations:
        return
    ordered = sorted(
        enumerate(accommodations),
        key=lambda item: (item[1].source_page is None, item[1].source_page or item[0]),
    )
    known = [expense for _index, expense in ordered if expense.start and expense.end]
    if not known:
        return
    for index, expense in ordered:
        if expense.start is not None or expense.end is not None:
            continue
        previous = next(
            (candidate for _candidate_index, candidate in reversed(ordered[:index])
             if candidate.start and candidate.end),
            None,
        )
        following = next(
            (candidate for _candidate_index, candidate in ordered[index + 1:]
             if candidate.start and candidate.end),
            None,
        )
        if previous and following and previous.end <= following.start:
            expense.start, expense.end = previous.end, following.start
            inferred_source = "ADJACENT_RECORD_INFERENCE"
        elif previous and previous.end < trip.end:
            expense.start, expense.end = previous.end, trip.end
            inferred_source = "TRIP_BOUNDARY_INFERENCE"
        elif following and trip.start < following.start:
            expense.start, expense.end = trip.start, following.start
            inferred_source = "TRIP_BOUNDARY_INFERENCE"
        else:
            inferred_source = ""
        if expense.start and expense.end and expense.end > expense.start and inferred_source:
            expense.accommodation_nights = (expense.end - expense.start).days
            expense.service_date_source = inferred_source
            expense.accommodation_nights_source = inferred_source


def _ocr_page_invoice_records(
    page: str,
    filename_category: Optional[str] = None,
) -> tuple[list[tuple[float, Optional[str]]], bool]:
    """Extract one or more invoice totals from a scanned page.

    Scanner packs sometimes place several invoices and a supporting Didi trip
    table on the same physical page. Treating the whole page as one supporting
    document used to discard every invoice on that page.
    """
    normalised = _normalise_text(page)
    upper = normalised.upper()
    is_supporting_trip_table = bool(
        re.search(r"DIDI\s*TRAVEL", upper) and "TRIP TABLE" in upper
    )
    if is_supporting_trip_table and "IAV_OUT_OF_SCOPE_TAXI_RECEIPT" in upper:
        return [], True
    didi_table_totals = (
        _didi_trip_table_totals(normalised) if is_supporting_trip_table else []
    )
    # Reviewer arithmetic above a Didi table is useful reconciliation
    # evidence, but neither that handwritten total nor the printed table sum
    # is a separate invoice. Keep both out of the payable-total matcher.
    payable_text = re.sub(
        r"IAV_DIDI_REVIEWED_TOTAL\s+¥\s*"
        r"[0-9][0-9,]*(?:\.\s*[0-9]{1,2})?",
        "",
        normalised,
        flags=re.IGNORECASE,
    )
    # Only a total printed inside the Didi table is supporting evidence.  A
    # separate invoice on the same scanned page can legitimately carry the same
    # amount as the table sum (for example a page containing both the invoice
    # and the trip table it reconciles to), so do not discard it by amount
    # alone.
    trip_table_start = None
    if is_supporting_trip_table:
        table_heading = re.search(r"DIDI\s*TRAVEL", payable_text.upper())
        trip_table_start = table_heading.start() if table_heading else None
    gross_matches = [
        match for match in _ocr_gross_total_matches(payable_text)
        if not is_supporting_trip_table
        or trip_table_start is None
        or match[1] < trip_table_start
        or not any(abs(match[0] - total) <= 0.01 for total in didi_table_totals)
    ]
    ticket_total_amounts = _ocr_ticket_total_matches(normalised, gross_matches)
    embedded_amounts = [
        amount
        for match in re.findall(
            r"IAV_EMBEDDED_TAXI_RECEIPT\s+¥\s*([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)",
            upper,
        )
        if (amount := _plausible_amount(match)) is not None
    ]
    toll_amounts = [
        amount
        for match in re.findall(
            r"IAV_TOLL_INVOICE_AMOUNT\s+¥\s*"
            r"([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)",
            upper,
        )
        if (amount := _plausible_amount(match)) is not None
    ]
    is_order_detail = "订单详情" in normalised and "发票号码" not in normalised
    is_supporting_hotel_folio = (
        _hotel_folio_evidence(normalised, _document_year(normalised, None)) is not None
    )
    is_invoice_detail_page = (
        "发票代码" in normalised
        and "发票号码" in normalised
        and not gross_matches
    )
    # Some valid physical highway invoices print the payable amount as
    # “通行费：131元” instead of using the standard “价税合计” label.  Treat
    # that labelled amount as payable evidence; an unlabelled OCR currency
    # fragment on an invoice-detail page remains supporting-only evidence.
    has_labelled_invoice_amount = bool(re.search(
        r"(?:通行费|金额|合计|总额)\s*[:：]?\s*"
        r"(?:¥|￥)?\s*[0-9][0-9,]*(?:\.\s*[0-9]{1,2})?\s*(?:元|YUAN)\b",
        normalised,
        flags=re.IGNORECASE,
    ))
    detail_currency_matches = re.findall(
        r"(?:¥|￥|RMB|CNY|(?<![A-Z])Y)\s*:?\s*"
        r"([0-9]{2,}[0-9,]*(?:\.\s*[0-9]{1,2})?)",
        upper,
    )
    detail_currency_amounts = [
        amount for match in detail_currency_matches
        if (amount := _plausible_amount(match)) is not None
    ]
    has_clear_detail_currency_amount = bool(detail_currency_amounts)
    has_invoice_evidence = bool(embedded_amounts) or bool(toll_amounts) or bool(gross_matches) or any(
        marker in normalised for marker in ("发票号码", "电子发票", "数电发票")
    )
    if (
        (is_order_detail and not has_invoice_evidence)
        or (is_supporting_trip_table and not has_invoice_evidence)
        or (is_supporting_hotel_folio and not has_invoice_evidence)
        or (
            is_invoice_detail_page
            and not embedded_amounts
            and not has_labelled_invoice_amount
            and not has_clear_detail_currency_amount
        )
    ):
        return [], True

    if gross_matches:
        records: list[tuple[float, Optional[str]]] = []
        previous_end = 0
        whole_page_category = _ocr_page_category(normalised, filename_category)
        for amount, label_start, match_end in gross_matches:
            context_start = previous_end if previous_end else 0
            # Keep invoice headings with their total, but cap an unusually long
            # OCR section so unrelated pages or another invoice cannot decide
            # this receipt's category.
            if label_start - context_start > 1800:
                context_start = max(context_start, label_start - 1800)
            context = normalised[context_start:match_end]
            category = _ocr_page_category(context, filename_category)
            if category is None:
                category = whole_page_category
            records.append((amount, category))
            previous_end = match_end
        records.extend(
            _ocr_unlabelled_invoice_records(
                normalised, gross_matches, filename_category,
            )
        )
        records.extend((amount, "taxi") for amount in embedded_amounts)
        records.extend((amount, "taxi") for amount in toll_amounts)
        records.extend((amount, "ticket") for amount in ticket_total_amounts)
        return records, False

    if embedded_amounts:
        records = [(amount, "taxi") for amount in embedded_amounts]
        records.extend((amount, "ticket") for amount in ticket_total_amounts)
        return records, False

    if toll_amounts:
        records = [(amount, "taxi") for amount in toll_amounts]
        records.extend((amount, "ticket") for amount in ticket_total_amounts)
        return records, False

    if ticket_total_amounts:
        return [(amount, "ticket") for amount in ticket_total_amounts], False

    if is_invoice_detail_page and detail_currency_amounts:
        # Physical highway invoices may contain both the payable passage fee
        # and smaller OCR-readable figures such as tax/metadata values. The
        # main currency-marked amount is the largest clear amount on this
        # single-invoice page; do not select the last OCR fragment.
        return [(max(detail_currency_amounts), _ocr_page_category(normalised, filename_category))], False

    matches = re.findall(
        r"(?:¥|RMB|CNY|(?<![A-Z])Y)\s*:?\s*([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)",
        upper,
    )
    amounts = [amount for match in matches if (amount := _plausible_amount(match)) is not None]
    if not amounts:
        # Faint hotel folios often retain the amount but turn the yuan glyph
        # into the correctly recognised suffix "元".
        suffixes = re.findall(
            r"([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)\s*(?:元|YUAN)\b",
            normalised,
            flags=re.IGNORECASE,
        )
        amounts = [amount for match in suffixes if (amount := _plausible_amount(match)) is not None]
    if not amounts:
        return [], False
    return [(amounts[-1], _ocr_page_category(normalised, filename_category))], False


def _ocr_page_invoice_total(page: str) -> tuple[Optional[float], bool]:
    """Return the last gross total on one OCR invoice page (compatibility API)."""
    records, is_supporting = _ocr_page_invoice_records(page)
    return (records[-1][0] if records else None), is_supporting


def _ocr_page_category(page: str, filename_category: Optional[str] = None) -> Optional[str]:
    """Classify a page only when its invoice type has unambiguous evidence.

    Filenames are a fallback. OCR labels on the actual invoice take precedence,
    and conflicting type cues remain unclassified for review instead of being
    forced into a neighbouring expense category.
    """
    lowered = _normalise_text(page).lower()
    hotel_markers = (
            "住宿费", "住宿服务", "酒店住宿", "酒店发票", "酒店管理", "消店管理",
            "宾馆", "客房", "hotel", "accommodation", "lodging",
    )
    taxi_markers = (
            "出租车", "出租汽车", "出租车发票", "网约车", "打车", "的士",
            "taxi", "uber", "ride hailing",
    )
    ticket_markers = (
            "铁路电子客票", "电子客票", "火车票", "机票", "航班", "航空运输",
            "train ticket", "railway", "flight", "airline",
    )
    special_markers = (
            "高速公路", "通行费", "过路费", "停车费", "停车发票", "油费", "加油站",
            "toll", "parking",
    )
    has_hotel = any(marker in lowered for marker in hotel_markers)
    has_taxi = any(marker in lowered for marker in taxi_markers)
    has_ticket = any(marker in lowered for marker in ticket_markers)
    has_special_transport = any(marker in lowered for marker in special_markers)
    if re.search(
        r"滴.{0,3}(?:出行|遇行|满行)|出行科技有限公司|滴滴.{0,8}科技有限公司",
        lowered,
    ):
        has_taxi = True

    # Fuel surcharge and toll wording are ancillary labels on some ticket
    # pages. They do not override an explicit ticket or hotel document type.
    if has_ticket and not has_hotel and not has_taxi:
        return "ticket"
    if has_hotel and not has_ticket and not has_taxi:
        return "accommodation"
    if has_taxi and not has_ticket and not has_hotel:
        return "taxi"
    if sum((has_ticket, has_hotel, has_taxi)) > 1:
        return None
    if has_special_transport:
        return "taxi"

    if re.search(
        r"\b(?:bike\s+(?:share|sharing|rental|rent|trip)|"
        r"(?:shared|rental)\s+bike|ride\s+hailing)\b",
        lowered,
    ):
        return "taxi"
    weak_markers = {
        "accommodation": ("酒店", "住宿", "hotel", "lodging"),
        "taxi": (
            "滴滴", "出行", "地铁", "单车", "自行车", "骑行", "租赁服务",
            "共享单车", "单车租赁", "自行车租赁", "共享电单车",
        ),
        "ticket": ("票价", "铁路", "火车", "高铁", "动车", "车票", "客运"),
        "special": tuple(keyword.lower() for keyword in SPECIAL_KEYWORDS),
    }
    weak_categories = {
        category for category, markers in weak_markers.items()
        if any(marker in lowered for marker in markers)
    }
    if len(weak_categories) == 1:
        return _normalise_expense_category(next(iter(weak_categories)))
    if filename_category and not weak_categories:
        return _normalise_expense_category(filename_category)
    return None


def _looks_like_route(line: str) -> bool:
    return bool(re.search(r"[\u4e00-\u9fffA-Za-z].*[-—－].*[\u4e00-\u9fffA-Za-z]", line))


def _is_date_only_line(line: str) -> bool:
    """Whether a paragraph is a date heading for the following expense."""
    remaining = FULL_DATE_RE.sub("", line)
    remaining = COMPACT_SPAN_RE.sub("", remaining)
    remaining = DATE_RE.sub("", remaining)
    remaining = CHINESE_DATE_RE.sub("", remaining)
    remaining = re.sub(r"[\s,，:：;；()（）\[\]【】-]", "", remaining)
    return not remaining


def _equation_components(line: str) -> list[float]:
    """Return individual invoice amounts from an expression such as 930+49=979."""
    normalised = _normalise_text(line).replace("＋", "+")
    if "=" not in normalised or "+" not in normalised.split("=", 1)[0]:
        return []
    left = normalised.split("=", 1)[0]
    left = FULL_DATE_RE.sub("", left)
    left = COMPACT_SPAN_RE.sub("", left)
    left = CHINESE_DATE_RE.sub("", left)

    def remove_short_date(match):
        month, day = int(match.group(1)), int(match.group(2))
        return "" if month <= 12 and day <= 31 else match.group(0)

    left = DATE_RE.sub(remove_short_date, left)
    values = []
    for value in NUMBER_RE.findall(left):
        amount = _plausible_amount(value)
        if amount is not None:
            values.append(amount)
    return values if len(values) >= 2 else []


def _validated_equation_total(line: str) -> tuple[float, list[float]] | None:
    """Return a document-written total only when its operands reconcile."""
    components = _equation_components(line)
    total = _amount_from_line(line)
    if not components or total is None:
        return None
    if abs(round(sum(components), 2) - round(total, 2)) > 0.01:
        return None
    return round(total, 2), components


def _dated_equation_evidence(page_text: str, default_year: int) -> list[dict]:
    """Read exact, date-headed equations already written in a PDF page."""
    evidence: list[dict] = []
    pending_date: Optional[date] = None
    pending_date_text = ""
    for raw_line in page_text.splitlines():
        line = _normalise_text(raw_line)
        if not line:
            continue
        line_start, line_end = _date_span(line, default_year)
        if line_start and _is_date_only_line(line):
            pending_date = line_start
            pending_date_text = line
            continue
        equation = _validated_equation_total(line)
        service_day = line_start or pending_date
        if equation is None or service_day is None:
            if line_start:
                pending_date = line_start
                pending_date_text = line
            continue
        evidence.append({
            "date": service_day,
            "total": equation[0],
            "components": equation[1],
            "text": line,
            "date_text": line if line_start else pending_date_text,
        })
        pending_date = None
        pending_date_text = ""
    return evidence


def _taxi_equation_evidence(page_text: str, default_year: int) -> list[dict]:
    """Read validated taxi equations, whether or not a date is written."""
    evidence: list[dict] = []
    pending_date: Optional[date] = None
    pending_date_text = ""
    in_taxi_section = bool(
        _looks_like_taxi_invoice(page_text)
        or re.search(r"滴滴|DIDI", page_text, flags=re.IGNORECASE)
    )
    for raw_line in page_text.splitlines():
        line = _normalise_text(raw_line)
        if not line:
            continue
        if re.search(r"机票|飞机|航空|住宿|酒店|补贴|招待|餐费", line):
            in_taxi_section = False
        if re.search(r"出租|打车|的士|网约车|TAXI", line, flags=re.IGNORECASE):
            in_taxi_section = True
        line_start, _line_end = _date_span(line, default_year)
        if line_start and _is_date_only_line(line):
            pending_date = line_start
            pending_date_text = line
            continue
        equation = _validated_equation_total(line) if in_taxi_section else None
        if equation is None:
            if line_start:
                pending_date = line_start
                pending_date_text = line
            elif line and not _is_date_only_line(line):
                pending_date = None
                pending_date_text = ""
            continue
        service_day = line_start or pending_date
        date_text = line if line_start else pending_date_text
        evidence.append({
            "date": service_day,
            "total": equation[0],
            "components": equation[1],
            "text": line,
            "date_text": date_text,
        })
        pending_date = None
        pending_date_text = ""
    return evidence


def _reconcile_dated_taxi_equations(
    expenses: list[Expense],
    pages: list[str],
    default_year: int,
) -> None:
    """Use exact written taxi subtotals without merging distant PDF pages.

    Receipt operands may continue for one or two pages after a handwritten
    subtotal. Matched receipts retain individual trace data, while the cell
    uses the PDF-written total. Missing operands stay as coverage candidates;
    they are never fabricated as invoice records.
    """
    used_expenses: set[int] = set()
    page_equations = {
        page_number: _taxi_equation_evidence(page, default_year)
        for page_number, page in enumerate(pages, 1)
    }
    nested_totals_by_page = {
        page_number: [
            {
                round(other["total"], 2)
                for other_index, other in enumerate(equations)
                if other_index != equation_index
            }
            for equation_index, _equation in enumerate(equations)
        ]
        for page_number, equations in page_equations.items()
    }
    for page_number, page in enumerate(pages, 1):
        equations = [
            (equation_index + 1, equation)
            for equation_index, equation in enumerate(page_equations.get(page_number, []))
            if not any(
                round(component, 2)
                in nested_totals_by_page.get(page_number, [set()])[equation_index]
                for component in equation["components"]
            )
        ]
        for equation_index, equation in equations:
            key = (
                page_number, equation["date"], equation["total"],
                tuple(round(value, 2) for value in equation["components"]),
            )
            selected: list[int] = []
            missing_component = False
            max_page_gap = 1 if equation["date"] is not None else 2
            for amount in equation["components"]:
                candidates = [
                    index for index, expense in enumerate(expenses)
                    if index not in used_expenses
                    and expense.category == "taxi"
                    and expense.source_page is not None
                    and abs(expense.source_page - page_number) <= max_page_gap
                    and abs(expense.amount - amount) <= 0.01
                    and (
                        equation["date"] is None
                        or not _has_confirmed_service_date(expense)
                        or expense.start == equation["date"]
                    )
                ]
                if not candidates:
                    missing_component = True
                    continue
                chosen = min(
                    candidates,
                    key=lambda index: (
                        abs((expenses[index].source_page or page_number) - page_number),
                        expenses[index].source_page or 0,
                        index,
                    ),
                )
                selected.append(chosen)
                used_expenses.add(chosen)
            if len(selected) < 2:
                used_expenses.difference_update(selected)
                continue
            members = [expenses[index] for index in selected]
            confirmed_dates = {
                member.start for member in members
                if _has_confirmed_service_date(member)
            }
            if len(confirmed_dates) > 1:
                used_expenses.difference_update(selected)
                continue
            linked_date = equation["date"] or next(iter(confirmed_dates), None)
            if linked_date is not None and any(
                member.allocation_anchor is not None
                and member.allocation_anchor != linked_date
                for member in members
            ):
                used_expenses.difference_update(selected)
                continue
            all_components_matched = (
                not missing_component
                and len(selected) == len(equation["components"])
                and abs(round(sum(member.amount for member in members), 2) - equation["total"]) <= 0.01
            )
            if round(sum(member.amount for member in members), 2) > equation["total"] + 0.01:
                used_expenses.difference_update(selected)
                continue

            group = f"pdf-dated-taxi-{page_number}-{equation_index}"
            for member in members:
                if linked_date is not None and not _has_confirmed_service_date(member):
                    member.start = linked_date
                    member.end = linked_date
                    member.service_date_source = (
                        "SOURCE_TEXT"
                        if equation["date"] is not None
                        and equation["date_text"] == equation["text"]
                        else "PAGE_DATE_HEADING"
                    )
                    dated_member = next((
                        peer for peer in members
                        if _has_confirmed_service_date(peer)
                        and peer.start == linked_date
                    ), None)
                    member.service_date_evidence_page = (
                        page_number if equation["date"] is not None
                        else dated_member.service_date_evidence_page if dated_member
                        else page_number
                    )
                    member.service_date_evidence_text = (
                        f"{equation['date_text']} {equation['text']}"
                        if equation["date"] is not None
                        else dated_member.service_date_evidence_text if dated_member
                        else equation["text"]
                    ).strip()
                member.allocation_group = group
                member.allocation_group_total = equation["total"]
                member.allocation_group_uses_written_total = not all_components_matched


def _written_ticket_summary_candidates(
    pages: list[str], expenses: list[Expense],
) -> dict[int, list[float]]:
    """Find ticket fare operands written in the PDF but missing as receipts."""
    ticket_expenses = [
        expense for expense in expenses
        if expense.category == "ticket" and not expense.inferred_from_document_summary
    ]
    candidates_by_page: dict[int, list[float]] = defaultdict(list)
    for page_number, page in enumerate(pages, 1):
        lines = page.splitlines()
        has_strong_ticket_context = bool(re.search(
            r"机票|飞机|航空运输|电子客票|航班号|乘机人|BOARDING\s*PASS|E[ -]?TICKET|FLIGHT",
            page,
            flags=re.IGNORECASE,
        ))
        has_ticket_context = has_strong_ticket_context or bool(
            re.search(r"行程单", page)
            and re.search(r"旅客姓名|有效身份证件号码|航空公司", page)
            and any(_looks_like_route(line) for line in lines)
            and not re.search(r"出租车|出租汽车|TAXI", page, flags=re.IGNORECASE)
        )
        for line in lines:
            if not re.search(r"机票|飞机|航空|FLIGHT|TICKET", line, flags=re.IGNORECASE) and not has_ticket_context:
                continue
            equation = _validated_equation_total(line)
            if equation is None:
                continue
            total, components = equation
            if any(abs(expense.amount - total) <= 0.01 for expense in ticket_expenses):
                continue
            unmatched = list(components)
            for expense in ticket_expenses:
                match_index = next(
                    (index for index, amount in enumerate(unmatched)
                     if abs(amount - expense.amount) <= 0.01),
                    None,
                )
                if match_index is not None:
                    unmatched.pop(match_index)
            candidates_by_page[page_number].extend(
                round(amount, 2) for amount in unmatched
            )
    return candidates_by_page


def _infer_written_ticket_summary_expenses(
    pages: list[str], expenses: list[Expense], source_file: str,
) -> list[Expense]:
    """Count likely ticket fares printed in a flight summary, with review risk."""
    inferred: list[Expense] = []
    for page_number, amounts in _written_ticket_summary_candidates(pages, expenses).items():
        page = pages[page_number - 1]
        route_lines = [
            _normalise_text(line)
            for line in page.splitlines()
            if _looks_like_route(_normalise_text(line))
        ]
        route_context = "；".join(route_lines[:3])
        for amount in amounts:
            if any(
                expense.category == "ticket"
                and abs(expense.amount - amount) <= 0.01
                for expense in [*expenses, *inferred]
            ):
                continue
            description = f"PDF 机票/航段汇总推定票价 ¥{amount:.2f}；原始票面金额未独立识别"
            if route_context:
                description += f"；航段线索：{route_context[:400]}"
            inferred.append(Expense(
                category="ticket",
                amount=round(amount, 2),
                start=None,
                end=None,
                description=description,
                source_file=source_file,
                source_page=page_number,
                inferred_from_document_summary=True,
            ))
    return inferred


def _written_taxi_summary_totals(pages: list[str]) -> dict[int, float]:
    """Read a page-level taxi summary equation, excluding dated day details."""
    summaries: dict[int, float] = {}
    for page_number, page in enumerate(pages, 1):
        active_taxi_section = False
        pending_date_heading = False
        equations = []
        for raw_line in page.splitlines():
            line = _normalise_text(raw_line)
            line_start, _line_end = _date_span(line, 2000)
            if line_start is not None and _is_date_only_line(line):
                pending_date_heading = True
                continue
            if re.search(r"机票|飞机|住宿|酒店|补贴|招待|餐费", line):
                active_taxi_section = False
            if re.search(r"出租|打车|的士|网约车|TAXI", line, flags=re.IGNORECASE):
                active_taxi_section = True
            equation = _validated_equation_total(line) if active_taxi_section else None
            if equation is not None:
                if len(equation[1]) >= 3 and not pending_date_heading:
                    equations.append(equation)
                pending_date_heading = False
            elif line:
                pending_date_heading = False
        if not equations:
            continue
        nested_totals = {
            round(value, 2)
            for _total, components in equations
            for value in components
        }
        outer_equations = [
            equation for equation in equations
            if round(equation[0], 2) not in nested_totals
        ]
        chosen = max(outer_equations or equations, key=lambda item: len(item[1]))
        summaries[page_number] = chosen[0]
    return summaries


def _taxi_equation_missing_components(
    pages: list[str], expenses: list[Expense], default_year: int,
) -> dict[int, list[float]]:
    """List equation operands absent from the receipts grouped to that total."""
    missing_by_page: dict[int, list[float]] = defaultdict(list)
    for page_number, page in enumerate(pages, 1):
        equations = _taxi_equation_evidence(page, default_year)
        nested_totals = [
            {
                round(other["total"], 2)
                for other_index, other in enumerate(equations)
                if other_index != equation_index
            }
            for equation_index, _equation in enumerate(equations)
        ]
        for equation_index, equation in enumerate(equations):
            if any(
                round(component, 2) in nested_totals[equation_index]
                for component in equation["components"]
            ):
                continue
            group_name = f"pdf-dated-taxi-{page_number}-{equation_index + 1}"
            members = [
                expense for expense in expenses
                if expense.category == "taxi"
                and expense.allocation_group == group_name
            ]
            if not members:
                continue
            remaining = list(equation["components"])
            for member in members:
                match_index = next(
                    (index for index, amount in enumerate(remaining)
                     if abs(amount - member.amount) <= 0.01),
                    None,
                )
                if match_index is not None:
                    remaining.pop(match_index)
            missing_by_page[page_number].extend(
                round(amount, 2) for amount in remaining
            )
    return missing_by_page


def _recognized_taxi_total(expenses: list[Expense]) -> float:
    """Count a written group total once, including its visibly missing operands."""
    total = 0.0
    seen_groups: set[str] = set()
    for expense in expenses:
        if expense.category != "taxi":
            continue
        group = expense.allocation_group
        if not group:
            total += expense.amount
            continue
        if group in seen_groups:
            continue
        seen_groups.add(group)
        members = [
            member for member in expenses
            if member.category == "taxi" and member.allocation_group == group
        ]
        totals = {
            round(float(member.allocation_group_total), 2)
            for member in members
            if member.allocation_group_total is not None
        }
        component_sum = round(sum(member.amount for member in members), 2)
        if len(totals) == 1:
            declared = next(iter(totals))
            if (
                abs(declared - component_sum) <= 0.01
                or any(member.allocation_group_uses_written_total for member in members)
            ) and declared >= component_sum - 0.01:
                total += declared
                continue
        total += component_sum
    return round(total, 2)


def _single_page_date_heading(page_text: str, default_year: int) -> tuple[date, str] | None:
    """Return a page's unambiguous standalone service-date heading."""
    headings: list[tuple[date, str]] = []
    for raw_line in page_text.splitlines():
        line = _normalise_text(raw_line)
        start, _end = _date_span(line, default_year)
        if start is not None and _is_date_only_line(line):
            headings.append((start, line))
    unique_dates = {day for day, _text in headings}
    if len(unique_dates) != 1 or not headings:
        return None
    return headings[0]


def _deduplicate_aggregate_expenses(expenses: list[Expense]) -> list[Expense]:
    """Remove section totals when the dated detail lines already add to them.

    DOCX travel packs commonly contain both a section total and its dated
    breakdown. Keeping both doubles the amount. If only undated expression
    totals exist, keep the largest one (the overall total) and discard the
    smaller route breakdowns.
    """
    result: list[Expense] = []
    for category in {expense.category for expense in expenses}:
        group = [expense for expense in expenses if expense.category == category]
        aggregates = [expense for expense in group if expense.aggregate]
        details = [expense for expense in group if not expense.aggregate]
        if not aggregates:
            result.extend(group)
            continue

        ordinary_details = [expense for expense in details if "发票" not in expense.description]
        largest_total = max(expense.amount for expense in aggregates)
        ordinary_total = round(sum(expense.amount for expense in ordinary_details), 2)
        if ordinary_details and abs(ordinary_total - largest_total) <= 0.01:
            # A line like “182元发票：机场-地铁站” can repeat a component
            # already included in the dated expression immediately above.
            result.extend(ordinary_details)
            continue

        if not details:
            result.append(max(aggregates, key=lambda expense: expense.amount))
            continue

        # If details do not reconcile to the section total, preserve both so
        # the reviewer can see the discrepancy instead of silently dropping it.
        result.extend(group)
    return result


def _is_toll_summary_page(page: str) -> bool:
    """Recognise a toll-invoice summary sheet, which repeats its invoices."""
    normalised = _normalise_text(page)
    return (
        "收费公路通行费电子票据汇总单" in normalised
        or (
            "电子票据汇总单" in normalised
            and ("票据索引" in normalised or "共1张" in normalised or "共 1 张" in normalised)
        )
    )


def _deduplicate_summary_sheet_expenses(
    expenses: list[Expense],
    pages: list[str],
) -> list[Expense]:
    """Drop a summary sheet's amount when its source invoice is also present.

    Toll-invoice summary sheets restate the amount of the invoices they index.
    Scanning both would claim the same toll twice. Keep the summary only when
    it is the sole evidence for that amount.
    """
    summary_pages = {
        index + 1 for index, page in enumerate(pages) if _is_toll_summary_page(page)
    }
    if not summary_pages:
        return expenses
    return [
        expense for expense in expenses
        if expense.source_page not in summary_pages
        or not any(
            other is not expense
            and other.category == expense.category
            and other.source_page not in summary_pages
            and abs(other.amount - expense.amount) <= 0.01
            for other in expenses
        )
    ]


def _deduplicate_repeated_invoice_notes(expenses: list[Expense]) -> list[Expense]:
    """Drop a note that repeats an invoice already present in a dated equation."""
    result: list[Expense] = []
    for index, expense in enumerate(expenses):
        description = expense.description.lower()
        is_invoice_note = "发票" in expense.description or "invoice" in description
        if is_invoice_note and any(
            other.category == expense.category
            and abs(other.amount - expense.amount) <= 0.01
            and other.start == expense.start
            and not ("发票" in other.description or "invoice" in other.description.lower())
            for other in expenses[index + 1:]
        ):
            continue
        if is_invoice_note and any(
            other.category == expense.category
            and abs(other.amount - expense.amount) <= 0.01
            and other.start == expense.start
            and not ("发票" in other.description or "invoice" in other.description.lower())
            for other in result
        ):
            continue
        result.append(expense)
    return result


def _page_coverage_records(
    pages: list[str],
    expenses: list[Expense],
    invoices: list[TravelInvoice],
    warnings: list[str],
    *,
    is_pdf: bool,
    trip_city: Optional[str] = None,
) -> list[dict]:
    """Report how each extracted source page maps to payable records.

    This is a coverage signal, not a claim that OCR is complete. Pages with
    invoice-like evidence but no Expense, failed OCR, or a mismatch between
    Expense and structured projections are explicitly marked for review.
    """
    failed_pages = {
        int(match.group(1))
        for warning in warnings
        for match in re.finditer(r"第\s*(\d+)\s*页\s*[:：]", warning)
    }
    failed_pages.update(
        int(page_number)
        for warning in warnings
        for match in re.finditer(r"OCR 未成功页码\s*[:：]\s*([0-9,，、 ]+)", warning)
        for page_number in re.findall(r"\d+", match.group(1))
    )
    written_ticket_candidates = _written_ticket_summary_candidates(pages, expenses)
    written_taxi_candidates = _taxi_equation_missing_components(
        pages, expenses, _document_year("\n".join(pages), None),
    )
    written_taxi_totals = _written_taxi_summary_totals(pages)
    recognized_taxi_total = _recognized_taxi_total(expenses)
    taxi_summary_mismatches = {
        page_number: {
            "written_total_cny": round(written_total, 2),
            "recognized_total_cny": recognized_taxi_total,
            "difference_cny": round(written_total - recognized_taxi_total, 2),
        }
        for page_number, written_total in written_taxi_totals.items()
        if abs(written_total - recognized_taxi_total) > 0.01
    }
    records: list[dict] = []
    for page_number, page in enumerate(pages, 1):
        page_expenses = [
            expense for expense in expenses
            if expense.source_page == page_number
            or (len(pages) == 1 and expense.source_page is None)
        ]
        inferred_ticket_expenses = [
            expense for expense in page_expenses
            if expense.category == "ticket" and expense.inferred_from_document_summary
        ]
        page_invoices = [
            invoice for invoice in invoices
            if invoice.source.get("page") == page_number
            or (len(pages) == 1 and invoice.source.get("page") is None)
        ]
        normalised = _normalise_text(page)
        upper = normalised.upper()
        ocr_candidates, candidates_are_supporting = _ocr_page_invoice_records(normalised)
        if candidates_are_supporting:
            ocr_candidates = []
        annotation_amount_matches = [
            amount for value in re.findall(
                r"IAV_TAXI_NOTE_AMOUNT_REVIEW\s+¥\s*"
                r"([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)",
                normalised,
                flags=re.IGNORECASE,
            )
            if (amount := _plausible_amount(value)) is not None
        ]
        document_summary_sources: dict[tuple[float, Optional[str]], str] = {}
        for amount in written_ticket_candidates.get(page_number, []):
            if not any(
                category == "ticket" and abs(existing - amount) <= 0.01
                for existing, category in ocr_candidates
            ):
                ocr_candidates.append((amount, "ticket"))
            document_summary_sources[(round(amount, 2), "ticket")] = "document_summary"
        for amount in written_taxi_candidates.get(page_number, []):
            if not any(
                category == "taxi" and abs(existing - amount) <= 0.01
                for existing, category in ocr_candidates
            ):
                ocr_candidates.append((amount, "taxi"))
            document_summary_sources[(round(amount, 2), "taxi")] = "document_summary"
        equation_evidence = _dated_equation_evidence(
            normalised, _document_year(normalised, None),
        )

        def equation_matches_adjacent_receipts(equation: dict) -> bool:
            declared_groups = {
                expense.allocation_group
                for expense in expenses
                if expense.category == "taxi"
                and expense.allocation_group
                and expense.source_page is not None
                and abs(expense.source_page - page_number) <= 1
                and expense.allocation_group_total is not None
                and abs(expense.allocation_group_total - equation["total"]) <= 0.01
            }
            for group_name in declared_groups:
                members = [
                    expense for expense in expenses
                    if expense.category == "taxi"
                    and expense.allocation_group == group_name
                ]
                if (
                    members
                    and not any(member.allocation_group_uses_written_total for member in members)
                    and abs(round(sum(member.amount for member in members), 2) - equation["total"]) <= 0.01
                ):
                    return True
            selected_indices: set[int] = set()
            for component in equation["components"]:
                candidates = [
                    index for index, expense in enumerate(expenses)
                    if index not in selected_indices
                    and expense.category == "taxi"
                    and expense.source_page is not None
                    and abs(expense.source_page - page_number) <= 1
                    and abs(expense.amount - component) <= 0.01
                ]
                if not candidates:
                    return False
                selected_indices.add(min(
                    candidates,
                    key=lambda index: (
                        abs((expenses[index].source_page or page_number) - page_number),
                        expenses[index].source_page or 0,
                        index,
                    ),
                ))
            return abs(round(sum(
                expenses[index].amount for index in selected_indices
            ), 2) - equation["total"]) <= 0.01

        def nearby_taxi_expense_matches(amount: float) -> bool:
            return any(
                expense.category == "taxi"
                and expense.source_page is not None
                and abs(expense.source_page - page_number) <= 1
                and abs(expense.amount - amount) <= 0.01
                for expense in expenses
            )

        verified_equations = [
            equation for equation in equation_evidence
            if equation_matches_adjacent_receipts(equation)
        ]
        verified_equation_values = {
            round(value, 2)
            for equation in verified_equations
            for value in (equation["total"], *equation["components"])
        }
        annotation_amount_matches = [
            amount for amount in annotation_amount_matches
            if not any(abs(amount - value) <= 0.01 for value in verified_equation_values)
        ]
        if verified_equations:
            verified_total_values = {
                round(equation["total"], 2) for equation in verified_equations
            }
            ocr_candidates = [
                (amount, category) for amount, category in ocr_candidates
                if category != "taxi"
                or not any(
                    abs(amount - value) <= 0.01
                    for equation in verified_equations
                    for value in (equation["total"], *equation["components"])
                )
            ]
            ocr_candidates.extend((value, "taxi") for value in sorted(verified_total_values))
        for amount in annotation_amount_matches:
            if not any(
                category == "taxi" and abs(existing - amount) <= 0.01
                for existing, category in ocr_candidates
            ):
                ocr_candidates.append((amount, "taxi"))
        unmatched_candidates = []
        category_names = {
            "accommodation": "住宿",
            "taxi": "打车/市内交通",
            "ticket": "机票/火车票",
            "special": "打车/市内交通",
        }
        for amount, category in ocr_candidates:
            matched = category is not None and any(
                abs(expense.amount - amount) <= 0.01
                and expense.category == category
                for expense in page_expenses
            )
            if not matched and category == "taxi":
                matched = any(
                    abs(equation["total"] - amount) <= 0.01
                    for equation in verified_equations
                ) or nearby_taxi_expense_matches(amount)
            if not matched:
                same_amount_categories = sorted({
                    expense.category for expense in page_expenses
                    if abs(expense.amount - amount) <= 0.01
                })
                unmatched_candidates.append({
                    "amount_cny": round(amount, 2),
                    "category": category,
                    "existing_categories": same_amount_categories,
                    "source": (
                        "handwritten_annotation"
                        if any(abs(amount - annotation) <= 0.01 for annotation in annotation_amount_matches)
                        else document_summary_sources.get((round(amount, 2), category), "invoice_text")
                    ),
                })
        has_didi_table = bool(re.search(r"DIDI\s*TRAVEL", upper) and "TRIP TABLE" in upper)
        didi_table_totals = _didi_trip_table_totals(normalised) if has_didi_table else []
        has_hotel_folio = (
            _hotel_folio_evidence(normalised, _document_year(normalised, None)) is not None
        )
        if not has_hotel_folio:
            lowered = normalised.lower()
            has_folio_layout = bool(
                re.search(r"\bfolio\s*(?:no\.?|number|#)", lowered)
                and re.search(r"\b(?:arrival|departure|check[ -]*in|check[ -]*out)\b", lowered)
                and any(marker in lowered for marker in ("hotel", "inn", "plaza", "酒店", "皇冠"))
            )
            has_hotel_folio = has_folio_layout
        is_order_detail = "订单详情" in normalised and "发票号码" not in normalised
        is_residence_registration = (
            "临时住宿登记" in normalised
            or "registration form of temporary residence" in normalised.casefold()
        )
        is_itinerary = any(
            marker in normalised
            for marker in ("行程单", "差旅行程", "行程安排", "出差申请", "出差审批")
        )
        is_toll_summary = _is_toll_summary_page(normalised)
        supporting_only = (
            has_didi_table or has_hotel_folio or is_order_detail
            or is_itinerary or is_residence_registration or is_toll_summary
        )
        has_receipt_cue = bool(re.search(
            r"发票|电子发票|数电发票|invoice|价税合计|发票代码|发票号码|"
            r"出租车票|出租汽车|机打发票|通行费票|客票|票号|taxi|receipt",
            normalised,
            flags=re.IGNORECASE,
        ))
        has_taxi_receipt_cue = _looks_like_taxi_invoice(normalised)
        has_ticket_receipt_cue = bool(re.search(
            r"机票|航空运输|航空客运|电子客票|航空公司|登机牌|乘机人|航班号|"
            r"BOARDING\s*PASS|E[ -]?TICKET",
            normalised,
            flags=re.IGNORECASE,
        ))
        has_currency_amount = bool(re.search(
            r"(?:¥|￥|RMB|CNY)\s*\d|\d[\d,]*\.\s*\d{1,2}\s*(?:元|RMB|CNY)",
            normalised,
            flags=re.IGNORECASE,
        ))
        has_possible_embedded_taxi = has_didi_table and bool(re.search(
            r"出租(?:车|汽车).{0,30}(?:票|发票|机打|金额|有限公司)|"
            r"(?:票|发票|机打|金额).{0,30}出租(?:车|汽车)|"
            r"(?:BEIJING|CHENGDU|SHANGHAI).{0,30}TAXI|TAXI\s*RECEIPT|"
            r"IAV_EMBEDDED_TAXI_RECEIPT",
            normalised,
            flags=re.IGNORECASE,
        ))
        reasons: list[str] = []
        review_required = False
        excluded_cross_city_receipts = []
        for match in re.finditer(
            r"IAV_OUT_OF_SCOPE_TAXI_RECEIPT\s+AMOUNT="
            r"([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)\s+"
            r"CITY=([^\s]+)\s+TRIP_CITY=([^\s]+)",
            normalised,
            flags=re.IGNORECASE,
        ):
            amount = _plausible_amount(match.group(1))
            if amount is not None:
                excluded_cross_city_receipts.append({
                    "amount": round(amount, 2),
                    "city": match.group(2).replace("_", " "),
                    "trip_city": match.group(3).replace("_", " "),
                })
        if page_number in failed_pages:
            status = "OCR_FAILED"
            reasons.append("页面 OCR 失败或未返回文字")
            review_required = True
        elif not normalised.strip():
            status = "NO_TEXT"
            reasons.append("页面没有可提取文字，需确认是否为空白页")
            review_required = True
        elif any(
            issue.get("code") == "DUPLICATE_INVOICE"
            for invoice in page_invoices
            for issue in invoice.validation.get("issues", [])
        ):
            status = "DUPLICATE_REVIEW"
            reasons.append("发票结构检测发现跨页重复候选，未自动删除")
            review_required = True
        elif annotation_amount_matches:
            status = "ANNOTATION_AMOUNT_REVIEW"
            amounts = "、".join(f"¥{amount:.2f}" for amount in annotation_amount_matches)
            reasons.append(
                f"从同页手写批注读到金额候选 {amounts}，尚未能与票面金额核实；请确认后再计入出租车费用"
            )
            review_required = True
        elif unmatched_candidates and not is_toll_summary:
            status = "PARTIAL_INVOICE_COVERAGE" if page_expenses else "UNMATCHED_INVOICE_CUE"
            summaries = "、".join(
                f"¥{candidate['amount_cny']:.2f}"
                + ("（类别不确定）" if candidate["category"] is None else "")
                + (
                    f"（识别为{category_names.get(candidate['category'], candidate['category'])}，"
                    f"当前记录为{'/'.join(category_names.get(item, item) for item in candidate['existing_categories'])}）"
                    if candidate["existing_categories"] and candidate["category"] is not None
                    else ""
                )
                for candidate in unmatched_candidates
            )
            if any(candidate["category"] is None for candidate in unmatched_candidates):
                reasons.append(
                    f"识别到金额候选 {summaries}，但费用类别不确定，未可靠归入费用"
                )
            elif any(
                candidate["source"] == "document_summary"
                and candidate["category"] == "ticket"
                for candidate in unmatched_candidates
            ):
                missing_summary_amounts = "、".join(
                    f"¥{candidate['amount_cny']:.2f}"
                    for candidate in unmatched_candidates
                    if candidate["source"] == "document_summary"
                    and candidate["category"] == "ticket"
                )
                reasons.append(
                    f"PDF 机票汇总列出金额 {missing_summary_amounts}，"
                    "但未匹配到对应机票票据；请核对是否漏了返程/另一程机票"
                )
            elif any(
                candidate["source"] == "document_summary"
                and candidate["category"] == "taxi"
                for candidate in unmatched_candidates
            ):
                missing_summary_amounts = "、".join(
                    f"¥{candidate['amount_cny']:.2f}"
                    for candidate in unmatched_candidates
                    if candidate["source"] == "document_summary"
                    and candidate["category"] == "taxi"
                )
                reasons.append(
                    f"PDF 打车等式列出金额 {missing_summary_amounts}，"
                    "但未匹配到相应票据；已按等式总额汇总，仍需核对缺失发票"
                )
            else:
                reasons.append(
                    f"识别到金额候选 {summaries}，但未匹配到同金额、同类别的费用记录"
                )
            review_required = True
        elif page_number in taxi_summary_mismatches:
            mismatch = taxi_summary_mismatches[page_number]
            status = "WRITTEN_SUMMARY_MISMATCH"
            reasons.append(
                f"PDF 打车汇总写明 ¥{mismatch['written_total_cny']:.2f}，"
                f"当前识别到的打车票据合计 ¥{mismatch['recognized_total_cny']:.2f}，"
                f"相差 ¥{mismatch['difference_cny']:.2f}；请核对是否漏识别票据"
            )
            review_required = True
        elif has_didi_table and didi_table_totals and not has_possible_embedded_taxi and any(
            expense.category == "taxi"
            and not any(abs(expense.amount - total) <= 0.01 for total in didi_table_totals)
            for expense in page_expenses
        ):
            status = "DIDI_AMOUNT_MISMATCH"
            reasons.append("Didi 行程汇总与同页出租车票金额不同，需确认是否属于本次行程")
            review_required = True
        elif inferred_ticket_expenses:
            status = "PARTIAL_INVOICE_COVERAGE"
            review_required = True
        elif len(page_expenses) != len(page_invoices):
            status = "PROJECTION_MISMATCH"
            reasons.append(
                f"Expense {len(page_expenses)} 条，与 structured invoice "
                f"{len(page_invoices)} 条不一致"
            )
            review_required = True
        elif has_ticket_receipt_cue and not any(
            expense.category == "ticket" for expense in page_expenses
        ):
            status = "UNMATCHED_TICKET_CUE"
            reasons.append("本页有机票/客票线索，但未匹配到机票费用记录")
            review_required = True
        elif page_expenses:
            status = "EXPENSES_FOUND"
        elif excluded_cross_city_receipts and not page_expenses and not page_invoices:
            status = "SUPPORTING_ONLY"
            reasons.append("本页仅有已排除的异地出租车票，不计入本次行程")
        elif has_possible_embedded_taxi:
            status = "UNMATCHED_RECEIPT_CUE"
            reasons.append("Didi 行程单页含出租车票线索，但未匹配到可计费记录")
            review_required = True
        elif has_taxi_receipt_cue and not any(
            expense.category == "taxi" for expense in page_expenses
        ):
            status = "UNMATCHED_RECEIPT_CUE"
            reasons.append("本页疑似出租车/市内交通票据，但未读到或未匹配到可核验金额")
            review_required = True
        elif is_toll_summary:
            status = "SUPPORTING_ONLY"
        elif supporting_only and not has_receipt_cue:
            status = "SUPPORTING_ONLY"
        elif has_receipt_cue or has_currency_amount:
            status = "UNMATCHED_INVOICE_CUE"
            reasons.append("页面含票据/金额线索，但未匹配到可计费记录")
            review_required = True
        else:
            status = "UNCLASSIFIED_PAGE"
            if is_pdf and len(normalised) >= 500:
                reasons.append(
                    "PDF 扫描页有文字但未形成可验证的票据/金额候选，可能存在方向或识别质量问题；请核对本页原图"
                )
            else:
                reasons.append("页面有文字，但未确认是费用票据还是辅助材料")
            review_required = True

        inferred_ticket_amounts = sorted({
            round(expense.amount, 2) for expense in inferred_ticket_expenses
        })
        inferred_ticket_risk = None
        if inferred_ticket_expenses:
            inferred_total = round(sum(expense.amount for expense in inferred_ticket_expenses), 2)
            amount_summary = "、".join(f"¥{amount:.2f}" for amount in inferred_ticket_amounts)
            inferred_ticket_risk = {
                "code": "WRITTEN_TICKET_AMOUNT_INFERRED",
                "severity": "WARNING",
                "source_page": page_number,
                "amount_cny": inferred_total,
                "amounts_cny": inferred_ticket_amounts,
                "description": (
                    f"PDF 航段/机票汇总支持机票金额 {amount_summary}，已按机票费用计入；"
                    "未识别到对应的独立票面金额，请核对原始机票及行程。"
                ),
            }
            reasons.append(inferred_ticket_risk["description"])
            review_required = True

        if excluded_cross_city_receipts:
            summary = "、".join(
                f"¥{item['amount']:.2f}（{item['city']}，行程城市 {item['trip_city']}）"
                for item in excluded_cross_city_receipts
            )
            reasons.append(f"同页异地出租车票已排除出本次行程：{summary}")

        risk_flags = []
        if status == "UNMATCHED_RECEIPT_CUE":
            risk_flags.append({
                "code": "UNMATCHED_TAXI_RECEIPT",
                "severity": "WARNING",
                "source_page": page_number,
                "description": (
                    "本页有出租车或市内交通票据线索，但未匹配到费用记录；请核对票面金额和归属。"
                    + ("同页滴滴行程单汇总不代表出租车票金额。" if has_didi_table else "")
                ),
            })
        elif status == "UNMATCHED_INVOICE_CUE":
            risk_flags.append({
                "code": "UNMATCHED_INVOICE",
                "severity": "WARNING",
                "source_page": page_number,
                "description": (
                    "本页有票据或金额线索，但未匹配到费用记录；请核对票面金额和费用类别。"
                ),
            })
        elif status == "PARTIAL_INVOICE_COVERAGE" and unmatched_candidates:
            risk_flags.append({
                "code": "INVOICE_COVERAGE_GAP",
                "severity": "WARNING",
                "source_page": page_number,
                "description": "本页有未匹配的发票金额候选；请核对是否漏票或类别识别错误。",
            })
        elif status == "UNMATCHED_TICKET_CUE":
            risk_flags.append({
                "code": "UNMATCHED_TICKET_RECEIPT",
                "severity": "WARNING",
                "source_page": page_number,
                "description": "本页有机票/客票线索，但未形成机票费用记录；请检查是否漏票。",
            })
        elif status == "ANNOTATION_AMOUNT_REVIEW":
            risk_flags.append({
                "code": "HANDWRITTEN_AMOUNT_REVIEW",
                "severity": "WARNING",
                "source_page": page_number,
                "description": "从出租车票同页手写批注读到金额候选；需确认批注金额与票面一致。",
            })
        if page_number in taxi_summary_mismatches:
            mismatch = taxi_summary_mismatches[page_number]
            risk_flags.append({
                "code": "WRITTEN_SUMMARY_MISMATCH",
                "severity": "WARNING",
                "source_page": page_number,
                **mismatch,
                "description": (
                    "PDF 中已写明的打车汇总与识别到的票据合计不一致；"
                    "按票据顺序复核缺失项目后再入账。"
                ),
            })
        if inferred_ticket_risk:
            risk_flags.append(inferred_ticket_risk)
        receipt_city = _out_of_trip_taxi_receipt_city(normalised, trip_city or "")
        if receipt_city:
            taxi_amounts = [
                expense.amount for expense in page_expenses
                if expense.category == "taxi"
            ]
            risk = {
                "code": "TAXI_RECEIPT_CITY_MISMATCH",
                "severity": "WARNING",
                "source_page": page_number,
                "description": (
                    f"出租车票据城市 {receipt_city} 与行程城市 "
                    f"{_normalise_city(trip_city or '')} 不一致，请复核票面金额及是否属于本次行程；"
                    "城市不一致本身不会自动排除费用。"
                ),
            }
            if taxi_amounts:
                risk["amount_cny"] = round(sum(taxi_amounts), 2)
            risk_flags.append(risk)
            reasons.append(risk["description"])
            review_required = True

        record = {
            "source_page": page_number,
            "unit": "pdf_page" if is_pdf else "document_text_unit",
            "status": status,
            "supporting_only": supporting_only,
            "has_didi_table": has_didi_table,
            "has_receipt_cue": has_receipt_cue,
            "has_possible_embedded_taxi": has_possible_embedded_taxi,
            "text_chars": len(normalised),
            "expense_count": len(page_expenses),
            "invoice_count": len(page_invoices),
            "inferred_expense_count": len(inferred_ticket_expenses),
            "amounts": sorted({round(expense.amount, 2) for expense in page_expenses}),
            "ocr_candidate_count": len(ocr_candidates),
            "written_summary_mismatch": taxi_summary_mismatches.get(page_number),
            "ocr_candidates": [
                {
                    "amount_cny": round(amount, 2),
                    "category": category,
                    "matched": category is not None and any(
                        abs(expense.amount - amount) <= 0.01
                        and expense.category == category
                        for expense in page_expenses
                    ) or (
                        category == "taxi"
                        and (
                            any(
                                abs(equation["total"] - amount) <= 0.01
                                for equation in verified_equations
                            )
                            or nearby_taxi_expense_matches(amount)
                        )
                    ),
                    "source": (
                        "handwritten_annotation"
                        if any(abs(amount - annotation) <= 0.01 for annotation in annotation_amount_matches)
                        else document_summary_sources.get((round(amount, 2), category), "invoice_text")
                    ),
                }
                for amount, category in ocr_candidates
            ],
            "excluded_out_of_scope_receipts": excluded_cross_city_receipts,
            "review_required": review_required,
            "reasons": reasons,
        }
        if risk_flags:
            record["risk_flags"] = risk_flags
        records.append(record)
    return records


def parse_trip_file(path: str, default_year: Optional[int] = None) -> Trip:
    text, warnings = extract_document_text(path)
    year = _document_year(text, default_year)
    start, end, name, city = _infer_trip_identity(text, path, year)
    records: list[dict] = []
    accommodation_folios: list[AccommodationFolioEvidence] = []
    didi_tables: list[DidiTripTableEvidence] = []
    filename_category = _category_from_filename(Path(path).name)
    local_ocr_used = any("OCR" in warning and "已执行" in warning for warning in warnings)

    # PDF page boundaries reset the category context. This prevents hotel
    # labels on one page from classifying taxi route lines on the next page.
    pages = text.replace("\r\n", "\n").split("\f")
    # A receipt's issuer city alone does not establish whether the expense
    # belongs to a trip. Keep every identified receipt available for review;
    # exclusion requires an explicit reviewed action tied to this source PDF.
    for page_index, page in enumerate(pages, 1):
        is_out_of_scope_taxi_page = "IAV_OUT_OF_SCOPE_TAXI_RECEIPT" in page.upper()
        folio = _hotel_folio_evidence(page, year, source_page=page_index)
        if folio is not None:
            folio.city = _hotel_page_city(page)
            accommodation_folios.append(folio)
        didi_table = (
            None if is_out_of_scope_taxi_page
            else _didi_trip_table_evidence(page, year, source_page=page_index)
        )
        if didi_table is not None:
            didi_tables.append(didi_table)
        if local_ocr_used:
            if is_out_of_scope_taxi_page:
                # The page's only detected receipt was explicitly identified
                # as outside this trip. Do not reinterpret its raw OCR digits
                # or Didi supporting total as payable amounts.
                continue
            page_records, supporting_trip_table = _ocr_page_invoice_records(page, filename_category)
            if supporting_trip_table:
                continue
            page_date_heading = _single_page_date_heading(page, year)
            dated_equations = _dated_equation_evidence(page, year)
            declared_equation = None
            if len(dated_equations) == 1 and (
                filename_category == "taxi" or _looks_like_taxi_invoice(page)
            ):
                equation = dated_equations[0]
                if (
                    page_records
                    and all(category == "taxi" for _amount, category in page_records)
                    and abs(round(sum(amount for amount, _category in page_records), 2) - equation["total"]) <= 0.01
                ):
                    declared_equation = equation
                elif not page_records:
                    page_records = [(equation["total"], "taxi")]
                    declared_equation = equation
            handwritten_taxi_candidates = [
                amount
                for match in re.findall(
                    r"IAV_TAXI_NOTE_AMOUNT_REVIEW\s+¥\s*"
                    r"([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)",
                    page,
                    flags=re.IGNORECASE,
                )
                if (amount := _plausible_amount(match)) is not None
            ]
            if declared_equation is not None:
                verified_values = [
                    declared_equation["total"], *declared_equation["components"],
                ]
                handwritten_taxi_candidates = [
                    amount for amount in handwritten_taxi_candidates
                    if not any(abs(amount - value) <= 0.01 for value in verified_values)
                ]
            for amount in handwritten_taxi_candidates:
                if not any(abs(existing[0] - amount) <= 0.01 for existing in page_records):
                    page_records.append((amount, "taxi"))
            embedded_markers = [
                amount
                for match in re.findall(
                    r"IAV_EMBEDDED_TAXI_RECEIPT\s+¥\s*"
                    r"([0-9][0-9,]*(?:\.\s*[0-9]{1,2})?)",
                    page,
                    flags=re.IGNORECASE,
                )
                if (amount := _plausible_amount(match)) is not None
            ]
            for page_total, page_category in page_records:
                if page_category is None:
                    continue
                embedded = next(
                    (
                        amount for amount in embedded_markers
                        if abs(amount - page_total) <= 0.01
                    ),
                    None,
                )
                if any(abs(candidate - page_total) <= 0.01 for candidate in handwritten_taxi_candidates):
                    description = (
                        f"OCR page {page_index} handwritten taxi note amount candidate: "
                        f"¥{page_total:.2f} (unverified)"
                    )
                elif embedded is not None:
                    embedded_markers.remove(embedded)
                    description = (
                        f"OCR page {page_index} embedded taxi receipt: "
                        f"¥{page_total:.2f}"
                    )
                else:
                    description = f"OCR page {page_index} gross total: ¥{page_total:.2f}"
                uses_declared_total = (
                    declared_equation is not None
                    and page_category == "taxi"
                    and (
                        abs(page_total - declared_equation["total"]) <= 0.01
                        or len(page_records) > 1
                    )
                )
                uses_page_date = (
                    page_category == "taxi"
                    and page_date_heading is not None
                    and not uses_declared_total
                )
                records.append({
                    "category": page_category,
                    "amount": page_total,
                    "components": [],
                    "start": (
                        declared_equation["date"] if uses_declared_total else
                        page_date_heading[0] if uses_page_date else None
                    ),
                    "end": (
                        declared_equation["date"] if uses_declared_total else
                        page_date_heading[0] if uses_page_date else None
                    ),
                    "service_date_source": (
                        "SOURCE_TEXT"
                        if uses_declared_total and declared_equation["date_text"] == declared_equation["text"]
                        else "PAGE_DATE_HEADING" if uses_declared_total
                        else "PAGE_DATE_HEADING" if uses_page_date else "UNCONFIRMED"
                    ),
                    "service_date_evidence_page": page_index if uses_declared_total or uses_page_date else None,
                    "service_date_evidence_text": (
                        f"{declared_equation['date_text']} {declared_equation['text']}"
                        if uses_declared_total else page_date_heading[1] if uses_page_date else ""
                    ),
                    "description": (
                        f"PDF written taxi total: ¥{declared_equation['total']:.2f}"
                        if uses_declared_total and len(page_records) == 1
                        else description
                    ),
                    "source_page": page_index,
                    "allocation_group": (
                        f"pdf-declared-total-{page_index}" if uses_declared_total else None
                    ),
                    "allocation_group_total": (
                        declared_equation["total"] if uses_declared_total else None
                    ),
                })
            # OCR line-by-line extraction is too prone to interpreting tax
            # amounts, room numbers, phone numbers and invoice IDs as separate
            # expenses. Only currency-marked gross totals are accepted.
            continue
        if folio is not None:
            _folio_records, is_supporting_folio = _ocr_page_invoice_records(page, filename_category)
            if is_supporting_folio:
                continue
        context: Optional[str] = None
        pending_date: Optional[date] = None
        pending_date_raw: Optional[str] = None
        for raw_line in page.splitlines():
            line = _normalise_text(raw_line)
            if not line or line.startswith("--- PAGE"):
                continue
            line_start, line_end = _date_span(line, year)
            date_heading = bool(line_start and _is_date_only_line(line))
            if date_heading:
                pending_date = line_start
                pending_date_raw = line
            explicit = _category_for_line(line)
            if (
                explicit == "taxi"
                and filename_category in {"ticket", "accommodation", "special"}
                and _looks_like_route(line)
                and not any(keyword.lower() in line.lower() for keyword in TAXI_KEYWORDS)
            ):
                # Route-like OCR noise contains hyphens and Latin letters and
                # would otherwise be classified as taxi before the filename
                # fallback gets a chance to run.  For a file explicitly named
                # as a ticket/hotel/special expense, retain that stronger hint.
                explicit = filename_category
            # Route-like flight legs such as “北京-成都” should inherit the
            # current ticket section instead of being mistaken for taxi.
            if explicit == "taxi" and context == "ticket" and not any(
                keyword.lower() in line.lower() for keyword in TAXI_KEYWORDS
            ):
                explicit = None
            if explicit:
                context = explicit

            category = explicit or context
            filename_fallback = False
            if not category and filename_category:
                # OCR of Chinese receipt labels may contain only numbers and
                # noise.  Use a category from the filename for monetary lines,
                # but do not let it classify invoice IDs or date headings.
                filename_fallback = bool(re.search(
                    r"(?:¥|￥|RMB|CNY|yuan|元|\d+\.\s*\d{1,2})",
                    line,
                    flags=re.IGNORECASE,
                ))
                if filename_fallback:
                    category = filename_category
            if not category and _looks_like_route(line):
                category = "taxi"
            if not category:
                continue

            if date_heading:
                # The heading itself has no expense amount; any accidental
                # date-like number should not become an invoice.
                continue
            if filename_fallback and not re.search(
                r"(?:¥|￥|RMB|CNY|yuan|元|\d+\.\s*\d{1,2})",
                line,
                flags=re.IGNORECASE,
            ):
                continue
            amount = _amount_from_line(line)
            components = _equation_components(line)
            if (amount is None or amount <= 0) and not components:
                continue
            item_start, item_end = (line_start, line_end) if line_start else (pending_date, pending_date)
            service_date_source = (
                "SOURCE_TEXT" if line_start else
                "PAGE_DATE_HEADING" if pending_date else
                "UNCONFIRMED"
            )
            written_equation = _validated_equation_total(line)
            declared_group = None
            declared_total = None
            if (
                category == "taxi"
                and written_equation is not None
                and item_start is not None
                and item_end is not None
                and abs(written_equation[0] - round(amount or 0.0, 2)) <= 0.01
            ):
                declared_total = written_equation[0]
                declared_group = f"pdf-declared-total-{page_index}-{len(records)}"
            records.append({
                "category": category,
                "amount": round(amount, 2) if amount is not None else None,
                "components": components,
                "start": item_start,
                "end": item_end,
                "service_date_source": service_date_source,
                "service_date_evidence_page": page_index if service_date_source in {"SOURCE_TEXT", "PAGE_DATE_HEADING"} else None,
                "service_date_evidence_text": (
                    line if service_date_source == "SOURCE_TEXT" else
                    pending_date_raw or "" if service_date_source == "PAGE_DATE_HEADING" else
                    ""
                ),
                "description": line,
                "source_page": page_index,
                "allocation_group": declared_group,
                "allocation_group_total": declared_total,
            })

    # A summary such as 979+1268=2247 may be followed by the two route-level
    # equations 930+49=979 and 1210+58=1268. Keep the invoice components and
    # drop the higher-level summary to avoid counting the same tickets twice.
    route_totals: dict[str, list[float]] = defaultdict(list)
    for record in records:
        if record["components"] and _looks_like_route(record["description"]):
            if record["amount"] is not None:
                route_totals[record["category"]].append(record["amount"])

    expenses: list[Expense] = []
    for record in records:
        components = record["components"]
        if components:
            dated_detail_total = round(sum(
                other["amount"] for other in records
                if other is not record
                and other["category"] == record["category"]
                and other["start"] is not None
                and other["amount"] is not None
            ), 2)
            is_outer_summary = (
                not _looks_like_route(record["description"])
                and len(components) >= 2
                and (
                    all(
                        any(abs(value - subtotal) <= 0.01 for subtotal in route_totals[record["category"]])
                        for value in components
                    )
                    or (
                        record["start"] is None
                        and record["amount"] is not None
                        and dated_detail_total
                        and abs(dated_detail_total - record["amount"]) <= 0.01
                    )
                )
            )
            if is_outer_summary:
                continue
            amounts = components
        else:
            amounts = [record["amount"]]
        for amount in amounts:
            expense = Expense(
                category=record["category"],
                amount=round(amount, 2),
                start=record["start"],
                end=record["end"],
                description=record["description"],
                source_file=os.path.basename(path),
                aggregate=(not components and "=" in record["description"]),
                source_page=record.get("source_page"),
                allocation_group=record.get("allocation_group"),
                allocation_group_total=record.get("allocation_group_total"),
                service_date_source=record.get("service_date_source", "UNCONFIRMED"),
                service_date_evidence_page=record.get("service_date_evidence_page"),
                service_date_evidence_text=record.get("service_date_evidence_text", ""),
            )
            if (
                expense.category == "accommodation"
                and expense.start is not None
                and expense.end is not None
                and expense.end > expense.start
                and expense.service_date_source in {"SOURCE_TEXT", "PAGE_DATE_HEADING"}
            ):
                expense.accommodation_nights = (expense.end - expense.start).days
                expense.accommodation_nights_source = "SOURCE_TEXT_RANGE"
            expenses.append(expense)

    expenses = _deduplicate_aggregate_expenses(expenses)
    expenses = _deduplicate_repeated_invoice_notes(expenses)
    expenses = _deduplicate_summary_sheet_expenses(expenses, pages)
    _reconcile_accommodation_folios(expenses, accommodation_folios, start, end)
    _reconcile_dated_taxi_equations(expenses, pages, year)
    _reconcile_didi_trip_tables(expenses, didi_tables, start, end)
    # Flight-summary operands are payable-category evidence even when OCR did
    # not produce a standalone ticket invoice. Count the likely fare, then
    # keep the missing invoice as an explicit review risk on its source page.
    expenses.extend(_infer_written_ticket_summary_expenses(
        pages, expenses, os.path.basename(path),
    ))
    # Later reconciliation may recover records after the original pass. Put
    # the final expense list back in stable PDF-page order; same-page line
    # order remains unchanged.
    expenses = [
        expense for _index, expense in sorted(
            enumerate(expenses),
            key=lambda item: (
                item[1].source_page is None,
                item[1].source_page if item[1].source_page is not None else 0,
                item[0],
            ),
        )
    ]
    inferred_trip = Trip(
        start=start,
        end=end,
        name=name,
        city=city,
        source_file=os.path.basename(path),
        expenses=expenses,
        warnings=warnings,
    )
    _infer_accommodation_spans(inferred_trip)
    dated_expense_days = [expense.start for expense in expenses if expense.start]
    if dated_expense_days and not any(start <= day <= end for day in dated_expense_days):
        # Some Word packs retain an old trip heading (e.g. 06/22-06/26) while
        # the invoice pages and filename are for a later trip (e.g. 08/16-08/20).
        # Prefer the dates attached to actual expense lines when the heading
        # has no overlap, otherwise all dated taxi rows collapse onto one day.
        corrected_start = min(dated_expense_days)
        corrected_end = max(dated_expense_days)
        warnings.append(
            f"文档标题日期 {start:%m/%d}-{end:%m/%d} 与发票明细日期不一致，已采用明细日期 {corrected_start:%m/%d}-{corrected_end:%m/%d}。"
        )
        start, end = corrected_start, corrected_end
    # Undated ticket route equations are normally the outbound and return
    # legs. Put them on the first and last trip days so each same-day group
    # remains visible as its own formula in column K.
    route_groups: list[str] = []
    for expense in expenses:
        if (
            expense.category == "ticket"
            and expense.start is None
            and not expense.inferred_from_document_summary
            and _looks_like_route(expense.description)
            and expense.description not in route_groups
        ):
            route_groups.append(expense.description)
    if route_groups:
        trip_days = _date_range(start, end)
        for index, description in enumerate(route_groups):
            if len(route_groups) == 1 or index == 0:
                target_day = start
            elif index == len(route_groups) - 1:
                target_day = end
            else:
                target_day = trip_days[min(index, len(trip_days) - 1)]
            for expense in expenses:
                if expense.description == description and expense.category == "ticket" and expense.start is None:
                    expense.start = target_day
                    expense.end = target_day
                    expense.service_date_source = "TRIP_ROUTE_INFERENCE"
    # Phase D structured projection.  The existing Expense list remains the
    # authority for workbook allocation; these records add field-level
    # evidence, arithmetic validation and duplicate-review metadata without
    # changing payable values.
    inferred_trip.start = start
    inferred_trip.end = end
    inferred_trip.city_by_day = _infer_trip_city_timeline(
        pages,
        accommodation_folios,
        start,
        end,
        year,
        city,
    )
    first_city = next(
        (inferred_trip.city_by_day.get(day) for day in inferred_trip.days
         if inferred_trip.city_by_day.get(day)),
        "",
    )
    if first_city:
        inferred_trip.city = first_city
    expense_pages = {
        expense.source_page for expense in expenses
        if expense.source_page is not None
    }
    structured_pages = []
    for page_number, page in enumerate(pages, 1):
        if "IAV_OUT_OF_SCOPE_TAXI_RECEIPT" in page.upper():
            structured_pages.append("")
            continue
        if page_number not in expense_pages and _didi_trip_table_evidence(
            page, year, source_page=page_number,
        ) is not None:
            # A Didi trip table can corroborate dates or reveal an unmatched
            # receipt cue, but its trip sum is not a standalone invoice.
            structured_pages.append("")
            continue
        structured_pages.append(re.sub(
            r"IAV_DIDI_REVIEWED_TOTAL\s+¥\s*"
            r"[0-9][0-9,]*(?:\.\s*[0-9]{1,2})?",
            "",
            page,
            flags=re.IGNORECASE,
        ))
    inferred_trip.structured_invoices = extract_travel_invoices(
        structured_pages,
        source_file=path,
        expenses=expenses,
        traveler=name,
        trip_start=start,
        trip_end=end,
        default_year=year,
        ocr_engine="tesseract" if local_ocr_used else "native_text",
    )
    inferred_trip.page_coverage = _page_coverage_records(
        pages,
        expenses,
        inferred_trip.structured_invoices,
        warnings,
        is_pdf=Path(path).suffix.lower() == ".pdf",
        trip_city=city,
    )
    # Structured projections are emitted in Expense order. Carry the same
    # per-invoice ID on the writing object so a reviewed correction can update
    # both representations without matching on amount alone.
    if len(inferred_trip.structured_invoices) == len(expenses):
        for expense, invoice in zip(expenses, inferred_trip.structured_invoices):
            if (
                expense.source_page == invoice.source.get("page")
                and expense.category == invoice.expense_type
                and abs(expense.amount - (invoice.total_amount or 0.0)) <= 0.01
            ):
                expense.document_id = invoice.document_id
    if not expenses:
        warnings.append("没有识别到带金额的机票/火车票、交通、住宿或特殊费用明细，请人工复核。")
    return inferred_trip


def _correction_value(value):
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return value


def _correction_values_match(actual, expected) -> bool:
    actual_value = _correction_value(actual)
    expected_value = _correction_value(expected)
    if isinstance(actual_value, (int, float)) and isinstance(expected_value, (int, float)):
        return abs(float(actual_value) - float(expected_value)) <= 0.01
    return actual_value == expected_value


def apply_manual_travel_corrections(trip: Trip, manifest: dict) -> list[dict]:
    """Apply source-reviewed field corrections to one Trip and both invoice views.

    Each item must identify an invoice by source hash, page, expense category,
    and original amount. Expected-current values guard against applying a stale
    correction to a changed OCR result. Ambiguous matches fail closed.
    """

    source_hash = str(manifest.get("source_sha256", "")).strip().lower()
    if len(source_hash) != 64 or any(char not in "0123456789abcdef" for char in source_hash):
        raise ValueError("人工修正清单必须包含完整 SHA-256。")
    reviewed_on = str(manifest.get("reviewed_on", "")).strip()
    correction_items = manifest.get("corrections")
    if not isinstance(correction_items, list) or not correction_items:
        raise ValueError("人工修正清单没有 corrections 项。")

    structured_fields = {
        "document_type", "vendor", "invoice_number", "invoice_date", "currency",
        "net_amount", "tax_amount", "total_amount", "expense_type",
        "travel_date_from", "travel_date_to",
    }
    expense_fields = {
        "start", "end", "description", "allocation_anchor", "amount", "category",
        "allocation_splits",
    }
    numeric_fields = {"net_amount", "tax_amount", "total_amount", "amount"}
    date_fields = {
        "invoice_date", "travel_date_from", "travel_date_to",
        "start", "end", "allocation_anchor",
    }

    def parse_field_value(field_name: str, value):
        if field_name in date_fields:
            if value in (None, ""):
                return None
            return date.fromisoformat(str(value))
        if field_name in numeric_fields:
            if value in (None, ""):
                return None
            number = float(value)
            if not math.isfinite(number):
                raise ValueError(f"人工修正字段 {field_name} 必须是有限数字。")
            return round(number, 2)
        if field_name == "allocation_splits":
            if value in (None, ""):
                return []
            if not isinstance(value, list):
                raise ValueError("allocation_splits 必须是列表。")
            splits = []
            for split in value:
                if not isinstance(split, dict):
                    raise ValueError("allocation_splits 的每项必须包含日期和金额。")
                split_date = date.fromisoformat(str(split.get("date", "")))
                split_amount = float(split.get("amount", 0))
                if not math.isfinite(split_amount) or split_amount <= 0:
                    raise ValueError("allocation_splits 的金额必须为正的有限数字。")
                splits.append({"date": split_date.isoformat(), "amount": round(split_amount, 2)})
            if not splits:
                return []
            if round(sum(split["amount"] for split in splits), 2) <= 0:
                raise ValueError("allocation_splits 合计必须为正数。")
            return splits
        if field_name in {"category", "expense_type"}:
            return _normalise_expense_category(value)
        return value

    applied: list[dict] = []
    for item in correction_items:
        match = item.get("invoice_match") or {}
        page = int(match.get("page", 0))
        category = _normalise_expense_category(match.get("expense_type", "")) or ""
        original_amount = float(match.get("original_total_amount", -1))
        if page < 1 or not category or original_amount < 0:
            raise ValueError("修正项必须指定页码、expense_type 和原始总金额。")

        invoices = [
            invoice for invoice in trip.structured_invoices
            if invoice.source.get("sha256", "").lower() == source_hash
            and int(invoice.source.get("page", 0)) == page
            and invoice.expense_type == category
            and invoice.total_amount is not None
            and abs(invoice.total_amount - original_amount) <= 0.01
        ]
        expenses = [
            expense for expense in trip.expenses
            if expense.source_page == page
            and expense.category == category
            and abs(expense.amount - original_amount) <= 0.01
        ]
        if len(invoices) != 1 or len(expenses) != 1:
            raise ValueError(
                f"修正项 {item.get('record_id', '<unknown>')} 匹配到 "
                f"{len(invoices)} 条 structured invoice、{len(expenses)} 条 Expense；"
                "请先消除歧义，不会按金额猜测。"
            )

        invoice, expense = invoices[0], expenses[0]
        changes: list[dict] = []
        evidence_location = str(item.get("evidence_location", "源 PDF 页面复核"))
        reason = str(item.get("reason", "源图人工复核"))

        for field_name, update in (item.get("structured_fields") or {}).items():
            if field_name not in structured_fields:
                raise ValueError(f"不允许人工修正 structured 字段: {field_name}")
            current = getattr(invoice, field_name)
            expected_current = parse_field_value(field_name, update.get("expected_current"))
            if not _correction_values_match(current, expected_current):
                raise ValueError(
                    f"{item.get('record_id')}.{field_name} 当前值与清单预期不符；"
                    "拒绝覆盖新识别结果。"
                )
            confirmed = parse_field_value(field_name, update.get("value"))
            if _correction_values_match(current, confirmed):
                continue
            setattr(invoice, field_name, confirmed)
            if field_name == "document_type":
                invoice.source["document_type"] = confirmed
            source_text = str(update.get("source_text", ""))
            changes.append({
                "object": "structured_invoice",
                "field": field_name,
                "original_value": _correction_value(current),
                "confirmed_value": _correction_value(confirmed),
                "source_text": source_text,
                "evidence_location": evidence_location,
                "reason": reason,
            })
            invoice.evidence[field_name] = {
                "field": field_name,
                "value": _correction_value(confirmed),
                "source_file": invoice.source.get("source_file", trip.source_file),
                "page": page,
                "bbox": None,
                "raw_text": source_text,
                "ocr_engine": "manual_source_review",
                "ocr_confidence": None,
                "validation_status": "MANUAL_CONFIRMED",
                "evidence_location": evidence_location,
            }

        confirm_service_date = bool(item.get("confirm_service_date"))
        accept_unknown_service_date = bool(item.get("accept_unknown_service_date"))
        if (
            any(change.get("field") in {"travel_date_from", "travel_date_to"} for change in changes)
            and not confirm_service_date
            and not accept_unknown_service_date
        ):
            previous_date_source = invoice.travel_date_source
            invoice.travel_date_source = "MANUAL_SOURCE_REVIEW"
            invoice.source["service_date_evidence_page"] = page
            changes.append({
                "object": "structured_invoice",
                "field": "travel_date_source",
                "original_value": previous_date_source,
                "confirmed_value": invoice.travel_date_source,
                "source_text": "源图人工复核",
                "evidence_location": evidence_location,
                "reason": reason,
            })

        if confirm_service_date and accept_unknown_service_date:
            raise ValueError(f"{item.get('record_id')} 不能同时确认日期并接受日期未知。")
        service_date_evidence_page = int(item.get("service_date_evidence_page") or page)
        if service_date_evidence_page < 1:
            raise ValueError(f"{item.get('record_id')} 的服务日期证据页码无效。")
        if confirm_service_date:
            if not all(
                field_name in (item.get("structured_fields") or {})
                and getattr(invoice, field_name) is not None
                for field_name in ("travel_date_from", "travel_date_to")
            ):
                raise ValueError(f"{item.get('record_id')} 确认日期时必须提供非空的起止日期。")
            previous_date_source = invoice.travel_date_source
            invoice.travel_date_source = "MANUAL_SOURCE_REVIEW"
            invoice.source["service_date_evidence_page"] = service_date_evidence_page
            for field_name in ("travel_date_from", "travel_date_to"):
                evidence = invoice.evidence.get(field_name)
                if evidence and evidence.get("ocr_engine") == "manual_source_review":
                    evidence["page"] = service_date_evidence_page
            changes.append({
                "object": "structured_invoice",
                "field": "travel_date_source",
                "original_value": previous_date_source,
                "confirmed_value": invoice.travel_date_source,
                "source_text": "源 PDF 逐票人工复核",
                "evidence_location": evidence_location,
                "reason": reason,
            })

        for field_name, update in (item.get("expense_fields") or {}).items():
            if field_name not in expense_fields:
                raise ValueError(f"不允许人工修正 Expense 字段: {field_name}")
            current = getattr(expense, field_name)
            expected_current = parse_field_value(field_name, update.get("expected_current"))
            if not _correction_values_match(current, expected_current):
                raise ValueError(
                    f"{item.get('record_id')}.Expense.{field_name} 当前值与清单预期不符；"
                    "拒绝覆盖新识别结果。"
                )
            confirmed = parse_field_value(field_name, update.get("value"))
            if _correction_values_match(current, confirmed):
                continue
            setattr(expense, field_name, confirmed)
            changes.append({
                "object": "Expense",
                "field": field_name,
                "original_value": _correction_value(current),
                "confirmed_value": _correction_value(confirmed),
                "source_text": str(update.get("source_text", "")),
                "evidence_location": evidence_location,
                "reason": reason,
            })

        if confirm_service_date or any(
            change.get("object") == "Expense" and change.get("field") in {"start", "end"}
            for change in changes
        ):
            if confirm_service_date:
                expense_updates = item.get("expense_fields") or {}
                if not all(
                    field_name in expense_updates
                    and _correction_values_match(
                        getattr(expense, field_name),
                        parse_field_value(
                            field_name,
                            (item.get("structured_fields") or {}).get(
                                "travel_date_from" if field_name == "start" else "travel_date_to",
                                {},
                            ).get("value"),
                        ),
                    )
                    for field_name in ("start", "end")
                ):
                    raise ValueError(f"{item.get('record_id')} 的 Expense 服务日期必须与票据起止日期一致。")
            previous_date_source = expense.service_date_source
            expense.service_date_source = "MANUAL_SOURCE_REVIEW"
            expense.service_date_evidence_page = service_date_evidence_page
            date_update = (item.get("expense_fields") or {}).get("start") or {}
            expense.service_date_evidence_text = str(
                date_update.get("source_text", reason)
            )
            changes.append({
                "object": "Expense",
                "field": "service_date_source",
                "original_value": previous_date_source,
                "confirmed_value": expense.service_date_source,
                "source_text": "源图人工复核",
                "evidence_location": evidence_location,
                "reason": reason,
            })

        if accept_unknown_service_date:
            structured_updates = item.get("structured_fields") or {}
            expense_updates = item.get("expense_fields") or {}
            if not all(
                field_name in structured_updates
                and parse_field_value(field_name, structured_updates[field_name].get("value")) is None
                for field_name in ("travel_date_from", "travel_date_to")
            ) or not all(
                field_name in expense_updates
                and parse_field_value(field_name, expense_updates[field_name].get("value")) is None
                for field_name in ("start", "end")
            ):
                raise ValueError(f"{item.get('record_id')} 接受日期未知时必须显式清空票据与 Expense 的起止日期。")
            marker = {
                "review_status": "approved",
                "source_sha256": source_hash,
                "source_page": page,
                "reviewed_on": reviewed_on,
                "record_id": str(item.get("record_id", "")),
                "evidence_location": evidence_location,
                "reason": reason,
            }
            invoice.travel_date_source = "REVIEWED_UNKNOWN_SERVICE_DATE"
            invoice.source["service_date_evidence_page"] = page
            invoice.source["accepted_unknown_service_date"] = copy.deepcopy(marker)
            expense.service_date_source = "REVIEWED_UNKNOWN_SERVICE_DATE"
            expense.service_date_evidence_page = page
            expense.service_date_evidence_text = reason
            changes.extend([
                {
                    "object": "structured_invoice",
                    "field": "service_date_status",
                    "original_value": "UNCONFIRMED",
                    "confirmed_value": "REVIEWED_UNKNOWN_SERVICE_DATE",
                    "source_text": reason,
                    "evidence_location": evidence_location,
                    "reason": reason,
                },
                {
                    "object": "Expense",
                    "field": "service_date_source",
                    "original_value": "UNCONFIRMED",
                    "confirmed_value": "REVIEWED_UNKNOWN_SERVICE_DATE",
                    "source_text": reason,
                    "evidence_location": evidence_location,
                    "reason": reason,
                },
            ])

        if expense.allocation_splits:
            split_total = round(sum(float(split["amount"]) for split in expense.allocation_splits), 2)
            if split_total != round(expense.amount, 2):
                raise ValueError(
                    f"{item.get('record_id')} 的 allocation_splits 合计 {split_total:.2f} "
                    f"与发票总额 {expense.amount:.2f} 不一致。"
                )
            for split in expense.allocation_splits:
                if not expense.start or not expense.end:
                    raise ValueError(f"{item.get('record_id')} 的逐日分配必须保留经确认的发票服务日期范围。")
                split_date = date.fromisoformat(split["date"])
                if split_date < expense.start or split_date > expense.end:
                    raise ValueError(f"{item.get('record_id')} 的逐日分配日期超出已确认的服务日期范围。")

        if not changes:
            raise ValueError(f"修正项 {item.get('record_id')} 没有任何字段变化。")
        stable_id = hashlib.sha256(
            f"{source_hash}|{page}|{category}|{invoice.invoice_number}|{invoice.total_amount:.2f}".encode("utf-8")
        ).hexdigest()[:20]
        old_expense_id = expense.document_id
        invoice.document_id = stable_id
        expense.document_id = stable_id
        if old_expense_id != stable_id:
            changes.append({
                "object": "Expense",
                "field": "document_id",
                "original_value": old_expense_id,
                "confirmed_value": stable_id,
                "source_text": "按源文件 SHA-256、页码、类别、票号和金额关联两种记录",
                "evidence_location": evidence_location,
                "reason": "将写表依据关联到同一 structured invoice。",
            })

        record = {
            "record_id": str(item.get("record_id", "")),
            "action": "correct_fields",
            "action_status": "FIELD_CORRECTION_APPLIED",
            "source_sha256": source_hash,
            "document_id": stable_id,
            "page": page,
            "reviewed_on": reviewed_on,
            "evidence_location": evidence_location,
            "reason": reason,
            "changes": changes,
        }
        invoice.validation = validate_travel_invoice(invoice)
        for evidence in invoice.evidence.values():
            if evidence.get("ocr_engine") != "manual_source_review":
                evidence["validation_status"] = invoice.validation["status"]
        invoice.correction_records.append(copy.deepcopy(record))
        expense.correction_records.append(copy.deepcopy(record))
        invoice.source.setdefault("manual_corrections", []).append(copy.deepcopy(record))
        trip.correction_records.append(copy.deepcopy(record))
        applied.append(record)

    for coverage in trip.page_coverage:
        page = int(coverage.get("source_page", 0))
        if page < 1:
            continue
        page_expenses = [expense for expense in trip.expenses if expense.source_page == page]
        page_invoices = [
            invoice for invoice in trip.structured_invoices
            if int(invoice.source.get("page", 0)) == page
        ]
        coverage.update({
            "expense_count": len(page_expenses),
            "invoice_count": len(page_invoices),
            "amounts": sorted({round(expense.amount, 2) for expense in page_expenses}),
        })

    return applied


def apply_reviewed_travel_actions(
    trip: Trip,
    manifest: dict,
    actual_source_sha256: str,
) -> list[dict]:
    """Apply approved review actions tied to one exact source PDF.

    This is the narrow path for a reviewer-confirmed receipt that OCR missed,
    or a receipt that must be kept out of the trip. It deliberately requires
    an approved manifest and a source hash match; ambiguous page matches fail
    closed instead of guessing by amount alone.
    """

    if not isinstance(manifest, dict) or manifest.get("schema_version") != 1:
        raise ValueError("人工动作清单 schema_version 必须为 1。")
    if str(manifest.get("review_status", "")).strip().lower() != "approved":
        raise ValueError("人工动作清单必须标记 review_status=approved。")
    reviewed_on = str(manifest.get("reviewed_on", "")).strip()
    if not reviewed_on:
        raise ValueError("人工动作清单必须记录 reviewed_on。")

    expected_hash = str(manifest.get("source_sha256", "")).strip().lower()
    actual_hash = str(actual_source_sha256 or "").strip().lower()
    if (
        len(expected_hash) != 64
        or any(char not in "0123456789abcdef" for char in expected_hash)
        or expected_hash != actual_hash
    ):
        raise ValueError("人工动作清单的 source_sha256 与当前 PDF 不一致，拒绝应用。")

    actions = manifest.get("actions", [])
    field_corrections = manifest.get("field_corrections", [])
    if not isinstance(actions, list) or (
        not actions and (not isinstance(field_corrections, list) or not field_corrections)
    ):
        raise ValueError("人工动作清单必须包含 actions 或 field_corrections。")

    page_by_number = {
        int(item.get("source_page", 0)): item
        for item in trip.page_coverage
        if int(item.get("source_page", 0)) > 0
    }

    def clear_resolved_page_risks(coverage: dict) -> None:
        """Drop only generic unmatched-page warnings after a reviewed action resolves them."""
        risk_flags = coverage.get("risk_flags")
        if not isinstance(risk_flags, list):
            return
        resolved_codes = {"UNMATCHED_INVOICE", "UNMATCHED_TAXI_RECEIPT"}
        remaining = [
            risk for risk in risk_flags
            if not isinstance(risk, dict) or risk.get("code") not in resolved_codes
        ]
        if remaining:
            coverage["risk_flags"] = remaining
        else:
            coverage.pop("risk_flags", None)

    applied: list[dict] = []
    seen_record_ids: set[str] = set()
    pages_with_additions: set[int] = set()
    for item in actions:
        if not isinstance(item, dict):
            raise ValueError("人工动作项必须为对象。")
        record_id = str(item.get("record_id", "")).strip()
        action = str(item.get("action", "")).strip().lower()
        if not record_id or record_id in seen_record_ids:
            raise ValueError("人工动作 record_id 必须非空且不能重复。")
        seen_record_ids.add(record_id)
        if action not in {
            "add", "exclude", "include_with_risk", "ignore_fragment", "classify_supporting",
        }:
            raise ValueError(
                f"不支持人工动作 {action!r}；仅支持 add/exclude/include_with_risk/ignore_fragment/classify_supporting。"
            )
        action_reviewed_on = str(item.get("reviewed_on") or reviewed_on).strip()
        if any(record.get("record_id") == record_id for record in trip.correction_records):
            raise ValueError(f"人工动作 {record_id} 已应用，拒绝重复执行。")

        page = int(item.get("source_page", 0))
        amount_value = item.get("amount_cny")
        amount = (
            0.0 if action in {"ignore_fragment", "classify_supporting"} and amount_value is None
            else float(amount_value or -1)
        )
        category = _normalise_expense_category(item.get("category", "taxi")) or "taxi"
        if page < 1 or (
            action in {"add", "exclude", "include_with_risk"}
            and (amount <= 0 or not math.isfinite(amount))
        ):
            raise ValueError(f"人工动作 {record_id} 的页码或金额无效。")
        if category not in {"ticket", "taxi", "accommodation"}:
            raise ValueError(f"人工动作 {record_id} 的费用类别无效。")
        coverage = page_by_number.get(page)
        if coverage is None:
            raise ValueError(f"人工动作 {record_id} 的第 {page} 页不在当前文件页覆盖清单内。")

        service_date_value = item.get("service_date")
        service_date = date.fromisoformat(str(service_date_value)) if service_date_value else None
        service_date_end_value = item.get("service_date_to")
        service_date_end = (
            date.fromisoformat(str(service_date_end_value))
            if service_date_end_value else service_date
        )
        date_evidence_page = int(item.get("service_date_evidence_page") or page)
        reason = str(item.get("reason", "源 PDF 人工复核")).strip()
        evidence_location = str(item.get("evidence_location", f"源 PDF 第 {page} 页")).strip()
        before = {
            "expense_count": int(coverage.get("expense_count", 0)),
            "invoice_count": int(coverage.get("invoice_count", 0)),
            "amounts": list(coverage.get("amounts", [])),
        }
        action_status = "APPLIED"
        stable_id = hashlib.sha256(
            f"{actual_hash}|{record_id}|{page}|{amount_value if amount_value is not None else 'region'}".encode("utf-8")
        ).hexdigest()[:20]

        if action == "add":
            page_expenses = [expense for expense in trip.expenses if expense.source_page == page]
            page_invoices = [
                invoice for invoice in trip.structured_invoices
                if int(invoice.source.get("page", 0)) == page
            ]
            if page_expenses or page_invoices:
                related_record_id = str(
                    item.get("related_field_correction_record_id", "")
                ).strip()
                related_corrections = [
                    correction for correction in field_corrections
                    if isinstance(correction, dict)
                    and correction.get("record_id") == related_record_id
                ] if isinstance(field_corrections, list) else []
                if (
                    item.get("allow_additional_same_page") is not True
                    or not related_record_id
                    or len(related_corrections) != 1
                    or len(page_expenses) != 1
                    or len(page_invoices) != 1
                ):
                    raise ValueError(
                        f"人工补录 {record_id} 的第 {page} 页已有费用/票据记录；"
                        "拒绝重复或含糊补录。"
                    )
                related = related_corrections[0]
                invoice_match = related.get("invoice_match", {})
                expense_fields = related.get("expense_fields", {})
                structured_fields = related.get("structured_fields", {})
                if not all(isinstance(value, dict) for value in (
                    invoice_match,
                    expense_fields,
                    structured_fields,
                )):
                    raise ValueError(
                        f"人工补录 {record_id} 的关联字段修正结构无效。"
                    )
                expense_amount_correction = expense_fields.get("amount", {})
                invoice_amount_correction = structured_fields.get("total_amount", {})
                if not isinstance(expense_amount_correction, dict) or not isinstance(
                    invoice_amount_correction, dict
                ):
                    raise ValueError(
                        f"人工补录 {record_id} 的关联字段修正缺少原金额/确认金额。"
                    )
                try:
                    related_page = int(invoice_match.get("page", 0))
                    existing_amount = float(invoice_match.get("original_total_amount"))
                    expected_expense_amount = float(expense_amount_correction.get("expected_current"))
                    expected_invoice_amount = float(invoice_amount_correction.get("expected_current"))
                    corrected_expense_amount = float(expense_amount_correction.get("value"))
                    corrected_invoice_amount = float(invoice_amount_correction.get("value"))
                    current_page_invoice_amount = float(page_invoices[0].total_amount)
                except (TypeError, ValueError):
                    raise ValueError(
                        f"人工补录 {record_id} 的关联字段修正没有完整的原金额/确认金额。"
                    )
                corrected_service_date = (
                    related.get("expense_fields", {}).get("start", {}).get("value")
                )
                if (
                    related_page != page
                    or invoice_match.get("expense_type") != category
                    or page_expenses[0].category != category
                    or page_invoices[0].expense_type != category
                    or abs(page_expenses[0].amount - existing_amount) > 0.01
                    or abs(current_page_invoice_amount - existing_amount) > 0.01
                    or abs(expected_expense_amount - existing_amount) > 0.01
                    or abs(expected_invoice_amount - existing_amount) > 0.01
                    or abs(corrected_expense_amount - corrected_invoice_amount) > 0.01
                    or abs(amount - existing_amount) <= 0.01
                    or abs(amount - corrected_expense_amount) <= 0.01
                    or not corrected_service_date
                    or service_date is None
                    or str(service_date) != str(corrected_service_date)
                    or coverage.get("status") != "EXPENSES_FOUND"
                ):
                    raise ValueError(
                        f"人工补录 {record_id} 与第 {page} 页已有票据及其关联字段修正不匹配。"
                    )
            elif coverage.get("status") not in {"UNMATCHED_RECEIPT_CUE", "UNMATCHED_INVOICE_CUE"}:
                raise ValueError(
                    f"人工补录 {record_id} 的第 {page} 页状态为 {coverage.get('status')}，"
                    "不符合未匹配票据线索的补录前置条件。"
                )
            if service_date is None or service_date_end is None or service_date_end < service_date:
                raise ValueError(f"人工补录 {record_id} 必须提供已复核的 service_date。")

            description = str(item.get("description", "人工复核补录费用")).strip()
            evidence_text = str(item.get("evidence_text", description)).strip()
            expense = Expense(
                category=category,
                amount=round(amount, 2),
                start=service_date,
                end=service_date_end,
                description=description,
                source_file=trip.source_file,
                source_page=page,
                accommodation_nights=(
                    (service_date_end - service_date).days
                    if category == "accommodation" and service_date_end > service_date else None
                ),
                document_id=stable_id,
                service_date_source="MANUAL_SOURCE_REVIEW",
                service_date_evidence_page=date_evidence_page,
                service_date_evidence_text=evidence_text,
                accommodation_nights_source=(
                    "MANUAL_SOURCE_REVIEW" if category == "accommodation" else "UNCONFIRMED"
                ),
            )
            invoice = TravelInvoice(
                document_type="taxi_receipt" if category == "taxi" else "travel_expense",
                vendor=str(item.get("vendor", "")),
                invoice_number=str(item.get("invoice_number", "")),
                currency="CNY",
                total_amount=round(amount, 2),
                travel_date_from=service_date,
                travel_date_to=service_date_end,
                travel_date_source="MANUAL_SOURCE_REVIEW",
                traveler=trip.name,
                expense_type=category,
                source={
                    "source_file": trip.source_file,
                    "sha256": actual_hash,
                    "page": page,
                    "document_type": "taxi_receipt" if category == "taxi" else "travel_expense",
                    "extraction_method": "approved_manual_action",
                },
                document_id=stable_id,
            )
            invoice.evidence["total_amount"] = {
                "field": "total_amount",
                "value": round(amount, 2),
                "source_file": trip.source_file,
                "page": page,
                "bbox": None,
                "raw_text": evidence_text,
                "ocr_engine": "manual_source_review",
                "ocr_confidence": None,
                "validation_status": "MANUAL_CONFIRMED",
                "evidence_location": evidence_location,
            }
            invoice.evidence["travel_date_from"] = {
                "field": "travel_date_from",
                "value": service_date.isoformat(),
                "source_file": trip.source_file,
                "page": date_evidence_page,
                "bbox": None,
                "raw_text": evidence_text,
                "ocr_engine": "manual_source_review",
                "ocr_confidence": None,
                "validation_status": "MANUAL_CONFIRMED",
                "evidence_location": evidence_location,
            }
            invoice.evidence["travel_date_to"] = copy.deepcopy(invoice.evidence["travel_date_from"])
            invoice.evidence["travel_date_to"]["value"] = service_date_end.isoformat()
            invoice.evidence["travel_date_to"]["page"] = date_evidence_page
            invoice.source["service_date_evidence_page"] = date_evidence_page
            trip.expenses.append(expense)
            trip.structured_invoices.append(invoice)
            pages_with_additions.add(page)
            page_expenses_after_add = [
                row for row in trip.expenses if row.source_page == page
            ]
            page_invoices_after_add = [
                row for row in trip.structured_invoices
                if int(row.source.get("page", 0)) == page
            ]
            coverage.update({
                "status": "EXPENSES_FOUND",
                "expense_count": len(page_expenses_after_add),
                "invoice_count": len(page_invoices_after_add),
                "amounts": sorted(round(row.amount, 2) for row in page_expenses_after_add),
                "review_required": False,
                "reasons": list(coverage.get("reasons", [])) + [
                    f"已应用已审核人工补录 {record_id}：票面金额 ¥{amount:.2f}。"
                ],
            })
            clear_resolved_page_risks(coverage)
        elif action == "exclude":
            matches_expense = [
                expense for expense in trip.expenses
                if expense.source_page == page
                and expense.category == category
                and abs(expense.amount - amount) <= 0.01
            ]
            matches_invoice = [
                invoice for invoice in trip.structured_invoices
                if int(invoice.source.get("page", 0)) == page
                and invoice.expense_type == category
                and invoice.total_amount is not None
                and abs(invoice.total_amount - amount) <= 0.01
            ]
            excluded_rows = coverage.setdefault("excluded_out_of_scope_receipts", [])
            already_excluded = any(
                abs(float(row.get("amount", -1)) - amount) <= 0.01
                for row in excluded_rows
            )
            if len(matches_expense) == 1 and len(matches_invoice) == 1:
                trip.expenses.remove(matches_expense[0])
                trip.structured_invoices.remove(matches_invoice[0])
                excluded_rows.append({
                    "amount": round(amount, 2),
                    "city": str(item.get("receipt_city", "")),
                    "trip_city": str(item.get("trip_city", trip.city)),
                })
            elif not matches_expense and not matches_invoice and already_excluded:
                action_status = "CONFIRMED_ALREADY_EXCLUDED"
            else:
                raise ValueError(
                    f"人工排除 {record_id} 在第 {page} 页匹配到 "
                    f"{len(matches_expense)} 条 Expense、{len(matches_invoice)} 条票据，"
                    "且没有同金额的已排除记录；拒绝按金额猜测。"
                )
            coverage.update({
                "status": "SUPPORTING_ONLY",
                "expense_count": sum(expense.source_page == page for expense in trip.expenses),
                "invoice_count": sum(
                    int(invoice.source.get("page", 0)) == page
                    for invoice in trip.structured_invoices
                ),
                "amounts": sorted({
                    round(expense.amount, 2)
                    for expense in trip.expenses
                    if expense.source_page == page
                }),
                "review_required": False,
            })
            coverage.setdefault("reasons", []).append(
                f"已应用已审核人工排除 {record_id}：{reason}"
            )
        elif action == "include_with_risk":
            matches_expense = [
                expense for expense in trip.expenses
                if expense.source_page == page
                and expense.category == category
                and abs(expense.amount - amount) <= 0.01
            ]
            matches_invoice = [
                invoice for invoice in trip.structured_invoices
                if int(invoice.source.get("page", 0)) == page
                and invoice.expense_type == category
                and invoice.total_amount is not None
                and abs(invoice.total_amount - amount) <= 0.01
            ]
            service_date = date.fromisoformat(str(item.get("service_date", "")))
            service_date_end = date.fromisoformat(str(item.get("service_date_to", "")))
            allocation_anchor_value = str(item.get("allocation_anchor", "")).strip()
            allocation_anchor = date.fromisoformat(allocation_anchor_value)
            risk_code = str(item.get("risk_code", "")).strip()
            risk_description = str(item.get("risk_description", "")).strip()
            if len(matches_expense) != 1 or len(matches_invoice) != 1:
                raise ValueError(
                    f"风险计入 {record_id} 在第 {page} 页必须唯一匹配一条同类别、同金额 Expense 和发票；"
                    f"实际匹配 {len(matches_expense)} 条 Expense、{len(matches_invoice)} 张发票。"
                )
            if (
                category != "accommodation"
                or service_date_end <= service_date
                or allocation_anchor < trip.start
                or allocation_anchor > trip.end
                or not risk_code
                or not risk_description
                or int(item.get("service_date_evidence_page") or 0) < 1
            ):
                raise ValueError(f"风险计入 {record_id} 的住宿日期、展示日期或风险证据无效。")
            if coverage.get("status") != "EXPENSES_FOUND":
                raise ValueError(
                    f"风险计入 {record_id} 的第 {page} 页状态为 {coverage.get('status')}，"
                    "不符合已解析费用记录的前置条件。"
                )

            expense = matches_expense[0]
            invoice = matches_invoice[0]
            evidence_page = int(item["service_date_evidence_page"])
            evidence_text = str(item.get("evidence_text", "")).strip()
            expense.start = service_date
            expense.end = service_date_end
            expense.accommodation_nights = (service_date_end - service_date).days
            expense.accommodation_nights_source = "HOTEL_FOLIO"
            expense.service_date_source = "HOTEL_FOLIO"
            expense.service_date_evidence_page = evidence_page
            expense.service_date_evidence_text = evidence_text
            expense.allocation_anchor = allocation_anchor
            if item.get("vendor"):
                invoice.vendor = str(item["vendor"]).strip()
            invoice.travel_date_from = service_date
            invoice.travel_date_to = service_date_end
            invoice.travel_date_source = "HOTEL_FOLIO"
            invoice.source["service_date_evidence_page"] = evidence_page
            date_evidence = {
                "field": "travel_date_from",
                "value": service_date.isoformat(),
                "source_file": trip.source_file,
                "page": evidence_page,
                "bbox": None,
                "raw_text": evidence_text,
                "ocr_engine": "manual_source_review",
                "ocr_confidence": None,
                "validation_status": "MANUAL_CONFIRMED",
                "evidence_location": evidence_location,
            }
            invoice.evidence["travel_date_from"] = copy.deepcopy(date_evidence)
            invoice.evidence["travel_date_to"] = {
                **date_evidence,
                "field": "travel_date_to",
                "value": service_date_end.isoformat(),
            }

            risk = {
                "record_id": record_id,
                "code": risk_code,
                "severity": "WARNING",
                "source_page": page,
                "evidence_page": evidence_page,
                "amount_cny": round(amount, 2),
                "actual_service_date_from": service_date.isoformat(),
                "actual_service_date_to": service_date_end.isoformat(),
                "allocation_anchor": allocation_anchor.isoformat(),
                "description": risk_description,
                "user_decision": str(item.get("user_decision", "")).strip(),
                "reviewed_on": action_reviewed_on,
            }
            if not risk["user_decision"]:
                raise ValueError(f"风险计入 {record_id} 必须保留用户决定说明。")
            coverage.setdefault("risk_flags", []).append(copy.deepcopy(risk))
            coverage["review_required"] = False
            coverage.setdefault("reasons", []).append(risk_description)
            action_status = "INCLUDED_WITH_RISK"
        elif action == "classify_supporting":
            if page_expenses := [expense for expense in trip.expenses if expense.source_page == page]:
                raise ValueError(
                    f"辅助材料标记 {record_id} 的第 {page} 页已有 {len(page_expenses)} 条费用，拒绝覆盖。"
                )
            if page_invoices := [
                invoice for invoice in trip.structured_invoices
                if int(invoice.source.get("page", 0)) == page
            ]:
                raise ValueError(
                    f"辅助材料标记 {record_id} 的第 {page} 页已有 {len(page_invoices)} 张发票，拒绝覆盖。"
                )
            if coverage.get("status") not in {
                "UNMATCHED_RECEIPT_CUE", "UNMATCHED_INVOICE_CUE", "UNCLASSIFIED_PAGE",
                "DIDI_AMOUNT_MISMATCH",
            }:
                raise ValueError(
                    f"辅助材料标记 {record_id} 的第 {page} 页状态为 {coverage.get('status')}，"
                    "不符合人工分类前置条件。"
                )
            clear_resolved_page_risks(coverage)
            coverage.update({
                "status": "SUPPORTING_ONLY",
                "expense_count": 0,
                "invoice_count": 0,
                "amounts": [],
                "review_required": False,
                "reasons": [f"已确认本页为辅助材料，不作为单独费用：{reason}"],
            })
            action_status = "CONFIRMED_SUPPORTING_ONLY"
        else:
            decision = str(item.get("decision", "")).strip().lower()
            region = str(item.get("region", "")).strip()
            if decision != "user_confirmed_ignore" or not region:
                raise ValueError(
                    f"忽略残片 {record_id} 必须有用户确认 decision 和具体 region。"
                )
            if coverage.get("status") != "DIDI_AMOUNT_MISMATCH":
                raise ValueError(
                    f"忽略残片 {record_id} 的第 {page} 页状态为 {coverage.get('status')}，"
                    "与清单确认的 Didi 汇总差异不符。"
                )
            current_reasons = list(coverage.get("reasons", []))
            if len(current_reasons) != 1 or "Didi 行程汇总" not in current_reasons[0]:
                raise ValueError(f"忽略残片 {record_id} 发现额外页面异常，拒绝清除待复核状态。")
            coverage.setdefault("resolved_issues", []).append({
                "status": coverage.get("status"),
                "reasons": current_reasons,
                "resolution": "USER_CONFIRMED_IGNORE_INCOMPLETE_FRAGMENT",
                "record_id": record_id,
            })
            coverage.setdefault("ignored_regions", []).append({
                "region": region,
                "decision": decision,
                "record_id": record_id,
                "reason": reason,
            })
            coverage.update({
                "status": "EXPENSES_FOUND" if before["expense_count"] else "SUPPORTING_ONLY",
                "review_required": False,
                "reasons": [
                    f"按用户确认忽略不完整残片：{region}；本页可读费用记录保留。"
                ],
            })
            action_status = "CONFIRMED_USER_IGNORE"

        record = {
            "record_id": record_id,
            "action": action,
            "action_status": action_status,
            "source_sha256": actual_hash,
            "page": page,
            "category": category,
            "amount_cny": round(amount, 2) if amount_value is not None else None,
            "reviewed_on": action_reviewed_on,
            "evidence_location": evidence_location,
            "reason": reason,
            "before": before,
            "after": {
                "expense_count": int(coverage.get("expense_count", 0)),
                "invoice_count": int(coverage.get("invoice_count", 0)),
                "amounts": list(coverage.get("amounts", [])),
            },
        }
        if action in {"add", "include_with_risk"}:
            if action == "include_with_risk":
                record.update({
                    "risk_code": risk_code,
                    "risk_description": risk_description,
                    "allocation_anchor": allocation_anchor.isoformat(),
                    "actual_service_date_from": service_date.isoformat(),
                    "actual_service_date_to": service_date_end.isoformat(),
                    "user_decision": risk["user_decision"],
                })
            expense.correction_records.append(copy.deepcopy(record))
            invoice.validation = validate_travel_invoice(invoice)
            invoice.correction_records.append(copy.deepcopy(record))
        trip.correction_records.append(copy.deepcopy(record))
        applied.append(record)

    if field_corrections:
        if not isinstance(field_corrections, list):
            raise ValueError("人工动作清单 field_corrections 必须为列表。")
        applied.extend(apply_manual_travel_corrections(
            trip,
            {
                "source_sha256": actual_hash,
                "reviewed_on": reviewed_on,
                "corrections": field_corrections,
            },
        ))

    for page in pages_with_additions:
        coverage = page_by_number.get(page)
        if coverage is None:
            continue
        page_expenses = [expense for expense in trip.expenses if expense.source_page == page]
        page_invoices = [
            invoice for invoice in trip.structured_invoices
            if int(invoice.source.get("page", 0)) == page
        ]
        coverage.update({
            "status": "EXPENSES_FOUND" if page_expenses else "SUPPORTING_ONLY",
            "expense_count": len(page_expenses),
            "invoice_count": len(page_invoices),
            "amounts": sorted(round(expense.amount, 2) for expense in page_expenses),
            "review_required": False,
        })

    return applied


def _date_range(start: date, end: date) -> list[date]:
    return [start + timedelta(days=i) for i in range((end - start).days + 1)]


def _clamp_day(day: Optional[date], trip_days: list[date]) -> date:
    if not day:
        return trip_days[0]
    if day < trip_days[0]:
        return trip_days[0]
    if day > trip_days[-1]:
        return trip_days[-1]
    return day


def _formula_or_value(amounts: Iterable[float]) -> Optional[float | str]:
    values = [round(float(value), 2) for value in amounts if value is not None]
    if not values:
        return None
    if len(values) == 1:
        return values[0]
    return "=" + "+".join(f"{value:g}" for value in values)


def _prepare_accommodation_dates(trip: Trip) -> list[Expense]:
    """Prepare hotel spans from folio evidence before applying the rate cap."""
    accommodations = [expense for expense in trip.expenses if expense.category == "accommodation"]
    _infer_accommodation_spans(trip)
    return accommodations


def _accommodation_invoice_nights(expense: Expense) -> int:
    if expense.accommodation_nights is not None:
        return max(1, expense.accommodation_nights)
    if expense.start and expense.end:
        return max(1, (expense.end - expense.start).days)
    return 1


def _accommodation_nights(trip: Trip) -> int:
    accommodations = _prepare_accommodation_dates(trip)
    return sum(_accommodation_invoice_nights(expense) for expense in accommodations)


def _accommodation_limit(trip: Trip) -> float:
    accommodations = _prepare_accommodation_dates(trip)
    engine = _travel_policy_engine(trip.policy_name)
    return round(sum(
        engine.calculate_accommodation(
            trip.city_for(expense.start or trip.start),
            expense.amount,
            _accommodation_invoice_nights(expense),
        )["output"].get("allowed_amount", 0.0)
        for expense in accommodations
    ), 2)


def allocate_trip_expenses(trip: Trip) -> dict:
    """Allocate expenses to rows while retaining each invoice amount in formulas."""
    for expense in trip.expenses:
        expense.category = _normalise_expense_category(expense.category) or "taxi"
    ordered_expenses = [
        expense for _index, expense in sorted(
            enumerate(trip.expenses),
            key=lambda item: (
                item[1].source_page is None,
                item[1].source_page if item[1].source_page is not None else 0,
                item[0],
            ),
        )
    ]
    days = trip.days
    _prepare_accommodation_dates(trip)
    policy_engine = _travel_policy_engine(trip.policy_name)
    policy_trace: list[dict] = []
    policy_exceptions: list[dict] = []
    allocation_trace: list[dict] = []
    bucket: dict[str, dict[date, list[float]]] = {
        "ticket": defaultdict(list),
        "taxi": defaultdict(list),
        "accommodation": defaultdict(list),
    }
    special: list[Expense] = []
    special_by_date: dict[date, list[float]] = defaultdict(list)
    special_notes_by_date: dict[date, list[str]] = defaultdict(list)
    # Keep a separate occupancy count for each expense category. A workbook
    # row has one cell per category, so an empty date cell is available even
    # when another category already uses that same row.
    category_counts: dict[str, dict[date, int]] = {
        category: defaultdict(int) for category in bucket
    }
    accommodation_notes: dict[date, list[str]] = defaultdict(list)

    processed_groups: set[str] = set()
    for expense in ordered_expenses:
        if expense.category == "special":
            special.append(expense)
            service_date_confirmed = _has_confirmed_service_date(expense)
            if service_date_confirmed and expense.start:
                target = _clamp_day(expense.start, days)
                display_date_source = "SERVICE_DATE"
            elif expense.allocation_anchor:
                target = _clamp_day(expense.allocation_anchor, days)
                display_date_source = "ALLOCATION_ANCHOR"
            else:
                target = trip.end
                display_date_source = "ROW_PLACEMENT_ONLY"
            special_by_date[target].append(expense.amount)
            page_label = f"Page {expense.source_page}: " if expense.source_page else ""
            special_notes_by_date[target].append(
                f"{page_label}special expense {expense.amount:.2f} - {expense.description}"
            )
            allocation_trace.append({
                "document_id": expense.document_id,
                "source_file": expense.source_file,
                "source_page": expense.source_page,
                "category": expense.category,
                "gross_amount": round(expense.amount, 2),
                "allowed_amount": round(expense.amount, 2),
                "actual_service_start": expense.start.isoformat() if service_date_confirmed and expense.start else None,
                "actual_service_end": expense.end.isoformat() if service_date_confirmed and expense.end else None,
                "service_date_source": expense.service_date_source,
                "date_candidate_start": expense.start.isoformat() if expense.start else None,
                "date_candidate_end": expense.end.isoformat() if expense.end else None,
                "display_date": target.isoformat(),
                "display_date_source": display_date_source,
                "accommodation_nights": None,
                "accommodation_nights_source": None,
            })
            continue
        if expense.category not in bucket:
            continue
        if expense.allocation_splits:
            split_total = round(sum(float(split["amount"]) for split in expense.allocation_splits), 2)
            if split_total != round(expense.amount, 2):
                raise ValueError(
                    f"第 {expense.source_page} 页发票分配明细合计 {split_total:.2f} "
                    f"与票面总额 {expense.amount:.2f} 不符。"
                )
            for component_index, split in enumerate(expense.allocation_splits, 1):
                target = date.fromisoformat(str(split["date"]))
                if target not in days:
                    raise ValueError(
                        f"第 {expense.source_page} 页发票分配日期 {target.isoformat()} 超出行程范围。"
                    )
                amount = round(float(split["amount"]), 2)
                bucket[expense.category][target].append(amount)
                category_counts[expense.category][target] += 1
                allocation_trace.append({
                    "document_id": expense.document_id,
                    "source_file": expense.source_file,
                    "source_page": expense.source_page,
                    "category": expense.category,
                    "gross_amount": amount,
                    "invoice_gross_amount": round(expense.amount, 2),
                    "allocation_component": component_index,
                    "allowed_amount": amount,
                    "actual_service_start": target.isoformat(),
                    "actual_service_end": target.isoformat(),
                    "service_date_source": expense.service_date_source,
                    "service_date_evidence_page": expense.service_date_evidence_page,
                    "service_date_evidence_text": expense.service_date_evidence_text,
                    "date_candidate_start": target.isoformat(),
                    "date_candidate_end": target.isoformat(),
                    "display_date": target.isoformat(),
                    "display_date_source": "INVOICE_SUPPORTED_ALLOCATION",
                    "accommodation_nights": None,
                    "accommodation_nights_source": None,
                })
            continue
        group_name = expense.allocation_group
        if group_name:
            if group_name in processed_groups:
                continue
            group_expenses = [
                member for member in ordered_expenses
                if member.category == expense.category
                and member.allocation_group == group_name
            ]
            confirmed_dates = {
                member.start for member in group_expenses
                if _has_confirmed_service_date(member)
            }
            anchors = {
                member.allocation_anchor for member in group_expenses
                if member.allocation_anchor is not None
            }
            group_conflicts = (
                len(confirmed_dates) > 1
                or any(day not in confirmed_dates for day in anchors if confirmed_dates)
            )
            if group_conflicts:
                # A shared OCR/review group must never move expenses with
                # different confirmed dates onto its first member's date.
                group_expenses = [expense]
            else:
                processed_groups.add(group_name)
        else:
            group_expenses = [expense]

        representative = group_expenses[0]
        item_start = _clamp_day(representative.start, days)
        item_end = _clamp_day(representative.end or item_start, days)
        if item_end < item_start:
            item_end = item_start
        allocation_anchor = next(
            (member.allocation_anchor for member in group_expenses if member.allocation_anchor),
            None,
        )
        if _has_confirmed_service_date(representative) and representative.start is not None:
            # The date on the individual invoice is stronger than an aggregate
            # or itinerary allocation anchor.
            target = _clamp_day(representative.start, days)
        elif allocation_anchor is not None:
            # A reviewed aggregate (for example a trip-table total plus a
            # separately scanned receipt) is only a fallback when no member
            # carries a confirmed individual service date.
            target = _clamp_day(allocation_anchor, days)
        elif representative.start is None and representative.end is None:
            # Unknown dates are display positions only. Put them into the
            # next unused day in PDF order, then use the least-loaded cell
            # only when every day already has an entry.
            available = [
                day for day in days
                if category_counts[representative.category][day] == 0
            ]
            target = available[0] if available else min(
                days,
                key=lambda value: (category_counts[representative.category][value], value),
            )
        elif representative.category == "taxi":
            # A taxi total with a service range but no daily breakdown is one
            # invoice; keep it on the first day of that range instead of
            # spreading it to a day chosen by cell occupancy.
            target = item_start
        else:
            # Tickets and hotel invoices are recorded on the first/earliest
            # applicable date. If several invoices explicitly belong to that
            # same date, they must share that date's cell and become a
            # formula; moving them to another date would misstate the claim.
            candidates = [item_start]
            target = min(
                candidates,
                key=lambda value: (category_counts[representative.category][value], value),
            )
        declared_total = None
        if len(group_expenses) > 1:
            declared_totals = {
                round(float(member.allocation_group_total), 2)
                for member in group_expenses
                if member.allocation_group_total is not None
            }
            component_total = round(sum(member.amount for member in group_expenses), 2)
            uses_written_total = any(
                member.allocation_group_uses_written_total
                for member in group_expenses
            )
            if (
                len(declared_totals) == 1
                and (
                    abs(next(iter(declared_totals)) - component_total) <= 0.01
                    or (
                        uses_written_total
                        and next(iter(declared_totals)) >= component_total - 0.01
                    )
                )
            ):
                declared_total = next(iter(declared_totals))
        component_total = round(sum(member.amount for member in group_expenses), 2)
        written_total_adjustment = (
            round(declared_total - component_total, 2)
            if declared_total is not None and declared_total > component_total + 0.01
            else 0.0
        )
        category_counts[representative.category][target] += len(group_expenses)
        for member in group_expenses:
            amount = member.amount
            if member.category == "accommodation":
                nights = _accommodation_invoice_nights(member)
                stay_city = trip.city_for(member.start or target)
                decision = policy_engine.calculate_accommodation(stay_city, member.amount, nights)
                decision["context"] = {
                    "category": member.category,
                    "description": member.description,
                    "source_file": member.source_file,
                    "source_page": member.source_page,
                    "allocation_date": target.isoformat(),
                    "service_date_source": member.service_date_source,
                    "service_date_evidence_page": member.service_date_evidence_page,
                    "service_date_evidence_text": member.service_date_evidence_text,
                    "accommodation_nights": _accommodation_invoice_nights(member),
                    "accommodation_nights_source": member.accommodation_nights_source,
                }
                if member.accommodation_nights_source not in {
                    "HOTEL_FOLIO", "HOTEL_FOLIO_RATE", "SOURCE_TEXT_RANGE", "MANUAL_SOURCE_REVIEW",
                }:
                    nights_issue = {
                        "code": "UNCONFIRMED_ACCOMMODATION_NIGHTS",
                        "severity": "WARNING",
                        "message": "住宿晚数不是由票面日期或匹配水单确认，差标金额需复核",
                    }
                    decision.setdefault("issues", []).append(nights_issue)
                    if decision.get("status") == "PASS":
                        decision["status"] = "WARNING"
                policy_trace.append(decision)
                for issue in decision["issues"]:
                    policy_exceptions.append({
                        **issue,
                        "rule_id": decision["rule_id"],
                        "source_file": member.source_file,
                        "source_page": member.source_page,
                        "description": member.description,
                    })
                amount = decision["output"].get("allowed_amount", 0.0)
                if amount < member.amount - 0.01:
                    folio_confirmation = ""
                    if member.accommodation_folio_total is not None:
                        folio_confirmation = (
                            f"，水单 {member.accommodation_folio_total:.2f} 元确认 {nights} 晚，"
                            f"发票折合 {member.amount / nights:.2f} 元/晚"
                        )
                    accommodation_notes[target].append(
                        f"住宿发票 {member.amount:.2f} 元（{nights} 晚{folio_confirmation}）超过 "
                        f"{_policy_level_key(stay_city, trip.policy_name)} 层城市标准 "
                        f"{_accommodation_rate(stay_city, trip.policy_name):.0f} 元/晚，已按 {amount:.2f} 元记账。"
                    )
            if declared_total is None:
                bucket[member.category][target].append(amount)
            service_date_confirmed = _has_confirmed_service_date(member)
            allocation_trace.append({
                "document_id": member.document_id,
                "source_file": member.source_file,
                "source_page": member.source_page,
                "category": member.category,
                "gross_amount": round(member.amount, 2),
                "allowed_amount": round(amount, 2),
                "actual_service_start": member.start.isoformat() if service_date_confirmed and member.start else None,
                "actual_service_end": member.end.isoformat() if service_date_confirmed and member.end else None,
                "service_date_source": member.service_date_source,
                "allocation_group_total": member.allocation_group_total,
                "allocation_group_uses_written_total": member.allocation_group_uses_written_total,
                "written_total_adjustment": (
                    written_total_adjustment if member is representative else 0.0
                ),
                "service_date_evidence_page": (
                    member.service_date_evidence_page if service_date_confirmed else None
                ),
                "service_date_evidence_text": (
                    member.service_date_evidence_text if service_date_confirmed else ""
                ),
                "date_candidate_start": member.start.isoformat() if member.start else None,
                "date_candidate_end": member.end.isoformat() if member.end else None,
                "display_date": target.isoformat(),
                "display_date_source": (
                    "SERVICE_DATE" if service_date_confirmed else
                    "ALLOCATION_ANCHOR" if allocation_anchor is not None else
                    "ROW_PLACEMENT_ONLY"
                ),
                "accommodation_nights": (
                    _accommodation_invoice_nights(member)
                    if member.category == "accommodation" else None
                ),
                "accommodation_nights_source": (
                    member.accommodation_nights_source
                    if member.category == "accommodation" else None
                ),
            })
        if declared_total is not None:
            # The PDF's verified written total is the cell value; individual
            # component amounts remain in the audit trace above.
            bucket[representative.category][target].append(declared_total)

    special_total = _formula_or_value(expense.amount for expense in special)
    special_note = "；".join(sorted({expense.category + ": " + expense.description for expense in special}))
    trip.policy_trace = policy_trace
    trip.policy_exceptions = policy_exceptions
    trip.allocation_trace = allocation_trace
    return {
        "ticket": bucket["ticket"],
        "taxi": bucket["taxi"],
        "accommodation": bucket["accommodation"],
        "special_by_date": dict(special_by_date),
        "special_notes_by_date": dict(special_notes_by_date),
        "special_total": special_total,
        "accommodation_note": "；".join(
            note for notes in accommodation_notes.values() for note in notes
        ) or None,
        "accommodation_notes": dict(accommodation_notes),
        "special_note": special_note or "特殊费用",
        "policy_trace": policy_trace,
        "policy_exceptions": policy_exceptions,
        "allocation_trace": allocation_trace,
    }


def _copy_row_style(ws, source_row: int, target_row: int) -> None:
    if source_row < 1:
        return
    for col in range(1, ws.max_column + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)
        if source.has_style:
            target._style = copy.copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.alignment:
            target.alignment = copy.copy(source.alignment)
        if source.protection:
            target.protection = copy.copy(source.protection)
    source_dimension = ws.row_dimensions[source_row]
    target_dimension = ws.row_dimensions[target_row]
    target_dimension.height = source_dimension.height
    target_dimension.hidden = source_dimension.hidden


def _clear_data_cells(ws, row: int) -> None:
    for col in range(2, 17):
        ws.cell(row, col).value = None
        ws.cell(row, col).comment = None


def _find_summary_row(ws) -> Optional[int]:
    for row in range(2, ws.max_row + 1):
        value = ws.cell(row, 15).value
        if isinstance(value, str) and "SUMIF" in value.upper():
            return row
    return None


def _last_data_row(ws, summary_row: int) -> int:
    last = 1
    # Ignore lookup/formula-only cells in I/N so preformatted capacity rows
    # remain available for new trips.
    for row in range(2, summary_row):
        # B:H, K:M and O:P are user-entered/summary data. I and N contain
        # prefilled lookup formulas in otherwise empty capacity rows and must
        # not make the capacity look occupied.
        data_columns = list(range(2, 9)) + list(range(11, 14)) + [15, 16]
        if any(ws.cell(row, col).value is not None for col in data_columns):
            last = row
    return last


def _find_existing_trip_block(
    ws,
    trip: Trip,
    summary_row: int,
    reserved_ranges: Iterable[tuple[int, int]] = (),
) -> tuple[Optional[int], int]:
    """Find one exact person/city/date block to refresh in the output copy."""
    days = trip.days
    if not days or any(expense.category == "special" for expense in trip.expenses):
        return None, 0

    def identity(value) -> str:
        return re.sub(r"\s+", " ", str(value or "")).strip().casefold()

    def cell_date(value) -> Optional[date]:
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        if isinstance(value, str):
            try:
                return date.fromisoformat(value.strip()[:10])
            except ValueError:
                return None
        return None

    normalized_name = identity(trip.name)
    normalized_city = _normalise_city(trip.city).casefold()
    reserved = list(reserved_ranges)
    matches = []
    last_start = summary_row - len(days)
    for start_row in range(2, last_start + 1):
        end_row = start_row + len(days) - 1
        if any(start_row <= reserved_end and reserved_start <= end_row
               for reserved_start, reserved_end in reserved):
            continue
        exact_match = True
        for offset, expected_day in enumerate(days):
            row = start_row + offset
            if (
                cell_date(ws.cell(row, 3).value) != expected_day
                or identity(ws.cell(row, 5).value) != normalized_name
                or _normalise_city(str(ws.cell(row, 8).value or "")).casefold() != normalized_city
            ):
                exact_match = False
                break
        if exact_match:
            matches.append(start_row)
    if len(matches) == 1:
        return matches[0], 1
    return None, len(matches)


def _refresh_summary_formulas(ws) -> None:
    summary_row = _find_summary_row(ws)
    if not summary_row:
        return
    end_row = summary_row - 1
    ws.cell(summary_row, 15).value = f"=SUMIF(D2:P{end_row}, Q{summary_row}, O2:O{end_row})"
    ws.cell(summary_row, 16).value = f"=SUMIF(D3:P{end_row}, Q{summary_row}, P3:P{end_row})"
    if summary_row + 1 <= ws.max_row and ws.cell(summary_row + 1, 15).value is not None:
        ws.cell(summary_row + 1, 15).value = f"=39240-O{summary_row}"


def _copy_formula_rows(ws, rows: Iterable[int], city: str, first_row: int, policy_name: str = "VGC") -> None:
    for row in rows:
        if policy_name == "VGC" and _is_defined_policy_city(city, policy_name):
            ws.cell(row, 9).value = f'=IFERROR(VLOOKUP(H{row},$V$2:$X$52,2,0),{_city_level(city, policy_name):g})'
        else:
            ws.cell(row, 9).value = f'={_city_level(city, policy_name):g}'
        if row == first_row:
            ws.cell(row, 14).value = None
        elif policy_name == "VGC" and _is_defined_policy_city(city, policy_name):
            ws.cell(row, 14).value = f'=IFERROR(VLOOKUP(H{row},$V$2:$Y$52,4,0),{_allowance_rate(city, policy_name):g})'
        else:
            ws.cell(row, 14).value = f'={_allowance_rate(city, policy_name):g}'


def _write_trip(ws, trip: Trip, start_row: int, po_no, style_source_row: int) -> tuple[int, float]:
    allocation = allocate_trip_expenses(trip)
    days = trip.days
    policy_engine = _travel_policy_engine(trip.policy_name)
    allowance_decisions = []
    city_order = list(dict.fromkeys(trip.city_for(day) for day in days if trip.city_for(day)))
    for city in city_order or [trip.city]:
        eligible_days = sum(1 for day in days[1:] if trip.city_for(day) == city)
        allowance_decision = policy_engine.calculate_allowance(
            city,
            trip_days=len(days),
            allowance_days=eligible_days,
        )
        allowance_decisions.append(allowance_decision)
        trip.policy_trace.append(allowance_decision)
        for issue in allowance_decision["issues"]:
            trip.policy_exceptions.append({
                **issue,
                "rule_id": allowance_decision["rule_id"],
                "source_file": trip.source_file,
                "source_page": None,
                "description": f"行程日补贴（{city}）",
            })
    allowance_summary = {
        "allowed_amount": round(sum(
            decision["output"].get("allowed_amount", 0.0)
            for decision in allowance_decisions
        ), 2),
        "eligible_days": sum(decision["input"].get("allowance_days", 0) for decision in allowance_decisions),
        "by_city": [
            {
                "city": decision["input"]["city"],
                "eligible_days": decision["input"].get("allowance_days", 0),
                "allowance_per_day": decision["output"].get("allowance_per_day", 0.0),
                "allowed_amount": decision["output"].get("allowed_amount", 0.0),
            }
            for decision in allowance_decisions
        ],
    }
    accommodation_claimed = round(
        sum(expense.amount for expense in trip.expenses if expense.category == "accommodation"), 2
    )
    accommodation_allowed = round(
        sum(
            item["output"].get("allowed_amount", 0.0)
            for item in allocation["policy_trace"]
            if item.get("rule_id") == "TRAVEL.HOTEL_CAP"
        ),
        2,
    )
    accommodation_over_limit = round(max(accommodation_claimed - accommodation_allowed, 0.0), 2)
    policy_status = "PASS"
    for decision in trip.policy_trace:
        if decision.get("status") == "BLOCKED":
            policy_status = "BLOCKED"
            break
        if decision.get("status") == "WARNING":
            policy_status = "WARNING"
    trip.policy_summary = {
        "status": policy_status,
        "rule_count": len(trip.policy_trace),
        "exception_count": len(trip.policy_exceptions),
        "accommodation": {
            "claimed_amount": accommodation_claimed,
            "allowed_amount": accommodation_allowed,
            "over_limit": accommodation_over_limit,
        },
        "allowance": allowance_summary,
        "policy_eligible_amount": round(
            accommodation_allowed + allowance_summary["allowed_amount"], 2
        ),
    }
    special_total = allocation["special_total"]
    accommodation_notes = allocation["accommodation_notes"]
    special_summary_row = start_row + len(days) if special_total is not None else None
    final_row = special_summary_row or (start_row + len(days) - 1)
    style_source_row = max(2, style_source_row)

    for offset, day in enumerate(days):
        row = start_row + offset
        day_city = trip.city_for(day)
        _copy_row_style(ws, style_source_row, row)
        _clear_data_cells(ws, row)
        ws.cell(row, 2).value = day.strftime("%b")
        ws.cell(row, 3).value = datetime.combine(day, datetime.min.time())
        ws.cell(row, 4).value = po_no
        ws.cell(row, 5).value = trip.name
        ws.cell(row, 6).value = None
        ws.cell(row, 7).value = None
        ws.cell(row, 8).value = day_city
        ws.cell(row, 10).value = None
        ws.cell(row, 11).value = _formula_or_value(allocation["ticket"].get(day, []))
        ws.cell(row, 12).value = _formula_or_value(
            allocation["taxi"].get(day, [])
            + allocation["special_by_date"].get(day, [])
        )
        ws.cell(row, 13).value = _formula_or_value(allocation["accommodation"].get(day, []))
        special_notes = allocation["special_notes_by_date"].get(day, [])
        if special_notes:
            ws.cell(row, 7).value = "；".join(special_notes)
        day_notes = accommodation_notes.get(day, [])
        if day_notes:
            note = "；".join(day_notes)
            ws.cell(row, 13).comment = Comment(note, "IAV AutoTools")
            for item in day_notes:
                if item not in trip.warnings:
                    trip.warnings.append(item)
        _copy_formula_rows(ws, [row], day_city, start_row, trip.policy_name)

    if special_summary_row is not None:
        row = special_summary_row
        _copy_row_style(ws, style_source_row, row)
        _clear_data_cells(ws, row)

    category_columns = {"ticket": 11, "taxi": 12, "accommodation": 13, "special": 12}
    for item in trip.allocation_trace:
        display_date = date.fromisoformat(item["display_date"])
        item_row = start_row + (display_date - trip.start).days
        item["workbook_row"] = item_row
        item["workbook_cell"] = f"{get_column_letter(category_columns[item['category']])}{item_row}"

    sum_end_row = final_row - 1 if special_summary_row is not None else final_row
    ws.cell(final_row, 15).value = f"=SUM(K{start_row}:N{sum_end_row})"
    ws.cell(final_row, 16).value = f"=O{final_row}*1.06"
    # Return the expected numeric total because openpyxl does not calculate formulas.
    numeric_total = sum(
        sum(values) for category in ("ticket", "taxi", "accommodation") for values in allocation[category].values()
    )
    numeric_total += sum(expense.amount for expense in trip.expenses if expense.category == "special")
    # First day allowance is intentionally deleted; subsequent days use the
    # city lookup value. The workbook lookup table is consulted by Excel.
    allowance = 0.0
    for day in days[1:]:
        allowance += _lookup_allowance(ws, trip.city_for(day), trip.policy_name)
    numeric_total += allowance
    expected_cells = {}
    for offset, day in enumerate(days):
        row = start_row + offset
        for category, column in (("ticket", 11), ("accommodation", 13)):
            expected_cells[f"{get_column_letter(column)}{row}"] = _formula_or_value(
                allocation[category].get(day, [])
            )
        expected_cells[f"L{row}"] = _formula_or_value(
            allocation["taxi"].get(day, [])
            + allocation["special_by_date"].get(day, [])
        )
    cell_mismatches = [
        {"cell": cell, "expected": expected, "actual": ws[cell].value}
        for cell, expected in expected_cells.items()
        if ws[cell].value != expected
    ]
    written_summary_adjustment_total = round(sum(
        float(item.get("written_total_adjustment", 0.0) or 0.0)
        for item in trip.allocation_trace
    ), 2)
    trace_total = round(sum(
        float(item["allowed_amount"])
        + float(item.get("written_total_adjustment", 0.0) or 0.0)
        for item in trip.allocation_trace
    ), 2)
    supported_expense_count = sum(
        expense.category in {"ticket", "taxi", "accommodation", "special"}
        for expense in trip.expenses
    )
    expected_allocation_component_count = sum(
        len(expense.allocation_splits) if expense.allocation_splits else 1
        for expense in trip.expenses
        if expense.category in {"ticket", "taxi", "accommodation", "special"}
    )
    trace_cell_totals: dict[str, float] = defaultdict(float)
    for item in trip.allocation_trace:
        if item.get("workbook_cell"):
            trace_cell_totals[item["workbook_cell"]] += (
                float(item["allowed_amount"])
                + float(item.get("written_total_adjustment", 0.0) or 0.0)
            )
    expected_trace_cell_totals: dict[str, float] = defaultdict(float)
    for offset, day in enumerate(days):
        row = start_row + offset
        for category, column in (("ticket", 11), ("accommodation", 13)):
            cell = f"{get_column_letter(column)}{row}"
            expected_trace_cell_totals[cell] += sum(allocation[category].get(day, []))
        expected_trace_cell_totals[f"L{row}"] += (
            sum(allocation["taxi"].get(day, []))
            + sum(allocation["special_by_date"].get(day, []))
        )
    trace_cell_mismatches = [
        {
            "cell": cell,
            "expected": round(expected_trace_cell_totals.get(cell, 0.0), 2),
            "actual": round(trace_cell_totals.get(cell, 0.0), 2),
        }
        for cell in sorted(set(trace_cell_totals) | set(expected_trace_cell_totals))
        if round(trace_cell_totals.get(cell, 0.0), 2) != round(expected_trace_cell_totals.get(cell, 0.0), 2)
    ]
    allowance_total = round(allowance, 2)
    reconciled_total = round(trace_total + allowance_total, 2)
    expected_o_formula = f"=SUM(K{start_row}:N{sum_end_row})"
    expected_p_formula = f"=O{final_row}*1.06"
    formula_mismatches = []
    if ws.cell(final_row, 15).value != expected_o_formula:
        formula_mismatches.append({"cell": f"O{final_row}", "expected": expected_o_formula, "actual": ws.cell(final_row, 15).value})
    if ws.cell(final_row, 16).value != expected_p_formula:
        formula_mismatches.append({"cell": f"P{final_row}", "expected": expected_p_formula, "actual": ws.cell(final_row, 16).value})
    trip.allocation_audit = {
        "status": "PASS" if (
            not cell_mismatches
            and not trace_cell_mismatches
            and not formula_mismatches
            and len(trip.allocation_trace) == expected_allocation_component_count
            and reconciled_total == round(numeric_total, 2)
        ) else "BLOCKED",
        "expense_count": len(trip.expenses),
        "trace_count": len(trip.allocation_trace),
        "supported_expense_count": supported_expense_count,
        "expected_allocation_component_count": expected_allocation_component_count,
        "cells_checked": len(expected_cells),
        "cell_mismatches": cell_mismatches,
        "trace_cell_mismatches": trace_cell_mismatches,
        "formula_mismatches": formula_mismatches,
        "trace_allowed_amount_total": trace_total,
        "written_summary_adjustment_total": written_summary_adjustment_total,
        "allowance_total": allowance_total,
        "reconciled_total": reconciled_total,
        "expected_sum_without_tax": round(numeric_total, 2),
        "expected_sum_with_tax": numeric_total * 1.06,
        "formula_cache_expectations": [
            {
                "sheet": ws.title,
                "cell": f"O{final_row}",
                "formula": expected_o_formula,
                "value": numeric_total,
                "tolerance": 0.005,
            },
            {
                "sheet": ws.title,
                "cell": f"P{final_row}",
                "formula": expected_p_formula,
                "value": numeric_total * 1.06,
                "tolerance": 0.005,
            },
        ],
        "formula_recalculation": "PENDING_EXCEL_RECALCULATION",
    }
    return final_row, round(numeric_total, 2)


def _lookup_allowance(ws, city: str, policy_name: str = "VGC") -> float:
    return _allowance_rate(city, policy_name)


def _derive_output_path(template_path: str, output_date: Optional[str], output_dir: Optional[str] = None) -> str:
    stamp = output_date or datetime.now().strftime("%y%m%d")
    if not re.fullmatch(r"\d{6}", stamp):
        raise ValueError("output_date 必须是 YYMMDD 六位日期")
    source = Path(template_path)
    # _vN is a collision suffix, not part of the template's date or basename.
    base_name = re.sub(r"_v\d+(?=\.xlsx$)", "", source.name, flags=re.IGNORECASE)
    name, date_replacements = re.subn(
        r"\d{6}(?=\.xlsx$)", stamp, base_name, count=1, flags=re.IGNORECASE,
    )
    if not date_replacements:
        name = f"{Path(base_name).stem}_{stamp}{source.suffix}"
    return str(Path(output_dir or source.parent) / name)


def assess_trip_delivery(trip: Trip) -> dict:
    """Decide whether a generated trip is ready or must remain a review draft."""
    invoice_validation = summarise_travel_validation(trip.structured_invoices)
    review_pages = [
        page for page in trip.page_coverage
        if page.get("review_required") is True
    ]
    policy_status = str((trip.policy_summary or {}).get("status", "PASS")).upper()
    issues = []
    if not trip.structured_invoices:
        issues.append({
            "code": "NO_STRUCTURED_INVOICES",
            "message": "没有可追溯的逐票记录",
        })
    if invoice_validation["review_required"]:
        issues.append({
            "code": "INVOICE_VALIDATION_REVIEW",
            "message": f"{invoice_validation['review_required']} 条票据存在关键字段或金额校验问题",
        })
    if review_pages:
        page_summaries = []
        for page in review_pages[:4]:
            details = "；".join(page.get("reasons", [])) or str(page.get("status", "需复核"))
            page_summaries.append(f"第 {page.get('source_page')} 页：{details}")
        remaining_pages = len(review_pages) - len(page_summaries)
        review_summary = "；".join(page_summaries)
        if remaining_pages > 0:
            review_summary += f"；另有 {remaining_pages} 页待复核"
        issues.append({
            "code": "PAGE_COVERAGE_REVIEW",
            "message": f"{len(review_pages)} 页仍需确认票据覆盖情况：{review_summary}",
        })
    if policy_status in {"WARNING", "BLOCKED"}:
        issues.append({
            "code": "TRAVEL_POLICY_REVIEW",
            "message": f"差旅政策状态为 {policy_status}",
        })
    allocation_audit = trip.allocation_audit or {}
    if allocation_audit.get("status") not in {None, "PASS"}:
        issues.append({
            "code": "ALLOCATION_RECONCILIATION_FAILED",
            "message": "费用追踪记录、写入单元格或包级金额未通过对账",
        })
    formula_recalculation = allocation_audit.get("formula_recalculation")
    if allocation_audit.get("status") == "PASS" and formula_recalculation == "BLOCKED":
        issues.append({
            "code": "EXCEL_RECALCULATION_FAILED",
            "message": "Excel 公式缓存存在错误或与独立期望金额不符",
        })
    elif allocation_audit.get("status") == "PASS" and formula_recalculation != "PASS":
        issues.append({
            "code": "EXCEL_RECALCULATION_PENDING",
            "message": "公式已写入但尚无 Excel 实际重算结果，不能标记为可直接使用",
        })
    status = "REVIEW_DRAFT" if issues else "READY_FOR_USE"
    return {
        "delivery_status": status,
        "ready_for_use": status == "READY_FOR_USE",
        "review_required": bool(issues),
        "delivery_issues": issues,
        "invoice_validation": invoice_validation,
        "review_page_count": len(review_pages),
        "policy_status": policy_status,
        "allocation_audit": copy.deepcopy(allocation_audit),
        "allocation_trace": copy.deepcopy(trip.allocation_trace),
    }


def create_review_workbook(
    template_path: str,
    trips: list[Trip],
    output_date: Optional[str] = None,
    output_path: Optional[str] = None,
    po_no=None,
    policy_name: str = "VGC",
    claim_workbook_path: Optional[str] = None,
) -> dict:
    """Copy the template, refresh exact matching trips, and append new blocks.

    ``claim_workbook_path`` optionally points at an expense-claim export.  The
    claim is used as a reconciliation baseline and adds a ``报销核对`` sheet to
    the generated copy; the template itself is never modified.
    """
    template = os.path.abspath(template_path)
    if not os.path.isfile(template):
        raise FileNotFoundError(f"模板不存在: {template}")
    if not trips:
        raise ValueError("没有可写入的差旅材料")
    template_adapter = TravelExpenseTemplateAdapter()
    template_validation = template_adapter.validate_template(template)
    if not template_validation.ok:
        raise ValueError(
            "差旅模板校验失败: " + "；".join(template_validation.errors)
        )
    if policy_name not in load_travel_policies().get("policies", {}):
        raise ValueError(f"差标不存在: {policy_name}")
    for trip in trips:
        trip.policy_name = policy_name

    destination = os.path.abspath(output_path or _derive_output_path(template, output_date))
    if destination == template:
        raise ValueError("输出文件不能覆盖原始模板")
    os.makedirs(os.path.dirname(destination), exist_ok=True)
    # Do not copy the template's filesystem metadata to the output. In
    # particular, copy2 preserves the Windows read-only attribute; openpyxl
    # then fails when saving the generated workbook even though the output
    # directory itself is writable.
    shutil.copyfile(template, destination)

    keep_vba = Path(template).suffix.lower() == ".xlsm"
    workbook = load_workbook(destination, data_only=False, keep_vba=keep_vba)
    worksheet = workbook.worksheets[0]
    summary_row = _find_summary_row(worksheet) or (worksheet.max_row + 1)
    cursor = _last_data_row(worksheet, summary_row)
    style_source = max(2, cursor)
    inferred_po = po_no
    if inferred_po in (None, ""):
        for row in range(cursor, 1, -1):
            value = worksheet.cell(row, 4).value
            if value not in (None, ""):
                inferred_po = value
                break

    written: list[dict] = []
    insertion_log: list[dict[str, int]] = []
    writable_row_ranges: list[tuple[int, int]] = []
    reserved_trip_ranges: list[tuple[int, int]] = []
    original_summary_row = template_validation.metadata.get("summary_row")
    for trip in trips:
        special_needed = any(expense.category == "special" for expense in trip.expenses)
        block_rows = len(trip.days) + (1 if special_needed else 0)
        existing_start, existing_match_count = _find_existing_trip_block(
            worksheet, trip, summary_row, reserved_trip_ranges,
        )
        if existing_start is not None:
            start_row = existing_start
            write_mode = "REPLACED_EXACT_EXISTING_TRIP"
            trip_style_source = max(2, existing_start)
        else:
            start_row = cursor + 2  # keep one blank row after the previous block
            write_mode = (
                "APPENDED_MULTIPLE_TEMPLATE_MATCHES"
                if existing_match_count > 1 else "APPENDED_NEW_TRIP"
            )
            if existing_match_count > 1:
                trip.warnings.append(
                    "模板中有多段与本次行程人员、城市和完整日期范围一致的旧记录；为避免覆盖错误，已追加新行程，请核对重复记录。"
                )
            separator_row = start_row + block_rows
            if separator_row >= summary_row:
                insert_count = separator_row - summary_row + 1
                insertion_log.append({"row": summary_row, "count": insert_count})
                worksheet.insert_rows(summary_row, amount=insert_count)
                summary_row += insert_count
                style_source = max(2, cursor)
            trip_style_source = style_source
        final_row, expected_total = _write_trip(
            worksheet, trip, start_row, inferred_po, trip_style_source,
        )
        is_replacement = existing_start is not None
        write_range_end = final_row if is_replacement else final_row + 1
        writable_row_ranges.append((start_row, write_range_end))
        delivery_assessment = assess_trip_delivery(trip)
        if not is_replacement:
            _copy_row_style(worksheet, trip_style_source, final_row + 1)
            _clear_data_cells(worksheet, final_row + 1)
        cursor = max(cursor, final_row)
        reserved_trip_ranges.append((start_row, write_range_end))
        written.append({
            "file": trip.source_file,
            "name": trip.name,
            "city": trip.city,
            "itinerary_cities": trip.itinerary_cities,
            "city_by_day": {
                day.isoformat(): trip.city_for(day) for day in trip.days
            },
            "start": trip.start.isoformat(),
            "end": trip.end.isoformat(),
            "rows": [start_row, final_row],
            "write_mode": write_mode,
            "expense_count": len(trip.expenses),
            "totals": trip.totals,
            "expenses": [expense.as_dict() for expense in trip.expenses],
            "correction_records": copy.deepcopy(trip.correction_records),
            "expected_sum_without_tax": expected_total,
            "warnings": trip.warnings,
            "page_coverage": copy.deepcopy(trip.page_coverage),
            "structured_invoices": [invoice.as_dict() for invoice in trip.structured_invoices],
            **delivery_assessment,
            "policy_summary": trip.policy_summary,
            "policy_trace": trip.policy_trace,
            "policy_exceptions": trip.policy_exceptions,
        })

    _refresh_summary_formulas(worksheet)

    claim_reconciliation = None
    claim_error = None
    if claim_workbook_path:
        try:
            claim_index = load_claim_index(claim_workbook_path)
            reconciliations = reconcile_trips(trips, claim_index)
            claim_reconciliation = add_reconciliation_sheets(
                workbook,
                reconciliations,
                claim_source=claim_index.source_path,
            )
            claim_reconciliation["claim_index"] = claim_index.as_dict()
            claim_reconciliation["trips"] = [item.as_dict() for item in reconciliations]
        except (ClaimWorkbookError, OSError, ValueError) as exc:
            claim_error = str(exc)

    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
    except AttributeError:  # openpyxl version compatibility
        pass
    workbook.save(destination)
    workbook.close()
    output_validation = template_adapter.validate_output(destination)
    if not output_validation.ok:
        raise ValueError(
            "差旅输出校验失败: " + "；".join(output_validation.errors)
        )
    preservation_audit = template_adapter.audit_preserved_content(
        template,
        destination,
        insertions=insertion_log,
        writable_row_ranges=writable_row_ranges,
        summary_row=original_summary_row,
    )
    output_validation.metadata["template_preservation"] = preservation_audit
    if preservation_audit["status"] != "PASS":
        output_validation.status = "BLOCKED"
        output_validation.errors.append("模板保留区域对账失败。")
        raise ValueError("差旅输出模板保留校验失败: 模板非写入区域发生变化。")

    formula_cache_results = []
    for trip, trip_result in zip(trips, written):
        cache_audit = template_adapter.audit_formula_calculation(
            destination,
            trip.allocation_audit.get("formula_cache_expectations", []),
        )
        trip.allocation_audit["formula_recalculation"] = cache_audit["status"]
        trip.allocation_audit["formula_cache_audit"] = cache_audit
        assessment = assess_trip_delivery(trip)
        for key in (
            "delivery_status", "ready_for_use", "review_required", "delivery_issues",
            "invoice_validation", "review_page_count", "policy_status",
            "allocation_audit", "allocation_trace",
        ):
            trip_result[key] = assessment[key]
        formula_cache_results.append({
            "file": trip.source_file,
            "status": cache_audit["status"],
            "formula_error_count": cache_audit.get("formula_error_count", 0),
            "mismatch_count": cache_audit.get("mismatch_count", 0),
        })
    if any(item["status"] == "BLOCKED" for item in formula_cache_results):
        formula_cache_status = "BLOCKED"
        output_validation.warnings.append("公式重算结果存在错误或与独立期望金额不符；相关行程保留为待复核草稿。")
    elif any(item["status"] != "PASS" for item in formula_cache_results):
        formula_cache_status = "PENDING_EXCEL_RECALCULATION"
        output_validation.warnings.append("公式已写入但尚无可核验的 Excel 重算缓存；相关行程保留为待复核草稿。")
    else:
        formula_cache_status = "PASS"
    output_validation.metadata["formula_cache_audit"] = {
        "status": formula_cache_status,
        "trips": formula_cache_results,
    }
    if claim_reconciliation is not None:
        output_validation.metadata["claim_reconciliation"] = claim_reconciliation
        if claim_reconciliation["status"] != "PASS":
            output_validation.warnings.append(
                "报销核对存在异常或未匹配项，详见“报销核对”工作表。"
            )
    elif claim_error:
        output_validation.warnings.append(f"报销明细无法读取，已跳过核对: {claim_error}")
    return {
        "output_path": destination,
        "output_name": os.path.basename(destination),
        "sheet": worksheet.title,
        "policy_name": policy_name,
        "template_validation": template_validation.as_dict(),
        "output_validation": output_validation.as_dict(),
        "trips": written,
        "claim_reconciliation": claim_reconciliation,
        "claim_error": claim_error,
    }


def preview_review_workbook(
    template_path: str,
    trips: list[Trip],
    output_date: Optional[str] = None,
    po_no=None,
    policy_name: str = "VGC",
    claim_workbook_path: Optional[str] = None,
) -> dict:
    """Render a temporary versioned copy and return a cell-level diff.

    The selected template is opened read-only for validation and is never
    overwritten.  The temporary workbook is removed when the preview ends.
    """
    template = os.path.abspath(template_path)
    adapter = TravelExpenseTemplateAdapter()
    template_validation = adapter.validate_template(template)
    if not template_validation.ok:
        return {
            "ok": False,
            "preview": True,
            "template_validation": template_validation.as_dict(),
            "error": "差旅模板校验失败: " + "；".join(template_validation.errors),
        }
    preview_root = Path(__file__).resolve().parent / "output"
    preview_root.mkdir(parents=True, exist_ok=True)
    preview_path = preview_root / f".travel-preview-{uuid.uuid4().hex}-{Path(template).name}"
    try:
        generated = create_review_workbook(
            template_path=template,
            trips=trips,
            output_date=output_date,
            output_path=str(preview_path),
            po_no=po_no,
            policy_name=policy_name,
            claim_workbook_path=claim_workbook_path,
        )
        diff = adapter.compare_workbooks(template, str(preview_path))
        return {
            "ok": True,
            "preview": True,
            "policy_name": policy_name,
            "template_validation": template_validation.as_dict(),
            "output_validation": generated["output_validation"],
            "trips": generated["trips"],
            "claim_reconciliation": generated.get("claim_reconciliation"),
            "claim_error": generated.get("claim_error"),
            **diff,
        }
    finally:
        try:
            preview_path.unlink(missing_ok=True)
        except OSError:
            pass
