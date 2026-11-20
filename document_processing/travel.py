"""Provider-neutral structured extraction for travel invoices.

The existing travel reviewer remains the authority for the payable expense
calculation and workbook layout.  This module adds a serialisable schema around
those results so every important field can point back to a source page and raw
text.  The extractor is deliberately deterministic: OCR provides text, while
this module parses, validates and flags values for review.
"""

from __future__ import annotations

import hashlib
import re
import unicodedata
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Mapping

from .models import Evidence


MAX_TRAVEL_AMOUNT = 1_000_000_000.0
CONFIRMED_TRAVEL_DATE_SOURCES = {
    "SOURCE_TEXT", "PAGE_DATE_HEADING", "HOTEL_FOLIO", "DIDI_ITINERARY",
    "MANUAL_SOURCE_REVIEW",
}
_DATE_RE = re.compile(r"(?<!\d)(?:(20\d{2})\s*[./-]\s*)?(\d{1,2})\s*[./-]\s*(\d{1,2})(?!\d)")
_DATE_LAST_YEAR_RE = re.compile(r"(?<!\d)(\d{1,2})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(20\d{2}|\d{2})(?!\d)")
_CHINESE_DATE_RE = re.compile(r"(?<!\d)(?:(20\d{2})\s*年\s*)?(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]?")
_ISO_DATE_RE = re.compile(r"(?<!\d)(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})(?!\d)")
_DATE_SPAN_RE = re.compile(
    r"(?P<first>(?:20\d{2}[-/.])?\d{1,2}[-/.]\d{1,2})\s*[-~至到]\s*"
    r"(?P<second>(?:20\d{2}[-/.])?\d{1,2}[-/.]\d{1,2})"
)
_NUMBER_RE = re.compile(
    r"(?<![A-Za-z0-9])\d+(?:(?:[.,]\s*|\s+)\d+)*(?![A-Za-z0-9])"
)
_CURRENCY_RE = re.compile(r"(?:¥|￥|元|RMB|CNY|USD|EUR|€|\$|GBP|£|JPY|日元)", re.IGNORECASE)


@dataclass(frozen=True)
class TravelLineItem:
    """A payable or supporting line associated with one structured invoice."""

    description: str
    amount: float
    category: str = ""
    date_from: date | None = None
    date_to: date | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "description": self.description,
            "amount": round(float(self.amount), 2),
            "category": self.category,
            "date_from": self.date_from.isoformat() if self.date_from else None,
            "date_to": self.date_to.isoformat() if self.date_to else None,
        }


@dataclass
class TravelInvoice:
    """Unified invoice/travel schema used by review and future mappers."""

    document_type: str = "travel_expense"
    vendor: str = ""
    invoice_number: str = ""
    invoice_date: date | None = None
    currency: str = "CNY"
    net_amount: float | None = None
    tax_amount: float | None = None
    total_amount: float | None = None
    travel_date_from: date | None = None
    travel_date_to: date | None = None
    travel_date_source: str = "UNCONFIRMED"
    origin: str = ""
    destination: str = ""
    traveler: str = ""
    expense_type: str = ""
    line_items: list[TravelLineItem] = field(default_factory=list)
    source: dict[str, Any] = field(default_factory=dict)
    evidence: dict[str, dict[str, Any]] = field(default_factory=dict)
    validation: dict[str, Any] = field(default_factory=dict)
    document_id: str = ""
    correction_records: list[dict[str, Any]] = field(default_factory=list)

    def as_dict(self) -> dict[str, Any]:
        return {
            "document_id": self.document_id,
            "document_type": self.document_type,
            "vendor": self.vendor,
            "invoice_number": self.invoice_number,
            "invoice_date": self.invoice_date.isoformat() if self.invoice_date else None,
            "currency": self.currency,
            "net_amount": self.net_amount,
            "tax_amount": self.tax_amount,
            "total_amount": self.total_amount,
            "travel_date_from": self.travel_date_from.isoformat() if self.travel_date_from else None,
            "travel_date_to": self.travel_date_to.isoformat() if self.travel_date_to else None,
            "travel_date_source": self.travel_date_source,
            "origin": self.origin,
            "destination": self.destination,
            "traveler": self.traveler,
            "expense_type": self.expense_type,
            "line_items": [item.as_dict() for item in self.line_items],
            "source": _json_value(self.source),
            "evidence": _json_value(self.evidence),
            "validation": _json_value(self.validation),
            "correction_records": _json_value(self.correction_records),
        }


def _json_value(value: Any) -> Any:
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, Mapping):
        return {str(key): _json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_value(item) for item in value]
    return value


def parse_locale_number(value: Any, locale: str = "auto") -> float | None:
    """Parse common Chinese, US and German amount formats.

    ``auto`` treats the right-most separator as decimal when both are present.
    For a single comma, three trailing digits mean a thousands separator; an
    explicit ``de``/``en``/``zh`` locale removes the remaining ambiguity.
    """

    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return round(number, 2) if abs(number) <= MAX_TRAVEL_AMOUNT else None

    text = unicodedata.normalize("NFKC", str(value)).strip()
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    text = re.sub(r"(?:¥|￥|RMB|CNY|USD|EUR|GBP|JPY|元|€|\$|£|日元|Y)\s*", "", text, flags=re.IGNORECASE)
    text = text.replace("\u00a0", "").replace(" ", "")
    text = re.sub(r"[^0-9,.+-]", "", text)
    if not text or text.count("+") + text.count("-") > 1:
        return None

    requested = (locale or "auto").casefold()
    if requested in {"de", "de-de", "german", "eu", "europe"}:
        decimal, thousands = ",", "."
    elif requested in {"en", "en-us", "zh", "zh-cn", "us", "cn", "english", "chinese"}:
        decimal, thousands = ".", ","
    elif "," in text and "." in text:
        decimal = "," if text.rfind(",") > text.rfind(".") else "."
        thousands = "." if decimal == "," else ","
    elif "," in text:
        tail = text.rsplit(",", 1)[-1]
        decimal = "," if len(tail) != 3 else ""
        thousands = "" if decimal else ","
    else:
        # Dot-only OCR output is the dominant format in the existing Chinese
        # corpus, so 1.234 is kept as 1.234 unless the caller specifies de.
        decimal, thousands = ".", ""

    if thousands:
        text = text.replace(thousands, "")
    if decimal:
        if text.count(decimal) > 1:
            return None
        text = text.replace(decimal, ".")
    if text.count(".") > 1 or not re.fullmatch(r"[+-]?\d+(?:\.\d+)?", text):
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    if negative:
        number = -number
    if abs(number) > MAX_TRAVEL_AMOUNT:
        return None
    return round(number, 2)


def parse_travel_date(value: Any, default_year: int | None = None, locale: str = "auto") -> date | None:
    """Parse ISO, Chinese and short invoice dates without guessing wildly."""

    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = unicodedata.normalize("NFKC", str(value or "")).strip()
    if not text:
        return None
    year_default = int(default_year or datetime.now().year)
    iso = _ISO_DATE_RE.search(text)
    if iso:
        try:
            return date(int(iso.group(1)), int(iso.group(2)), int(iso.group(3)))
        except ValueError:
            return None
    chinese = _CHINESE_DATE_RE.search(text)
    if chinese:
        year = int(chinese.group(1) or year_default)
        try:
            return date(year, int(chinese.group(2)), int(chinese.group(3)))
        except ValueError:
            return None
    last_year = _DATE_LAST_YEAR_RE.search(text)
    if last_year:
        first, second, year_text = last_year.groups()
        year = int(year_text)
        if len(year_text) == 2:
            year += 2000
        is_european = (locale or "auto").casefold() in {"de", "de-de", "german", "eu"}
        first, second = int(first), int(second)
        day, month = (first, second) if is_european or first > 12 else (second, first)
        try:
            return date(year, month, day)
        except ValueError:
            return None
    match = _DATE_RE.search(text)
    if not match:
        return None
    year_text, first_text, second_text = match.groups()
    first, second = int(first_text), int(second_text)
    if year_text:
        year = int(year_text)
        month, day = first, second
    else:
        year = year_default
        is_european = (locale or "auto").casefold() in {"de", "de-de", "german", "eu"}
        if is_european or first > 12:
            day, month = first, second
        else:
            month, day = first, second
    try:
        return date(year, month, day)
    except ValueError:
        return None


def _normalise_text(value: Any) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", str(value or ""))).strip()


def _normalise_blocks(blocks: Iterable[Any]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for block in blocks or ():
        if hasattr(block, "as_dict"):
            block = block.as_dict()
        if isinstance(block, Mapping):
            result.append(dict(block))
    return result


def _amount_tokens(line: str) -> list[tuple[float, str]]:
    result: list[tuple[float, str]] = []
    for match in _NUMBER_RE.finditer(line):
        raw = match.group(0)
        parsed = parse_locale_number(raw)
        if parsed is not None:
            result.append((parsed, raw))
    return result


def _label_value(lines: list[str], patterns: Iterable[str]) -> tuple[str, str, int] | None:
    compiled = [re.compile(pattern, re.IGNORECASE) for pattern in patterns]
    for index, line in enumerate(lines):
        for pattern in compiled:
            match = pattern.search(line)
            if match:
                tail = line[match.end():].strip(" :：-—")
                if tail:
                    return tail, line, index
                if index + 1 < len(lines) and lines[index + 1].strip():
                    return lines[index + 1].strip(), line + " " + lines[index + 1].strip(), index
    return None


def _label_amount(lines: list[str], patterns: Iterable[str]) -> tuple[float, str, int] | None:
    compiled = [re.compile(pattern, re.IGNORECASE) for pattern in patterns]
    for index, line in enumerate(lines):
        if not any(pattern.search(line) for pattern in compiled):
            continue
        amounts = _amount_tokens(line)
        if amounts:
            amount, _raw = amounts[-1]
            return amount, line, index
        if index + 1 < len(lines):
            amounts = _amount_tokens(lines[index + 1])
            if amounts:
                amount, _raw = amounts[-1]
                return amount, line + " " + lines[index + 1].strip(), index
    return None


_TOTAL_LABEL_RE = re.compile(
    r"(?:价税合计(?:（小写）|\(小写\))?|含税合计|合计金额|应付金额|"
    r"gross\s*amount|total\s*(?:amount|due)?|grand\s*total|合计)",
    re.IGNORECASE,
)


def _invoice_page_fragment(text: str, total_amount: float | None, ordinal: int) -> str:
    """Return the page segment belonging to one of several same-page invoices."""

    if not text or total_amount is None:
        return text or ""
    candidates: list[tuple[int, int, float]] = []
    for label in _TOTAL_LABEL_RE.finditer(text):
        # OCR sometimes places the small-value marker before ``价税合计`` and
        # sometimes after it. Prefer a currency-marked amount in a bounded
        # window; this also avoids reading the trailing ``00`` from ``x10. 00``
        # as a separate amount when the preceding ``10`` is attached to noise.
        window_start = max(0, label.start() - 180)
        window_end = min(len(text), label.end() + 180)
        window = text[window_start:window_end]
        currency_matches = list(re.finditer(
            r"(?:¥|￥|RMB|CNY)\s*([0-9][0-9\s,.]*)",
            window,
            flags=re.IGNORECASE,
        ))
        amount_match = None
        parsed = None
        if currency_matches:
            currency_match = min(
                currency_matches,
                key=lambda item: abs((window_start + item.start()) - label.start()),
            )
            parsed = parse_locale_number(currency_match.group(1))
            amount_match = currency_match
        if parsed is None:
            tail = text[label.end():label.end() + 140]
            amount_match = _NUMBER_RE.search(tail)
            parsed = parse_locale_number(amount_match.group(0)) if amount_match else None
        if parsed is None:
            continue
        amount_end = window_start + amount_match.end() if currency_matches and amount_match else label.end() + amount_match.end()
        candidates.append((label.start(), amount_end, parsed))
    matching = [index for index, item in enumerate(candidates) if abs(item[2] - total_amount) <= 0.01]
    if not matching:
        return text
    target_index = matching[min(max(ordinal, 0), len(matching) - 1)]
    start = candidates[target_index - 1][1] if target_index else 0
    end = candidates[target_index + 1][0] if target_index + 1 < len(candidates) else len(text)
    return text[start:end]


def _currency(text: str) -> str:
    match = _CURRENCY_RE.search(text or "")
    if not match:
        return "CNY"
    token = match.group(0).upper()
    if token in {"¥", "￥", "元", "RMB", "CNY", "Y"}:
        return "CNY"
    if token in {"€", "EUR"}:
        return "EUR"
    if token in {"$", "USD"}:
        return "USD"
    if token in {"£", "GBP"}:
        return "GBP"
    if token in {"日元", "JPY"}:
        return "JPY"
    return token


def _clean_label(value: str) -> str:
    value = _normalise_text(value).strip(" :：,，;")
    return re.split(r"\s+(?:发票号码|发票代码|金额|税额|Total|VAT)\b", value, maxsplit=1, flags=re.IGNORECASE)[0]


def _route(text: str) -> tuple[str, str]:
    for line in text.splitlines():
        match = re.search(r"([^\s,，:：]{2,30})\s*[-—－至到]\s*([^\s,，:：]{2,30})", line)
        if not match:
            continue
        first, second = match.groups()
        # A date range is not a travel origin/destination route.
        if parse_travel_date(first) and parse_travel_date(second):
            continue
        if any(char.isalpha() or "\u4e00" <= char <= "\u9fff" for char in match.group(0)):
            return _clean_label(first), _clean_label(second)

    origin_labels = re.compile(
        r"出发地|出发站|出发城市|起飞地|始发地|始发站|起点|\bFROM\b|\bORIGIN\b",
        re.IGNORECASE,
    )
    destination_labels = re.compile(
        r"到达地|到达站|到达城市|目的地|目的站|终点|降落地|\bTO\b|\bDESTINATION\b",
        re.IGNORECASE,
    )
    all_labels = re.compile(
        rf"(?:{origin_labels.pattern})|(?:{destination_labels.pattern})",
        re.IGNORECASE,
    )

    def labelled_value(labels: re.Pattern) -> str:
        for label in labels.finditer(text):
            next_label = all_labels.search(text, label.end())
            line_end = text.find("\n", label.end())
            stop = min(
                item for item in (next_label.start() if next_label else len(text),
                                  line_end if line_end >= 0 else len(text))
                if item >= label.end()
            )
            value = _clean_label(text[label.end():stop])
            if value:
                return value
        return ""

    origin = labelled_value(origin_labels)
    destination = labelled_value(destination_labels)
    if origin and destination:
        return origin, destination
    return "", ""


def _date_span(text: str, default_year: int | None) -> tuple[date | None, date | None]:
    match = _DATE_SPAN_RE.search(text or "")
    if match:
        start = parse_travel_date(match.group("first"), default_year)
        end = parse_travel_date(match.group("second"), default_year)
        return start, end
    dates = []
    for match in _ISO_DATE_RE.finditer(text or ""):
        parsed = parse_travel_date(match.group(0), default_year)
        if parsed and parsed not in dates:
            dates.append(parsed)
    return (dates[0], dates[-1]) if dates else (None, None)


def _bbox_for_value(value: Any, raw_text: str, blocks: list[dict[str, Any]]) -> list[float] | None:
    target = _normalise_text(value).casefold()
    numeric_target = parse_locale_number(value)
    selected: list[dict[str, Any]] = []
    for block in blocks:
        token = _normalise_text(block.get("text", ""))
        if not token:
            continue
        if target and (target in token.casefold() or token.casefold() in target):
            selected.append(block)
            continue
        numeric_token = parse_locale_number(token)
        if numeric_target is not None and numeric_token is not None and abs(numeric_target - numeric_token) <= 0.01:
            selected.append(block)
    if not selected:
        return None
    bboxes = []
    for block in selected:
        bbox = block.get("bbox")
        if bbox and len(bbox) == 4:
            bboxes.append([float(item) for item in bbox])
            continue
        if all(block.get(key) is not None for key in ("left", "top", "width", "height")):
            bboxes.append([
                float(block["left"]), float(block["top"]),
                float(block["left"]) + float(block["width"]),
                float(block["top"]) + float(block["height"]),
            ])
    if not bboxes:
        return None
    return [
        min(item[0] for item in bboxes), min(item[1] for item in bboxes),
        max(item[2] for item in bboxes), max(item[3] for item in bboxes),
    ]


def _evidence(
    field_name: str,
    value: Any,
    source_file: str,
    page: int,
    raw_text: str,
    blocks: list[dict[str, Any]],
    ocr_engine: str,
    ocr_confidence: float | None,
    validation_status: str,
) -> dict[str, Any]:
    return Evidence(
        field=field_name,
        value=value,
        source_file=source_file,
        page=page,
        bbox=_bbox_for_value(value, raw_text, blocks),
        raw_text=raw_text[:500],
        ocr_engine=ocr_engine,
        ocr_confidence=ocr_confidence,
        validation_status=validation_status,
    ).as_dict()


def _source_sha256(source_file: str) -> str:
    try:
        path = Path(source_file)
        if not path.is_file():
            return ""
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()
    except OSError:
        return ""


def _document_type(category: str, text: str) -> str:
    lowered = (text or "").casefold()
    category = (category or "").casefold()
    if category in {"special", "toll", "toll_receipt", "special_expense"}:
        category = "taxi"
    # The existing Travel parser has already separated same-page invoices by
    # category.  Honour that business classification before looking at noisy
    # page-wide OCR keywords (for example a toll invoice beside a taxi ticket).
    if category == "accommodation":
        return "hotel_invoice"
    if category == "ticket":
        return "flight_ticket" if any(item in lowered for item in ("flight", "航空", "航班", "机票", "电子客票")) else "train_ticket"
    if category == "taxi":
        return "taxi_receipt"
    if any(item in lowered for item in ("hotel", "酒店", "住宿", "客房")):
        return "hotel_invoice"
    if any(item in lowered for item in ("taxi", "滴滴", "出租", "网约车")):
        return "taxi_receipt"
    return "travel_expense"


def _extract_vendor(lines: list[str], text: str) -> tuple[str, str, int]:
    found = _label_value(lines, (
        r"(?:销售方名称|销售方|商户名称|酒店名称|收款方|供应商|服务商)\s*",
        r"(?:vendor|supplier|merchant|hotel|airline)\s*(?:name)?\s*[:：]?\s*",
    ))
    if found:
        value, raw, index = found
        return _clean_label(value), raw, index
    lowered = text.casefold()
    if "didi" in lowered or "滴滴" in text:
        index = next((i for i, line in enumerate(lines) if "didi" in line.casefold() or "滴滴" in line), 0)
        return "Didi", lines[index], index
    return "", "", 0


def _extract_invoice_number(lines: list[str]) -> tuple[str, str, int]:
    found = _label_value(lines, (
        r"(?:发票号码|发票号|电子发票号码)\s*",
        r"(?:invoice\s*(?:no|number)|invoice#|folio\s*no)\s*[:：#]?\s*",
    ))
    if not found:
        return "", "", 0
    value, raw, index = found
    match = re.search(r"[A-Za-z0-9-]{4,}", value)
    return (match.group(0) if match else _clean_label(value)), raw, index


def _extract_invoice_date(lines: list[str], default_year: int | None) -> tuple[date | None, str, int]:
    found = _label_value(lines, (
        r"(?:开票日期|开票时间|出票日期|发票日期)\s*",
        r"(?:invoice\s*date|issue\s*date|date)\s*[:：]?\s*",
    ))
    if not found:
        return None, "", 0
    value, raw, index = found
    return parse_travel_date(value, default_year), raw, index


def extract_travel_invoice(
    text: str,
    *,
    source_file: str = "",
    page: int = 1,
    category: str = "",
    total_amount: float | None = None,
    traveler: str = "",
    trip_start: date | None = None,
    trip_end: date | None = None,
    expense_start: date | None = None,
    expense_end: date | None = None,
    expense_date_source: str = "",
    expense_date_evidence_page: int | None = None,
    expense_date_evidence_text: str = "",
    default_year: int | None = None,
    blocks: Iterable[Any] = (),
    ocr_engine: str = "",
    ocr_confidence: float | None = None,
    project_id: str = "",
    run_id: str = "",
    line_description: str = "",
) -> TravelInvoice:
    """Extract one deterministic structured invoice from one source page."""

    category = (category or "").casefold()
    if category in {"special", "toll", "toll_receipt", "special_expense"}:
        category = "taxi"
    normalised = _normalise_text(text)
    lines = [_normalise_text(line) for line in (text or "").splitlines() if _normalise_text(line)]
    block_list = _normalise_blocks(blocks)
    year = default_year or next((int(match.group(1)) for match in _ISO_DATE_RE.finditer(text or "")), None)

    vendor, vendor_raw, vendor_index = _extract_vendor(lines, text)
    invoice_number, invoice_raw, invoice_index = _extract_invoice_number(lines)
    invoice_date, invoice_date_raw, invoice_date_index = _extract_invoice_date(lines, year)
    route_origin, route_destination = _route(text)
    date_from, date_to = _date_span(text, year)
    labeled_travel = _label_value(lines, (r"(?:行程日期|乘车日期|旅行日期|travel\s*date)\s*",))
    if labeled_travel:
        travel_from, travel_to = _date_span(labeled_travel[0], year)
        date_from, date_to = travel_from or date_from, travel_to or date_to
    elif invoice_date and date_from == invoice_date and date_to == invoice_date:
        # A lone ``Invoice Date`` is not evidence of the service/travel date.
        # Prefer the expense/trip fallback below instead of copying the issue
        # date into the travel interval.
        date_from = date_to = None

    net_item = _label_amount(lines, (
        r"(?:不含税金额|不含税|税前金额|金额|小计)\s*",
        r"(?:net\s*amount|subtotal|amount\s*excl(?:usive)?\s*tax)\s*[:：]?\s*",
    ))
    tax_item = _label_amount(lines, (
        r"(?:税额|增值税|税款)\s*",
        r"(?:tax|vat)\s*(?:amount)?\s*[:：]?\s*",
    ))
    total_item = _label_amount(lines, (
        r"(?:价税合计(?:（小写）|\(小写\))?|含税合计|合计金额|应付金额)\s*",
        r"(?:gross\s*amount|total\s*(?:amount|due)?|grand\s*total)\s*[:：]?\s*",
    ))
    if total_amount is None and total_item:
        total_amount = total_item[0]
    # Bare OCR table headers such as ``金额[元]`` are not proof of an invoice
    # net amount. Require a currency marker for the generic Chinese ``金额``
    # label; explicit net/subtotal labels remain usable without one.
    if net_item and "金额" in net_item[1] and not re.search(
        r"(?:不含税|税前|小计|net|subtotal)", net_item[1], re.IGNORECASE,
    ) and not _CURRENCY_RE.search(net_item[1]):
        net_item = None
    currency = _currency(normalised)
    if net_item and tax_item and total_amount is not None:
        # Keep the locale parser as the source of numeric values, but leave
        # arithmetic reconciliation to the validation function below.
        pass

    effective_from = expense_start or date_from or trip_start
    effective_to = expense_end or date_to or trip_end or effective_from
    if effective_from and not effective_to:
        effective_to = effective_from
    if effective_to and not effective_from:
        effective_from = effective_to
    if expense_start or expense_end:
        travel_date_source = expense_date_source or "EXPENSE_RECORD"
    elif labeled_travel and (date_from or date_to):
        travel_date_source = "SOURCE_TEXT"
    elif date_from or date_to:
        travel_date_source = "PAGE_DATE_CANDIDATE"
    elif trip_start or trip_end:
        travel_date_source = "TRIP_FALLBACK"
    else:
        travel_date_source = "UNCONFIRMED"

    source = {
        "source_file": source_file,
        "sha256": _source_sha256(source_file),
        "page": page,
        "document_type": _document_type(category, text),
        "service_date_evidence_page": (
            (expense_date_evidence_page or page)
            if travel_date_source in CONFIRMED_TRAVEL_DATE_SOURCES else None
        ),
        "extraction_method": "ocr" if ocr_engine and ocr_engine != "native_text" else "native_text",
        "run_id": run_id,
        "project_id": project_id,
    }
    # ``total_amount`` is the invoice-level payable amount, not a line item.
    # Keeping this list empty until a real item/quantity parser is available
    # prevents validation from trivially proving total == total.
    line_items: list[TravelLineItem] = []
    invoice = TravelInvoice(
        document_type=source["document_type"],
        vendor=vendor,
        invoice_number=invoice_number,
        invoice_date=invoice_date,
        currency=currency,
        net_amount=net_item[0] if net_item else None,
        tax_amount=tax_item[0] if tax_item else None,
        total_amount=round(float(total_amount), 2) if total_amount is not None else None,
        travel_date_from=effective_from,
        travel_date_to=effective_to,
        travel_date_source=travel_date_source,
        origin=route_origin,
        destination=route_destination,
        traveler=traveler,
        expense_type=category,
        line_items=line_items,
        source=source,
    )
    invoice.document_id = hashlib.sha256(
        f"{source_file}|{page}|{category}|{invoice.invoice_number}|{invoice.total_amount}|{line_description}".encode("utf-8")
    ).hexdigest()[:20]

    validation = validate_travel_invoice(invoice)
    invoice.validation = validation
    travel_date_evidence = (
        expense_date_evidence_text
        or (labeled_travel[1] if labeled_travel else "")
        or line_description
        or normalised
    )
    date_is_confirmed = travel_date_source in CONFIRMED_TRAVEL_DATE_SOURCES
    evidence_sources = {
        "vendor": (vendor, vendor_raw),
        "invoice_number": (invoice_number, invoice_raw),
        "invoice_date": (invoice_date, invoice_date_raw),
        "net_amount": (invoice.net_amount, net_item[1] if net_item else ""),
        "tax_amount": (invoice.tax_amount, tax_item[1] if tax_item else ""),
        "total_amount": (invoice.total_amount, total_item[1] if total_item else line_description),
        "travel_date_from": (
            invoice.travel_date_from if date_is_confirmed else None,
            travel_date_evidence if date_is_confirmed else "",
        ),
        "travel_date_to": (
            invoice.travel_date_to if date_is_confirmed else None,
            travel_date_evidence if date_is_confirmed else "",
        ),
        "origin": (invoice.origin, next((line for line in lines if invoice.origin and invoice.origin in line), "")),
        "destination": (invoice.destination, next((line for line in lines if invoice.destination and invoice.destination in line), "")),
        "traveler": (invoice.traveler, next((line for line in lines if invoice.traveler and invoice.traveler in line), "")),
        "currency": (invoice.currency, normalised),
        "expense_type": (invoice.expense_type, normalised),
    }
    for field_name, (value, raw) in evidence_sources.items():
        if value in (None, ""):
            continue
        evidence_page = page
        if field_name in {"travel_date_from", "travel_date_to"}:
            evidence_page = int(invoice.source.get("service_date_evidence_page") or page)
        invoice.evidence[field_name] = _evidence(
            field_name, value, source_file, evidence_page, raw or normalised,
            block_list, ocr_engine, ocr_confidence, validation["status"],
        )
    return invoice


def _expense_value(expense: Any, key: str, default: Any = None) -> Any:
    if isinstance(expense, Mapping):
        return expense.get(key, default)
    return getattr(expense, key, default)


def _normalise_pages(pages: Any) -> list[dict[str, Any]]:
    if isinstance(pages, str):
        pages = pages.split("\f")
    result = []
    for index, page in enumerate(pages or (), 1):
        if isinstance(page, Mapping):
            result.append({
                "text": str(page.get("text", "")),
                "page": int(page.get("page", index)),
                "blocks": page.get("blocks", ()),
                "ocr_engine": page.get("ocr_engine", ""),
                "ocr_confidence": page.get("ocr_confidence"),
            })
        else:
            result.append({"text": str(page or ""), "page": index, "blocks": (), "ocr_engine": "", "ocr_confidence": None})
    return result


def extract_travel_invoices(
    pages: Any,
    *,
    source_file: str = "",
    expenses: Iterable[Any] = (),
    traveler: str = "",
    trip_start: date | None = None,
    trip_end: date | None = None,
    default_year: int | None = None,
    ocr_engine: str = "",
    ocr_confidence: float | None = None,
    project_id: str = "",
    run_id: str = "",
) -> list[TravelInvoice]:
    """Project existing Travel ``Expense`` records into the unified schema.

    One record is produced per payable Expense, preserving the existing
    same-page multi-invoice behaviour.  If no expenses are supplied, clearly
    labelled page totals are still extracted as standalone candidates.
    """

    page_list = _normalise_pages(pages)
    by_page = {item["page"]: item for item in page_list}
    expense_list = list(expenses or ())
    records: list[TravelInvoice] = []
    page_ordinals: dict[int, int] = {}
    if expense_list:
        for expense in expense_list:
            page = int(_expense_value(expense, "source_page", 1) or 1)
            page_data = by_page.get(page, {"text": "", "blocks": (), "ocr_engine": "", "ocr_confidence": None})
            amount = parse_locale_number(_expense_value(expense, "amount"))
            ordinal = page_ordinals.get(page, 0)
            page_ordinals[page] = ordinal + 1
            page_text = _invoice_page_fragment(page_data["text"], amount, ordinal)
            records.append(extract_travel_invoice(
                page_text, source_file=source_file, page=page,
                category=str(_expense_value(expense, "category", "") or ""),
                total_amount=amount,
                traveler=traveler, trip_start=trip_start, trip_end=trip_end,
                expense_start=_expense_value(expense, "start"), expense_end=_expense_value(expense, "end"),
                expense_date_source=str(_expense_value(expense, "service_date_source", "") or ""),
                expense_date_evidence_page=_expense_value(expense, "service_date_evidence_page"),
                expense_date_evidence_text=str(_expense_value(expense, "service_date_evidence_text", "") or ""),
                default_year=default_year, blocks=page_data["blocks"],
                ocr_engine=ocr_engine or page_data["ocr_engine"],
                ocr_confidence=ocr_confidence if ocr_confidence is not None else page_data["ocr_confidence"],
                project_id=project_id, run_id=run_id,
                line_description=str(_expense_value(expense, "description", "") or ""),
            ))
    else:
        for page_data in page_list:
            total_item = _label_amount(
                [_normalise_text(line) for line in page_data["text"].splitlines() if _normalise_text(line)],
                (r"(?:价税合计|含税合计|合计金额|total|grand\s*total)\s*",),
            )
            if not total_item:
                continue
            records.append(extract_travel_invoice(
                page_data["text"], source_file=source_file, page=page_data["page"],
                total_amount=total_item[0], traveler=traveler,
                trip_start=trip_start, trip_end=trip_end, default_year=default_year,
                blocks=page_data["blocks"], ocr_engine=ocr_engine or page_data["ocr_engine"],
                ocr_confidence=ocr_confidence if ocr_confidence is not None else page_data["ocr_confidence"],
                project_id=project_id, run_id=run_id,
            ))

    # A scanned page can contain several invoices.  If the page-level OCR
    # exposes one vendor or invoice number for every projected expense, that
    # value is not safely attributable to each payable item.  Remove only
    # these ambiguous identity fields and keep the amount/category/date
    # records intact for review.
    by_page_records: dict[int, list[TravelInvoice]] = {}
    for invoice in records:
        by_page_records.setdefault(int(invoice.source.get("page", 1)), []).append(invoice)
    for page_records in by_page_records.values():
        if len(page_records) < 2:
            continue
        for field_name in ("vendor", "invoice_number", "net_amount", "tax_amount"):
            values = [getattr(invoice, field_name) for invoice in page_records]
            present = [value for value in values if value not in (None, "")]
            if len(present) < 2 or len({str(value) for value in present}) != 1:
                continue
            for invoice in page_records:
                setattr(invoice, field_name, "" if field_name in {"vendor", "invoice_number"} else None)
                invoice.evidence.pop(field_name, None)
        for invoice in page_records:
            invoice.validation = validate_travel_invoice(invoice)
            for evidence in invoice.evidence.values():
                evidence["validation_status"] = invoice.validation["status"]

    duplicate_rows = detect_duplicate_invoices(records)
    for row in duplicate_rows:
        invoice = records[row["index"]]
        invoice.validation.setdefault("issues", []).append({
            "code": "DUPLICATE_INVOICE",
            "message": f"可能与第 {row['duplicate_of'] + 1} 条发票重复",
            "severity": "WARNING",
            "field": "invoice_number" if invoice.invoice_number else "total_amount",
        })
        invoice.validation["status"] = "WARNING"
        invoice.validation["duplicate_of"] = records[row["duplicate_of"]].document_id
    return records


def validate_travel_invoice(invoice: TravelInvoice) -> dict[str, Any]:
    """Run deterministic arithmetic and basic business validation."""

    issues: list[dict[str, Any]] = []
    source_file = str(invoice.source.get("source_file", "") or "").strip()
    source_page = invoice.source.get("page")
    try:
        source_page_valid = int(source_page) >= 1
    except (TypeError, ValueError):
        source_page_valid = False
    accepted_unknown = invoice.source.get("accepted_unknown_service_date")
    accepted_unknown_service_date = bool(
        isinstance(accepted_unknown, dict)
        and accepted_unknown.get("review_status") == "approved"
        and accepted_unknown.get("source_sha256")
        == str(invoice.source.get("sha256", "")).lower()
        and accepted_unknown.get("source_page") == source_page
        and bool(accepted_unknown.get("reviewed_on"))
        and invoice.travel_date_from is None
        and invoice.travel_date_to is None
    )
    checks: dict[str, Any] = {
        "total_present": invoice.total_amount is not None,
        "arithmetic_checked": False,
        "date_range_valid": True,
        "source_file_present": bool(source_file),
        "source_page_valid": source_page_valid,
        "service_date_present": bool(invoice.travel_date_from and invoice.travel_date_to),
        "service_date_source": invoice.travel_date_source,
        "service_date_unavailable_accepted": accepted_unknown_service_date,
    }
    if not source_file:
        issues.append({
            "code": "MISSING_SOURCE_FILE",
            "message": "缺少原始来源文件，无法回查票面",
            "severity": "WARNING",
            "field": "source",
        })
    if not source_page_valid:
        issues.append({
            "code": "INVALID_SOURCE_PAGE",
            "message": "缺少有效来源页码，无法定位票面",
            "severity": "WARNING",
            "field": "source.page",
        })
    if not checks["service_date_present"]:
        issues.append({
            "code": (
                "SERVICE_DATE_UNKNOWN_ACCEPTED"
                if accepted_unknown_service_date else "MISSING_SERVICE_DATE"
            ),
            "message": (
                "票面无法确认实际服务日期；已按人工审核显式保留未知状态"
                if accepted_unknown_service_date else "没有可确认的实际服务日期"
            ),
            "severity": "INFO" if accepted_unknown_service_date else "WARNING",
            "field": "travel_date_from",
        })
    elif invoice.travel_date_source not in CONFIRMED_TRAVEL_DATE_SOURCES:
        issues.append({
            "code": "UNCONFIRMED_SERVICE_DATE",
            "message": "当前日期来自行程/记录回退值，未确认是票面实际服务日期",
            "severity": "WARNING",
            "field": "travel_date_from",
        })
    if invoice.total_amount is None:
        issues.append({"code": "MISSING_TOTAL", "message": "未识别价税合计/总金额", "severity": "BLOCKING", "field": "total_amount"})
    elif invoice.total_amount < 0 or invoice.total_amount > MAX_TRAVEL_AMOUNT:
        issues.append({"code": "INVALID_TOTAL", "message": "总金额超出合理范围", "severity": "BLOCKING", "field": "total_amount"})

    if invoice.net_amount is not None and invoice.tax_amount is not None and invoice.total_amount is not None:
        checks["arithmetic_checked"] = True
        checks["net_plus_tax"] = round(invoice.net_amount + invoice.tax_amount, 2)
        checks["total_amount"] = round(invoice.total_amount, 2)
        if abs(checks["net_plus_tax"] - checks["total_amount"]) > 0.02:
            issues.append({
                "code": "NET_TAX_TOTAL_MISMATCH",
                "message": f"金额+税额 {checks['net_plus_tax']:.2f} 与总额 {checks['total_amount']:.2f} 不一致",
                "severity": "WARNING",
                "field": "total_amount",
            })
    line_total = round(sum(item.amount for item in invoice.line_items), 2)
    if invoice.line_items and invoice.total_amount is not None:
        checks["line_items_total"] = line_total
        if abs(line_total - invoice.total_amount) > 0.02:
            issues.append({
                "code": "LINE_ITEMS_TOTAL_MISMATCH",
                "message": f"明细合计 {line_total:.2f} 与总额 {invoice.total_amount:.2f} 不一致",
                "severity": "WARNING",
                "field": "line_items",
            })
    if invoice.travel_date_from and invoice.travel_date_to and invoice.travel_date_to < invoice.travel_date_from:
        checks["date_range_valid"] = False
        issues.append({"code": "INVALID_DATE_RANGE", "message": "行程结束日期早于开始日期", "severity": "WARNING", "field": "travel_date_to"})
    if invoice.net_amount is not None and invoice.net_amount < 0:
        issues.append({"code": "INVALID_NET_AMOUNT", "message": "不含税金额不能为负数", "severity": "WARNING", "field": "net_amount"})
    if invoice.tax_amount is not None and invoice.tax_amount < 0:
        issues.append({"code": "INVALID_TAX_AMOUNT", "message": "税额不能为负数", "severity": "WARNING", "field": "tax_amount"})
    status = (
        "BLOCKED" if any(item["severity"] == "BLOCKING" for item in issues)
        else "WARNING" if any(item["severity"] == "WARNING" for item in issues)
        else "PASS"
    )
    return {"status": status, "issues": issues, "checks": checks}


def _duplicate_key(invoice: TravelInvoice) -> tuple[Any, ...]:
    vendor = re.sub(r"\W", "", invoice.vendor.casefold())
    number = re.sub(r"[^A-Za-z0-9]", "", invoice.invoice_number.casefold())
    total = round(invoice.total_amount, 2) if invoice.total_amount is not None else None
    if number:
        return ("invoice_number", number, vendor, total)
    return (
        "fallback", vendor, invoice.document_type, total,
        invoice.invoice_date.isoformat() if invoice.invoice_date else "",
        invoice.travel_date_from.isoformat() if invoice.travel_date_from else "",
        invoice.destination.casefold(),
    )


def detect_duplicate_invoices(invoices: Iterable[TravelInvoice]) -> list[dict[str, Any]]:
    """Return duplicate candidates without deleting or merging any record."""

    invoice_list = list(invoices)
    seen: dict[tuple[Any, ...], int] = {}
    duplicates: list[dict[str, Any]] = []
    for index, invoice in enumerate(invoice_list):
        key = _duplicate_key(invoice)
        if key[0] == "fallback" and not invoice.vendor and invoice.total_amount is not None:
            # Same-page OCR often produces multiple payable records with no
            # vendor/number. Keep those visible; page-level evidence separates
            # them better than a weak amount-only fingerprint.
            continue
        previous = seen.get(key)
        if previous is not None:
            previous_invoice = invoice_list[previous]
            if previous_invoice.source.get("source_file") == invoice.source.get("source_file") and previous_invoice.source.get("page") == invoice.source.get("page"):
                continue
            duplicates.append({
                "index": index,
                "duplicate_of": previous,
                "key": list(key),
                "reason": "invoice_number_and_total" if key[0] == "invoice_number" else "vendor_type_total_date",
            })
        else:
            seen[key] = index
    return duplicates


def summarise_travel_validation(invoices: Iterable[TravelInvoice]) -> dict[str, Any]:
    records = list(invoices)
    statuses = {"PASS": 0, "WARNING": 0, "BLOCKED": 0}
    for invoice in records:
        status = str(invoice.validation.get("status", "WARNING"))
        statuses[status] = statuses.get(status, 0) + 1
    return {
        "invoice_count": len(records),
        "status_counts": statuses,
        "review_required": statuses.get("WARNING", 0) + statuses.get("BLOCKED", 0),
        "duplicate_count": sum(1 for invoice in records if any(item.get("code") == "DUPLICATE_INVOICE" for item in invoice.validation.get("issues", []))),
    }


__all__ = [
    "TravelInvoice",
    "TravelLineItem",
    "detect_duplicate_invoices",
    "extract_travel_invoice",
    "extract_travel_invoices",
    "parse_locale_number",
    "parse_travel_date",
    "summarise_travel_validation",
    "validate_travel_invoice",
]
