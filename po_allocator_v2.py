# -*- coding: utf-8 -*-
"""
PO 工时规划器 V2 — 基于实际数据的前瞻性规划

核心改进 (vs 旧 V2):
1. 数据源优先级修正: 工时清单表 > 外援情况 > Budget Report > Timesheet > MC(仅核查)
2. 输出从"分配推测"变为"实际用量统计 + 正向推演"
3. 支持自动填补缺失的 PO start/end month
4. 5-Sheet 报告: 状态 / 矩阵 / 月度汇总 / 正向推演 / 提醒

用法:
    python po_allocator_v2.py                    # 默认 CNTS-25165
    python po_allocator_v2.py --project CNTS-25165
    python po_allocator_v2.py --as-of 2026-07    # 指定推演起始月
"""

import os
import sys
import re
import math
from datetime import datetime, timedelta
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import config
from utils import (
    setup_utf8_stdout, safe_load_workbook, read_sheet, read_sheet_raw,
    find_latest_file, safe_str, normalize_name_last_first, parse_date_month,
    infer_level_from_rate,
)
from timesheet_parser import scan_timesheet_directory, records_to_summary


# ============================================================
# 常量
# ============================================================
HOURS_PER_DAY = config.HOURS_PER_DAY
WORKDAYS = getattr(config, 'WORKDAYS_2026', None) or {
    1: 19, 2: 17, 3: 22, 4: 22, 5: 19, 6: 22,
    7: 23, 8: 21, 9: 22, 10: 17, 11: 21, 12: 23,
}
LOOKBACK = getattr(config, 'CONSUMPTION_LOOKBACK_MONTHS', 3)
MC_THRESHOLD = getattr(config, 'MC_VERIFY_THRESHOLD', 0.20)
RATE_TO_LEVEL = getattr(config, 'RATE_TO_LEVEL', {
    2460: "Testing Engineer-Expert",
    2260: "Testing Engineer-Senior",
    2100: "Testing Engineer-Experienced",
    1600: "Testing Engineer-Junior",
})

MONTH_NAMES_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
    "january": 1, "february": 2, "march": 3, "april": 4,
    "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
}


# ============================================================
# 数据模型
# ============================================================
@dataclass
class POAllocationRecord:
    """单条分配记录: 某人某月在某个 PO 上的天数"""
    person_name: str        # 标准化名称 "Tang Qiwei"
    month: str              # "2026-01"
    po_id: str              # "2601000337"
    days: float             # 天数
    rate: float = 0.0       # 单价
    level: str = ""         # Level
    source: str = ""        # 数据来源
    po_days_quota: float = 0.0      # 该PO给此人的配额
    po_days_remaining: float = 0.0  # 该PO给此人的剩余


@dataclass
class POLevelConfig:
    """单张 PO 内某个 Level 的配额"""
    project: str = "CNTS-25165"
    po_id: str = ""
    level: str = ""
    days_quota: float = 0.0
    unit_price: float = 0.0
    value_rmb: float = 0.0
    start_month: str = ""
    end_month: str = ""
    priority: int = 1
    allocation_mode: str = "usage"

    def __post_init__(self):
        self.days_remaining = self.days_quota


@dataclass
class LevelConsumption:
    """某个 Level 的月度消耗统计"""
    level: str
    monthly_totals: dict = field(default_factory=dict)  # {month: total_days}
    monthly_avg: float = 0.0
    active_people: list = field(default_factory=list)
    latest_monthly_rate: float = 0.0  # 最近 N 月平均


@dataclass
class ForwardProjection:
    """单月正向推演结果 (某个 Level)"""
    month: str
    active_po_id: str
    monthly_demand: float
    consumed_from_po: float
    po_remaining_after: float
    transition: bool = False
    transition_date: str = ""
    transition_from_po: str = ""
    transition_to_po: str = ""
    split_days_old: float = 0.0
    split_days_new: float = 0.0
    alert: str = ""


def _normalize_to_canonical(raw_name: str, name_lookup: dict) -> str:
    """
    将任意名字格式转为 config 中定义的标准名 (canonical name)。
    流程: raw → "First Last" → alias lookup → canonical
    """
    if not raw_name:
        return raw_name
    step1 = normalize_name_last_first(raw_name)
    # 通过 alias 查找 canonical
    canonical = name_lookup.get(step1.lower())
    if canonical:
        return canonical
    # 模糊匹配
    for alias, cname in name_lookup.items():
        if alias in step1.lower() or step1.lower() in alias:
            return cname
    # 反向: "First Last" → "Last First" 再试一次
    parts = step1.split()
    if len(parts) == 2:
        reversed_name = f"{parts[1]} {parts[0]}"
        canonical2 = name_lookup.get(reversed_name.lower())
        if canonical2:
            return canonical2
        for alias, cname in name_lookup.items():
            if alias in reversed_name.lower() or reversed_name.lower() in alias:
                return cname
    return step1  # 无法匹配，返回标准化后的名字


# ============================================================
# Phase B1: 工时清单表加载
# ============================================================
def load_gongshi_qingdan(gr_dir: str = None, personnel: dict = None) -> list[POAllocationRecord]:
    """
    加载 IAV Q1/Q2/Q3 工时清单表.xlsx 文件。
    这些文件有显式 Level 列，是最可靠的数据源。
    """
    records = []
    if gr_dir is None:
        gr_dir = getattr(config, 'GR_DOCUMENTS_DIR',
                         os.path.join(config.ROOT, "CNTS-25165", "03_GR Documents"))
    vctc_dir = os.path.join(gr_dir, "2026", "VCTC")

    name_lookup = config.build_name_lookup(personnel) if personnel else {}

    q_files = {
        "Q1": os.path.join(vctc_dir, "Q1", "IAV Q1 工时清单表.xlsx"),
        "Q2": os.path.join(vctc_dir, "Q2", "IAV Q2 工时清单表.xlsx"),
        "Q3": os.path.join(vctc_dir, "Q3", "IAV Q3 工时清单表.xlsx"),
    }

    # Level 简写 → 全称
    level_short_to_full = {
        "senior": "Testing Engineer-Senior",
        "expert": "Testing Engineer-Expert",
        "experienced": "Testing Engineer-Experienced",
        "junior": "Testing Engineer-Junior",
    }

    for quarter, filepath in q_files.items():
        if not os.path.exists(filepath):
            continue

        print(f"  [工时清单表] Loading {quarter}: {os.path.basename(filepath)}")
        # 工时清单表可能有不同 sheet 名称 (Q3 有 typo: "breakdow")
        wb = safe_load_workbook(filepath)
        if not wb:
            continue

        gongshi_sheets = [s for s in wb.sheetnames
                          if 'consultant' in s.lower() and 'cost' in s.lower()]
        wb.close()

        for sheet_name in gongshi_sheets:
            headers, rows = read_sheet_raw(filepath, sheet_name)
            if not headers:
                continue

            # 定位列
            col_name = col_po = col_level = col_rate = None
            col_po_days = col_remaining = None
            month_cols = {}  # {month_str: col_index}

            for i, h in enumerate(headers):
                hs = str(h).strip().lower()
                if hs in ('name', 'fte'):
                    col_name = i
                elif hs in ('po nr', 'po nr.', 'po nr'):
                    col_po = i
                elif hs == 'level':
                    col_level = i
                elif 'rate' in hs:
                    col_rate = i
                elif hs in ('po days', 'po_days'):
                    col_po_days = i
                elif 'remaining' in hs:
                    col_remaining = i
                else:
                    m = parse_date_month(h)
                    if m:
                        month_cols[m] = i

            if col_name is None or col_po is None or not month_cols:
                continue

            record_count = 0
            for row in rows:
                if len(row) <= max(max(month_cols.values()), col_po, col_name):
                    continue

                name_raw = row[col_name]
                po_id = str(row[col_po]).strip() if row[col_po] else ""
                if not name_raw or not po_id or not po_id.strip():
                    continue

                # 标准化名称为 canonical name
                name = _normalize_to_canonical(str(name_raw), name_lookup)
                # 跳过占位符和空名
                if not name or name.lower() in ('xxx', 'tbd', 'n/a', ''):
                    continue
                level = ""
                if col_level is not None and col_level < len(row) and row[col_level]:
                    lv = str(row[col_level]).strip().lower()
                    level = level_short_to_full.get(lv, str(row[col_level]).strip())
                if not level and col_rate is not None and col_rate < len(row):
                    try:
                        rate_val = float(row[col_rate]) if row[col_rate] else 0
                        level = infer_level_from_rate(rate_val, RATE_TO_LEVEL) or ""
                    except (ValueError, TypeError):
                        pass
                # 如果 level 仍为空，从 personnel 配置获取
                if not level and name in (personnel or {}):
                    level = personnel[name].get("level", "")

                # PO Days (配额) 和 Remaining
                quota = 0.0
                remaining = 0.0
                if col_po_days is not None and col_po_days < len(row):
                    try:
                        quota = float(row[col_po_days]) if row[col_po_days] else 0.0
                    except (ValueError, TypeError):
                        pass
                if col_remaining is not None and col_remaining < len(row):
                    try:
                        remaining = float(row[col_remaining]) if row[col_remaining] else 0.0
                    except (ValueError, TypeError):
                        pass

                rate_val = 0.0
                if col_rate is not None and col_rate < len(row):
                    try:
                        rate_val = float(row[col_rate]) if row[col_rate] else 0.0
                    except (ValueError, TypeError):
                        pass

                for month_str, col in sorted(month_cols.items()):
                    if col >= len(row):
                        continue
                    val = row[col]
                    if val is None:
                        continue
                    try:
                        days = float(val)
                    except (ValueError, TypeError):
                        continue
                    if days > 0:
                        records.append(POAllocationRecord(
                            person_name=name, month=month_str, po_id=po_id,
                            days=days, rate=rate_val, level=level,
                            source=f"工时清单表_{quarter}",
                            po_days_quota=quota, po_days_remaining=remaining,
                        ))
                        record_count += 1

            print(f"    Sheet '{sheet_name}': {record_count} allocation records")

    return records


# ============================================================
# Phase B2: 外援情况加载 (Priority 1b — GROUND TRUTH, Level from Rate)
# ============================================================
def load_waiyuan_qingkuang(gr_dir: str = None, personnel: dict = None) -> list[POAllocationRecord]:
    """
    加载外援情况文件 (Q1: AI IAV 外援情况_26Q1.xlsx, Q2: Cost Breakdown_26Q2)

    文件格式:
      Supervisor | Supplier | FTE | Role Rate (Net) | PO Nr. | GR Nr. |
      total GR day | total GR Amount | PO Days | GRed Days | Remaining |
      2026-Jan | 2026-Feb | 2026-Mar | Total GR mandays | GR Nr | GR Amount | Remark

    FTE 列: "Tang, Qiwei" 格式
    月份列: "2026-Jan", "2026-Feb" 等
    无显式 Level 列 → 从 Role Rate 推断
    """
    records = []
    if gr_dir is None:
        gr_dir = getattr(config, 'GR_DOCUMENTS_DIR',
                         os.path.join(config.ROOT, "CNTS-25165", "03_GR Documents"))
    vctc_dir = os.path.join(gr_dir, "2026", "VCTC")

    name_lookup = config.build_name_lookup(personnel) if personnel else {}

    waiyuan_files = [
        # (quarter_label, path, default_sheet)
        ("Q1", os.path.join(vctc_dir, "Q1", "AI IAV 外援情况_26Q1.xlsx"), "26Q1"),
        ("Q2", os.path.join(vctc_dir, "Q2", "Cost Breakdown_26Q2_260608.xlsx"), "26Q2"),
    ]

    for quarter, filepath, default_sheet in waiyuan_files:
        if not os.path.exists(filepath):
            continue

        print(f"  [外援情况] Loading {quarter}: {os.path.basename(filepath)}")

        # 查找含月份数据的 sheet
        wb = safe_load_workbook(filepath)
        if not wb:
            continue

        candidate_sheets = [s for s in wb.sheetnames
                           if s.startswith("26Q") and 'Sheet' not in s]
        wb.close()

        sheets_to_load = candidate_sheets if candidate_sheets else [default_sheet]

        for sheet_name in sheets_to_load:
            try:
                headers, rows = read_sheet_raw(filepath, sheet_name)
            except Exception:
                continue
            if not headers:
                continue

            col_name = col_po = col_rate = None
            col_po_days = col_remaining = None
            month_cols = {}

            for i, h in enumerate(headers):
                hs = str(h).strip().lower()
                if hs in ('fte', 'name'):
                    col_name = i
                elif hs in ('po nr.', 'po nr', 'po no.'):
                    col_po = i
                elif 'rate' in hs and 'role' in hs:
                    col_rate = i
                elif hs in ('po days',):
                    col_po_days = i
                elif hs == 'remaining':
                    col_remaining = i
                else:
                    m = parse_date_month(str(h).strip())
                    if m:
                        month_cols[m] = i

            if col_name is None or col_po is None or not month_cols:
                continue

            record_count = 0
            for row in rows:
                if len(row) <= max(max(month_cols.values()), col_po, col_name):
                    continue

                name_raw = row[col_name]
                po_id = str(row[col_po]).strip() if row[col_po] else ""
                if not name_raw or not po_id:
                    continue

                name = _normalize_to_canonical(str(name_raw), name_lookup)
                if not name or name.lower() in ('xxx', 'tbd', 'n/a', ''):
                    continue

                # Rate → Level
                rate_val = 0.0
                if col_rate is not None and col_rate < len(row):
                    try:
                        rate_val = float(row[col_rate]) if row[col_rate] else 0.0
                    except (ValueError, TypeError):
                        pass
                level = infer_level_from_rate(rate_val, RATE_TO_LEVEL) or ""

                quota = 0.0
                remaining = 0.0
                if col_po_days is not None and col_po_days < len(row):
                    try:
                        quota = float(row[col_po_days]) if row[col_po_days] else 0.0
                    except (ValueError, TypeError):
                        pass
                if col_remaining is not None and col_remaining < len(row):
                    try:
                        remaining = float(row[col_remaining]) if row[col_remaining] else 0.0
                    except (ValueError, TypeError):
                        pass

                for month_str, col in sorted(month_cols.items()):
                    if col >= len(row):
                        continue
                    val = row[col]
                    if val is None:
                        continue
                    try:
                        days = float(val)
                    except (ValueError, TypeError):
                        continue
                    if days > 0:
                        records.append(POAllocationRecord(
                            person_name=name, month=month_str, po_id=po_id,
                            days=days, rate=rate_val, level=level,
                            source=f"外援情况_{quarter}",
                            po_days_quota=quota, po_days_remaining=remaining,
                        ))
                        record_count += 1

            print(f"    Sheet '{sheet_name}': {record_count} allocation records")

    return records


# ============================================================
# Phase B3: Budget Report 加载 (Priority 2 — 补缺 + 预测)
# ============================================================
def load_budget_report_actuals_and_forecast(
    project_dir: str, personnel: dict, project_id: str | None = None
) -> tuple[dict, dict]:
    """
    从 Budget Report 的 Internal Budget_FC sheet 读取实际和预测数据。

    Sheet 结构 (Internal Budget_FC 2026 Q2):
      - Rows 1-8: PO 摘要
      - Rows 9-29: Forecast 区（每人每月计划天数）
      - Rows 31+: 分隔/子标题行
      - Rows 33-54: Actual 区（每人每月实际天数 + 实际费率）

    Forecast 区列结构 (从 Row 9 header + Row 32 sub-header 推断):
      Col 9: Name ("Tang Qiwei" 格式，但可能 "Tang, Qiwei")
      Col 10: VCTC Rate
      Col 11: Level
      Col 22-33: 2026年1-12月
      Col 34: Total

    Returns:
      actuals:  {name: {month_str: days}}
      forecast: {name: {month_str: days}}
    """
    actuals = defaultdict(lambda: defaultdict(float))
    forecast = defaultdict(lambda: defaultdict(float))

    filepath = _latest_active_budget_report(project_id, project_dir)
    if not filepath:
        print("  [WARN] Budget Report not found")
        return dict(actuals), dict(forecast)

    # 读取原始数据
    wb = safe_load_workbook(filepath)
    if not wb:
        return dict(actuals), dict(forecast)

    sheet_name = None
    for sn in wb.sheetnames:
        if 'Internal Budget' in sn and '2026' in sn:
            sheet_name = sn
            break
    if not sheet_name:
        wb.close()
        return dict(actuals), dict(forecast)

    ws = wb[sheet_name]
    all_rows = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
        all_rows.append(list(row))
    wb.close()

    # 构建名称查找表
    name_lookup = config.build_name_lookup(personnel)

    def normalize_br_name(raw):
        """Budget Report 中名字格式: 'Tang, Qiwei' 或 'Tang Qiwei'"""
        if not raw:
            return None
        n = str(raw).strip()
        # "Last, First" → "First Last"
        if ',' in n:
            parts = n.split(',', 1)
            n = f"{parts[1].strip()} {parts[0].strip()}"
        # 通过 alias 匹配 canonical name
        canonical = name_lookup.get(n.lower())
        if canonical:
            return canonical
        # 模糊匹配
        for alias, cname in name_lookup.items():
            if alias in n.lower() or n.lower() in alias:
                return cname
        return n

    def parse_br_section(rows_slice, section_label):
        """解析 Budget Report 的一个数据段"""
        result = defaultdict(lambda: defaultdict(float))
        if not rows_slice:
            return result

        # 找到列映射: 在数据行中找含月份值的列
        month_col_map = {}
        col_name = col_level = 9  # 默认列
        col_level_idx = 11

        for i in range(len(rows_slice)):
            row = rows_slice[i]
            # 尝试在每行中识别 Name 和 Level
            for j, val in enumerate(row):
                if val is None:
                    continue
                vs = str(val).strip()
                parsed = parse_date_month(vs)
                if parsed:
                    # 检查上下文：附近行是否有数据来确认年月
                    month_col_map[j] = parsed

            if month_col_map:
                break

        # 如果没有找到月份列，使用固定映射 (从 row 9 header)
        if not month_col_map:
            # 默认: col 22 = Jan, 23 = Feb, ..., 33 = Dec
            for m in range(1, 13):
                month_col_map[22 + m - 1] = f"2026-{m:02d}"

        for row in rows_slice:
            # 跳过分隔行和汇总行
            name_val = None
            level_val = ""
            for j in range(min(len(row), 12)):
                val = row[j] if j < len(row) else None
                if val is None:
                    continue
                vs = str(val).strip()
                if not vs:
                    continue

            # Name 在 col 9
            if len(row) > col_name and row[col_name]:
                name_val = row[col_name]
            if len(row) > col_level_idx and row[col_level_idx]:
                level_val = str(row[col_level_idx]).strip()

            if not name_val:
                # 尝试在其他位置找名字
                for j in range(8, 12):
                    if j < len(row) and row[j] and not isinstance(row[j], (int, float)):
                        candidate = str(row[j]).strip()
                        if candidate and not candidate.startswith('EM') and \
                           candidate.lower() not in ('senior', 'expert', 'experienced', 'junior'):
                            # 检查是否匹配已知人员
                            for alias in name_lookup:
                                if alias.replace(',', '') in candidate.lower().replace(',', ''):
                                    name_val = candidate
                                    break
                            if name_val:
                                break

            if not name_val:
                continue

            canonical = normalize_br_name(str(name_val))
            if not canonical or canonical not in personnel:
                continue

            for col, month_str in month_col_map.items():
                if col >= len(row):
                    continue
                try:
                    days = float(row[col]) if row[col] is not None else 0.0
                except (ValueError, TypeError):
                    days = 0.0
                if days > 0:
                    result[canonical][month_str] += days

        return result

    # 查找 Forecast 段 (rows 9-29, 约 0-indexed 8-28)
    forecast_start = None
    actuals_start = None
    for i, row in enumerate(all_rows):
        if not row:
            continue
        first_col = str(row[9]) if len(row) > 9 and row[9] else ""
        if 'Forecast' in str(first_col) or 'VCTC Rate' in str(first_col):
            if forecast_start is None:
                forecast_start = i + 1  # 下一行开始是数据
        if 'Actual' in str(first_col) or 'Cost Rate' in str(first_col):
            if actuals_start is None and forecast_start is not None:
                actuals_start = i + 1

    # Fallback: 固定行范围
    if forecast_start is None:
        forecast_start = 9  # 0-indexed row 9
    if actuals_start is None:
        actuals_start = 33  # 0-indexed row 33

    fc_rows = all_rows[forecast_start:]
    ac_rows = all_rows[actuals_start:]

    forecast = parse_br_section(fc_rows, "Forecast")
    actuals = parse_br_section(ac_rows, "Actual")

    fc_people = len(forecast)
    ac_people = len(actuals)
    print(f"  [Budget Report] Forecast: {fc_people} people, Actual: {ac_people} people")

    return dict(actuals), dict(forecast)


def _latest_active_budget_report(project_id: str | None, project_dir: str) -> str | None:
    """Use the project registry's shared active-report rules for every writer."""
    if project_id:
        try:
            from project_registry import build_profile, latest_budget_report
            report = latest_budget_report(build_profile(project_id))
            if report is not None:
                return str(report)
        except (KeyError, OSError, ValueError) as exc:
            print(f"  [WARN] Project registry could not locate the active Budget Report: {exc}")

    import glob as glob_mod
    excluded = ("template", "backup", "_diff", "~$", "po_usage_status", "_working")
    candidates = [
        path for path in glob_mod.glob(os.path.join(project_dir, "IAV_Budget_Report_*.xlsx"))
        if not any(token in os.path.basename(path).casefold() for token in excluded)
    ]
    return max(candidates, key=os.path.getmtime) if candidates else None


# ============================================================
# Phase B4: Timesheet + Monthly Closing (Priority 3 & 4)
# ============================================================
def load_timesheet_days(project_dir: str, personnel: dict) -> dict:
    """
    Priority 3: 第三方人员工时 PDF
    """
    name_lookup = config.build_name_lookup(personnel)
    timesheet_dir = os.path.join(
        os.path.dirname(project_dir), "01_IAV TMS & WP"
    )
    if not os.path.isdir(timesheet_dir):
        return {}

    hours_raw = load_timesheet_hours_from_dir(timesheet_dir)
    days_result = defaultdict(lambda: defaultdict(float))
    for name, months in hours_raw.items():
        canonical = name_lookup.get(name.lower(), name)
        for m, h in months.items():
            days_result[canonical][m] += h / HOURS_PER_DAY
    return dict(days_result)


def load_timesheet_hours_from_dir(base_dir: str, months: list = None) -> dict:
    """Wrapper around timesheet_parser (existing functionality)."""
    records = scan_timesheet_directory(base_dir, months)
    return records_to_summary(records)


def load_monthly_closing_for_verification(project_wbs: str, personnel: dict) -> dict:
    """
    Priority 4 (仅核查): Monthly Closing 数据库
    返回: {name: {month_str: hours}}
    """
    name_lookup = config.build_name_lookup(personnel)
    result = defaultdict(lambda: defaultdict(float))

    pattern = config.MONTHLY_CLOSING_PATTERN
    filepath = find_latest_file(os.path.dirname(pattern), os.path.basename(pattern))
    if not filepath:
        print("  [WARN] Monthly Closing file not found (verification skipped)")
        return dict(result)

    wb = safe_load_workbook(filepath)
    if not wb:
        return dict(result)

    for sheet_name in wb.sheetnames:
        month_num = None
        for name_part, num in MONTH_NAMES_MAP.items():
            if sheet_name.lower().startswith(name_part.lower()):
                month_num = num
                break
        if month_num is None:
            continue

        month_str = f"2026-{month_num:02d}"
        ws = wb[sheet_name]

        header_row = [str(c.value) if c.value else "" for c in
                      next(ws.iter_rows(min_row=2, max_row=2))]
        col_wbs = 10  # Project Number column
        col_hours = None
        col_last = 0
        col_first = 1

        for i, h in enumerate(header_row):
            hl = h.lower()
            if 'project' in hl and 'number' in hl:
                col_wbs = i
            elif 'total' in hl and 'quantity' in hl:
                col_hours = i
            elif 'receiver' in hl and 'wbs' in hl and col_wbs == 10:
                col_wbs = i

        if col_hours is None:
            continue

        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
            if len(row) <= max(col_wbs, col_hours):
                continue
            wbs_val = str(row[col_wbs]) if col_wbs < len(row) and row[col_wbs] else ""
            if project_wbs not in wbs_val:
                continue

            try:
                hours = float(row[col_hours]) if row[col_hours] else 0.0
            except (ValueError, TypeError):
                hours = 0.0
            if hours <= 0:
                continue

            last = str(row[col_last]).strip() if len(row) > col_last and row[col_last] else ""
            first = str(row[col_first]).strip() if len(row) > col_first and row[col_first] else ""
            name_raw = f"{first} {last}".strip() if first or last else ""
            if not name_raw:
                continue

            canonical = None
            for alias, cname in name_lookup.items():
                if alias in name_raw.lower() or name_raw.lower() in alias:
                    canonical = cname
                    break
            if not canonical:
                continue

            result[canonical][month_str] += hours

    wb.close()
    mc_people = len(result)
    print(f"  [Monthly Closing] {mc_people} matched people (verification only)")
    return dict(result)


# ============================================================
# Phase C1: 数据合并 (Ground Truth → Gap Fill)
# ============================================================
def merge_ground_truth_allocation(
    gongshi_records: list[POAllocationRecord],
    waiyuan_records: list[POAllocationRecord],
    budget_actuals: dict,
    timesheet_days: dict,
    personnel: dict,
    po_configs: list[POLevelConfig],
    project_id: str = "CNTS-25165",
) -> dict:
    """
    按优先级合并所有数据源 → {name: {month: {po_id: days}}}

    Priority:
      1. 工时清单表 (最优先，有 Level 列)
      2. 外援情况 (补缺，Level 从 Rate 推断)
      3. Budget Report Actuals (补缺，无 PO 归属 → 按 Level 匹配当前活跃 PO)
      4. Timesheet PDFs (第三方补充)
    """
    # 构建 Level→Priority→PO 映射 (用于补缺时确定活跃 PO)
    level_po_map = defaultdict(list)  # {level: [(priority, start_month, po_id), ...]}
    for c in po_configs:
        level_po_map[c.level].append((c.priority, c.start_month, c.po_id))
    for lv in level_po_map:
        level_po_map[lv].sort()

    merged = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    source_tracker = {}  # {(name, month, po_id): source}

    # --- Priority 1: 工时清单表 ---
    for rec in gongshi_records:
        key = (rec.person_name, rec.month, rec.po_id)
        merged[rec.person_name][rec.month][rec.po_id] = rec.days
        source_tracker[key] = rec.source

    # --- Priority 2: 外援情况 (不覆盖 P1) ---
    for rec in waiyuan_records:
        key = (rec.person_name, rec.month, rec.po_id)
        if key in source_tracker:
            continue  # P1 已有
        merged[rec.person_name][rec.month][rec.po_id] = rec.days
        source_tracker[key] = rec.source

    # --- Priority 3: Budget Report Actuals (补缺：找到该人该月有没有任何 PO 分配) ---
    br_filled = 0
    for name, months in budget_actuals.items():
        if name not in personnel:
            continue
        person_level = personnel[name].get("level", "")
        for month_str, days in months.items():
            # 检查该人该月是否已有 PO 分配
            existing = sum(merged[name][month_str].values()) if name in merged else 0
            if existing > 0:
                continue  # 已有 P1/P2 数据，不覆盖

            # 找该 Level 在该月的活跃 PO
            active_po = _find_active_po_for_month(person_level, month_str, level_po_map)
            if active_po:
                merged[name][month_str][active_po] = days
                source_tracker[(name, month_str, active_po)] = "Budget_Report_Inferred"
                br_filled += 1

    if br_filled:
        print(f"  [Merge] Budget Report gap-filled: {br_filled} person-month entries")

    # --- Priority 4: Timesheet PDFs (第三方补充) ---
    ts_filled = 0
    for name, months in timesheet_days.items():
        if name not in personnel:
            continue
        person_level = personnel[name].get("level", "")
        for month_str, days in months.items():
            existing = sum(merged[name][month_str].values()) if name in merged else 0
            if existing > 0:
                continue
            active_po = _find_active_po_for_month(person_level, month_str, level_po_map)
            if active_po:
                merged[name][month_str][active_po] = days
                source_tracker[(name, month_str, active_po)] = "Timesheet_PDF"
                ts_filled += 1

    if ts_filled:
        print(f"  [Merge] Timesheet gap-filled: {ts_filled} person-month entries")

    # The Review Center's approved facts are the canonical, PO-specific result
    # of reviewing scanned Timesheet PDFs. The legacy folder scan does not see
    # facts that were reviewed in the Workbench, so read the current approvals
    # from the shared fact store and apply them last. Same person/month/PO keys
    # replace lower-priority workbook values instead of being double counted.
    from fact_store import load_approved_timesheet_facts
    approved_facts = load_approved_timesheet_facts(project_id)
    if approved_facts:
        name_lookup = config.build_name_lookup(personnel)
        compact_name_lookup = {}
        for canonical, info in personnel.items():
            for alias in (canonical, *(info.get("aliases") or [])):
                token = re.sub(r"[^a-z0-9]", "", str(alias).casefold())
                if token:
                    compact_name_lookup.setdefault(token, canonical)
        approved_days = defaultdict(float)
        approved_hours = 0.0
        skipped_unmapped = 0
        skipped_invalid = 0
        for fact in approved_facts:
            raw_name = str(fact.get("canonical_name") or fact.get("person_key") or "").strip()
            person_name = raw_name if raw_name in personnel else name_lookup.get(raw_name.lower())
            if not person_name:
                token = re.sub(r"[^a-z0-9]", "", raw_name.casefold())
                person_name = compact_name_lookup.get(token)
            if not person_name:
                skipped_unmapped += 1
                continue
            month = str(fact.get("month") or "").strip()
            po_id = str(fact.get("po_id") or "").strip()
            try:
                hours = float(fact.get("hours") or 0)
            except (TypeError, ValueError):
                hours = 0.0
            if not month or not po_id or not math.isfinite(hours) or hours <= 0:
                skipped_invalid += 1
                continue
            approved_days[(person_name, month, po_id)] += hours / HOURS_PER_DAY
            approved_hours += hours

        overlaps = 0
        for (person_name, month, po_id), days in approved_days.items():
            existing = merged[person_name][month].get(po_id)
            if existing is not None:
                overlaps += 1
            merged[person_name][month][po_id] = round(days, 8)
            source_tracker[(person_name, month, po_id)] = "Approved_Timesheet_Fact"

        print(
            f"  [Merge] Approved Review Center facts: {len(approved_facts)} records, "
            f"{len(approved_days)} person-month-PO keys, {approved_hours:.1f} hours; "
            f"{overlaps} existing keys replaced, {skipped_unmapped} names unmapped, "
            f"{skipped_invalid} invalid records skipped"
        )
    else:
        print(
            "  [Review Center] 0 current approved Timesheet facts loaded "
            f"from {config.FACT_STORE_PATH}; only configured source fallbacks are available"
        )

    # 统计
    total_people = len(merged)
    total_entries = sum(
        sum(len(po_days) for po_days in months.values())
        for months in merged.values()
    )
    print(f"  [Merge] Final: {total_people} people, {total_entries} person-month-PO entries")

    return dict(merged)


def _find_active_po_for_month(level, month_str, level_po_map):
    """找到某个 Level 在指定月份的活跃 PO (按 Priority 第一个覆盖该月的)"""
    if level not in level_po_map:
        return None
    for priority, start_month, po_id in level_po_map[level]:
        if start_month <= month_str:
            # 检查是否还活跃 (简单判断：不是所有 later PO 都覆盖)
            pass
    # 找最后一个 start_month <= month_str 的
    best = None
    for priority, start_month, po_id in sorted(level_po_map.get(level, [])):
        if start_month <= month_str:
            best = po_id
    return best


# ============================================================
# Phase C2: Monthly Closing 验证
# ============================================================
def verify_with_monthly_closing(
    merged_allocation: dict,
    mc_hours: dict,
    personnel: dict,
    threshold: float = MC_THRESHOLD,
) -> list[dict]:
    """
    用 Monthly Closing 数据验证合并结果 (只出警告，不修改)。
    只对超过阈值且绝对差异 >= 3 天的情况发出警告。
    """
    warnings = []
    for name, months_hours in mc_hours.items():
        if name not in personnel:
            continue
        for month_str, hours in months_hours.items():
            mc_days = hours / HOURS_PER_DAY
            merged_days = 0.0
            if name in merged_allocation and month_str in merged_allocation[name]:
                merged_days = sum(merged_allocation[name][month_str].values())

            if mc_days <= 0 and merged_days <= 0:
                continue

            max_val = max(mc_days, merged_days)
            if max_val <= 0:
                continue

            diff_pct = abs(mc_days - merged_days) / max_val
            abs_diff = abs(mc_days - merged_days)

            # 需要同时满足: 差异 > 阈值 AND 绝对差异 >= 3 天
            if diff_pct > threshold and abs_diff >= 3.0:
                severity = "WARNING" if diff_pct < 0.50 else "SEVERE"
                warnings.append({
                    "name": name, "month": month_str,
                    "merged_days": round(merged_days, 1),
                    "mc_days": round(mc_days, 1),
                    "diff_pct": round(diff_pct * 100, 1),
                    "severity": severity,
                })

    if warnings:
        severe_count = sum(1 for w in warnings if w["severity"] == "SEVERE")
        print(f"  [Verify] MC check: {len(warnings)} warnings ({severe_count} severe)")
    else:
        print(f"  [Verify] MC check: All within threshold ({threshold*100:.0f}%)")

    return warnings


# ============================================================
# Phase D1: PO 配置处理
# ============================================================
def load_po_level_config(project: str = "CNTS-25165") -> list[POLevelConfig]:
    """加载 PO Level 配置 (优先 Excel，其次硬编码)。project 参数用于筛选。"""
    configs = []

    excel_path = os.path.join(config.INPUT_DIR, "po_config_v2.xlsx")
    if os.path.exists(excel_path):
        headers, rows = read_sheet(excel_path, "PO_Level_Config")
        if headers:
            # 检测格式: 新格式有 Project 列在最前面
            has_project_col = headers and 'project' in str(headers[0]).lower()
            for row in rows:
                if len(row) < 6:
                    continue
                try:
                    if has_project_col:
                        row_project = str(row[0]).strip() if row[0] else ""
                        offset = 1
                    else:
                        row_project = "CNTS-25165"  # 旧格式默认 CNTS-25165
                        offset = 0

                    # 按 project 参数筛选 (空字符串 = 全部)
                    if project and row_project and row_project != project:
                        continue

                    po_id = str(row[offset + 0]).strip()
                    level = str(row[offset + 1]).strip()
                    days_quota = float(row[offset + 2]) if row[offset + 2] else 0
                    unit_price = float(row[offset + 3]) if row[offset + 3] else 0
                    value_rmb = float(row[offset + 4]) if len(row) > offset + 4 and row[offset + 4] else days_quota * unit_price
                    start_month = str(row[offset + 5]).strip() if len(row) > offset + 5 and row[offset + 5] else ""
                    end_month = str(row[offset + 6]).strip() if len(row) > offset + 6 and row[offset + 6] else ""
                    priority = int(row[offset + 7]) if len(row) > offset + 7 and row[offset + 7] else 1
                    allocation_mode = str(row[offset + 8]).strip() if len(row) > offset + 8 and row[offset + 8] else "usage"

                    if po_id and level and days_quota > 0:
                        configs.append(POLevelConfig(
                            row_project, po_id, level, days_quota, unit_price,
                            value_rmb, start_month, end_month, priority,
                            allocation_mode,
                        ))
                except (ValueError, IndexError) as e:
                    print(f"  [WARN] Skipping po_config_v2 row: {row[:6]} — {e}")

    if not configs and project == "CNTS-25165":
        configs = _get_default_cnst25165_config()

    return configs


def _get_default_cnst25165_config() -> list[POLevelConfig]:
    """CNTS-25165 默认配置 (基于客户 PO PDF 实际配额)"""
    return [
        POLevelConfig("CNTS-25165", "2601000336", "Testing Engineer-Expert", 132, 2460, 324720,
                      "2026-01", "2026-04", 1),
        POLevelConfig("CNTS-25165", "2601000337", "Testing Engineer-Expert", 132, 2460, 324720,
                      "2026-01", "2026-04", 1),
        POLevelConfig("CNTS-25165", "2601000337", "Testing Engineer-Senior", 132, 2260, 298320,
                      "2026-01", "2026-04", 1),
        POLevelConfig("CNTS-25165", "2601000338", "Testing Engineer-Expert", 66, 2460, 162360,
                      "2026-01", "2026-04", 1),
        POLevelConfig("CNTS-25165", "2601000338", "Testing Engineer-Experienced", 330, 2100, 693000,
                      "2026-01", "2026-04", 1),
        POLevelConfig("CNTS-25165", "2604000799", "Testing Engineer-Experienced", 350, 2100, 735000,
                      "2026-04", "2026-09", 2),
        POLevelConfig("CNTS-25165", "2604000799", "Testing Engineer-Senior", 280, 2260, 632800,
                      "2026-04", "2026-09", 2),
        POLevelConfig("CNTS-25165", "2604000799", "Testing Engineer-Expert", 210, 2460, 516600,
                      "2026-04", "2026-09", 2),
        POLevelConfig("CNTS-25165", "2607000312", "Testing Engineer-Experienced", 180, 2100, 378000,
                      "2026-07", "2026-12", 3),
        POLevelConfig("CNTS-25165", "2607000312", "Testing Engineer-Senior", 85, 2260, 192100,
                      "2026-07", "2026-12", 3),
        POLevelConfig("CNTS-25165", "2607000312", "Testing Engineer-Expert", 464, 2460, 1141440,
                      "2026-07", "2026-12", 3),
    ]


def resolve_po_start_end_gaps(
    po_configs: list[POLevelConfig]
) -> tuple[list[POLevelConfig], list[str]]:
    """
    自动填补缺失的 start_month / end_month。
    按 Priority 排序后，缺失的 start 从前一个 PO 的 end+1 月推断。
    """
    alerts = []
    if not po_configs:
        return po_configs, alerts

    # 按 Priority 分组 (所有 Level)
    priority_groups = defaultdict(list)
    for c in po_configs:
        priority_groups[c.priority].append(c)

    sorted_priorities = sorted(priority_groups.keys())
    prev_end = None

    for i, p in enumerate(sorted_priorities):
        configs = priority_groups[p]
        # 取该 priority 组中第一个有效的 end_month
        group_end = None
        group_start = None
        for c in configs:
            if c.end_month and (group_end is None or c.end_month > group_end):
                group_end = c.end_month
            if c.start_month and (group_start is None or c.start_month < group_start):
                group_start = c.start_month

        for c in configs:
            if not c.start_month and prev_end:
                # 推断: prev_end + 1 month
                y, m = map(int, prev_end.split('-'))
                m += 1
                if m > 12:
                    y += 1
                    m = 1
                c.start_month = f"{y}-{m:02d}"
                msg = (f"  [GAP FILL] PO {c.po_id} ({c.level}): "
                       f"start_month auto-set to {c.start_month}")
                print(msg)
                alerts.append(msg)
            elif not c.start_month:
                c.start_month = "2026-01"
                msg = (f"  [GAP FILL] PO {c.po_id} ({c.level}): "
                       f"start_month defaulted to {c.start_month}")
                print(msg)
                alerts.append(msg)

            if not c.end_month:
                # 找下一个 priority 的 start_month
                next_start = None
                for np in sorted_priorities:
                    if np > p:
                        for nc in priority_groups[np]:
                            if nc.start_month:
                                next_start = nc.start_month
                                break
                        if next_start:
                            break
                if next_start:
                    # end = next_start - 1 month
                    y, m = map(int, next_start.split('-'))
                    m -= 1
                    if m < 1:
                        y -= 1
                        m = 12
                    c.end_month = f"{y}-{m:02d}"
                elif c.start_month:
                    # 无后续 PO → 从 start_month 起推 6 个月
                    y, m = map(int, c.start_month.split('-'))
                    m += 5
                    if m > 12:
                        y += 1
                        m -= 12
                    c.end_month = f"{y}-{m:02d}"
                else:
                    c.end_month = "2026-12"
                msg = (f"  [GAP FILL] PO {c.po_id} ({c.level}): "
                       f"end_month auto-set to {c.end_month}")
                print(msg)
                alerts.append(msg)

            # 确保 end_month >= start_month
            if c.start_month and c.end_month and c.end_month < c.start_month:
                y, m = map(int, c.start_month.split('-'))
                m += 5
                if m > 12:
                    y += 1
                    m -= 12
                c.end_month = f"{y}-{m:02d}"
                msg = (f"  [FIX] PO {c.po_id} ({c.level}): "
                       f"end_month adjusted from before start to {c.end_month}")
                print(msg)
                alerts.append(msg)

        if group_end:
            prev_end = group_end

    return po_configs, alerts


# ============================================================
# Phase D2: PO-Level 库存快照
# ============================================================
def build_po_level_inventory(
    po_configs: list[POLevelConfig],
    merged_allocation: dict,
    personnel: dict,
) -> dict:
    """
    构建 PO-Level 库存快照。
    使用人员的 Level 正确归属消耗到对应配额。
    返回: {po_id: {level: {quota, consumed, remaining, usage_pct, ...}}}
    """
    inventory = {}

    # 初始化所有 PO-Level
    for c in po_configs:
        if c.po_id not in inventory:
            inventory[c.po_id] = {}
        inventory[c.po_id][c.level] = {
            "quota": c.days_quota,
            "consumed": 0.0,
            "remaining": c.days_quota,
            "unit_price": c.unit_price,
            "value_rmb": c.value_rmb if c.value_rmb else c.days_quota * c.unit_price,
            "start_month": c.start_month,
            "end_month": c.end_month,
            "priority": c.priority,
            "monthly_consumed": defaultdict(float),
        }

    # 汇总实际消耗 — 按人员的 Level 归属到正确的 PO-Level 配额
    for name, months in merged_allocation.items():
        person_level = personnel.get(name, {}).get("level", "")
        if not person_level:
            continue
        for month_str, po_days in months.items():
            for po_id, days in po_days.items():
                if po_id not in inventory:
                    continue
                if person_level not in inventory[po_id]:
                    continue
                if days > 0:
                    inventory[po_id][person_level]["consumed"] += days
                    inventory[po_id][person_level]["monthly_consumed"][month_str] += days

    # 计算 remaining 和 usage%
    for po_id in inventory:
        for level in inventory[po_id]:
            info = inventory[po_id][level]
            info["consumed"] = round(info["consumed"], 1)
            info["remaining"] = round(info["quota"] - info["consumed"], 1)
            info["usage_pct"] = round(
                info["consumed"] / info["quota"] * 100, 1
            ) if info["quota"] > 0 else 0.0

    return inventory


# ============================================================
# Phase D3: Level 月消耗率计算
# ============================================================
def calculate_level_consumption_rates(
    merged_allocation: dict,
    personnel: dict,
    inventory: dict,
    lookback: int = LOOKBACK,
) -> dict[str, LevelConsumption]:
    """
    计算每个 Level 的实际月消耗率 (基于最近 N 个月的数据)
    """
    # 聚合每个 Level 每月总天数
    level_monthly = defaultdict(lambda: defaultdict(float))

    for name, months in merged_allocation.items():
        if name not in personnel:
            continue
        level = personnel[name].get("level", "")
        if not level:
            continue
        for month_str, po_days in months.items():
            level_monthly[level][month_str] += sum(po_days.values())

    # 计算每个 Level 的统计
    result = {}
    available_months = sorted(set(
        m for months in level_monthly.values() for m in months.keys()
    ))

    for level, monthly in level_monthly.items():
        lc = LevelConsumption(level=level, monthly_totals=dict(monthly))

        # 活跃人数
        active = [n for n, info in personnel.items()
                  if info.get("level") == level]
        lc.active_people = active

        # 最近 N 月平均值
        recent_months = available_months[-lookback:] if len(available_months) >= lookback else available_months
        recent_values = [monthly.get(m, 0) for m in recent_months]
        if recent_values:
            lc.latest_monthly_rate = sum(recent_values) / len(recent_values)
        total = sum(monthly.values())
        num = len(monthly)
        lc.monthly_avg = total / num if num > 0 else 0

        result[level] = lc

    return result


# ============================================================
# Phase D3.5: 人员级分配 (Person-Level Allocation)
# ============================================================

@dataclass
class PersonMonthAllocation:
    """某人某月的 PO 分配"""
    person_name: str
    month: str
    level: str
    total_days: float
    po_days: dict          # {po_id: days}
    transition: bool = False
    historical_rate: float = 0.0  # 历史月均天数


def compute_person_historical_rates(
    merged_allocation: dict,
    personnel: dict,
    lookback_months: int = 3,
) -> dict:
    """
    计算每人每月的历史平均消耗天数（用于分配权重）。

    Args:
        merged_allocation: {name: {month: {po_id: days}}}
        personnel: 人员配置
        lookback_months: 回溯窗口

    Returns:
        {name: avg_days_per_month}
    """
    rates = {}
    for name, person_info in personnel.items():
        months_data = merged_allocation.get(name, {})
        if not months_data:
            rates[name] = 0.0
            continue

        monthly_totals = {}
        for month_str, po_days in months_data.items():
            monthly_totals[month_str] = sum(po_days.values())

        all_months = sorted(monthly_totals.keys())
        recent_months = all_months[-lookback_months:] if len(all_months) >= lookback_months else all_months
        recent_values = [monthly_totals.get(m, 0) for m in recent_months]
        rates[name] = sum(recent_values) / len(recent_values) if recent_values else 0.0

    return rates


def allocate_level_to_persons(
    level: str,
    month: str,
    monthly_demand: float,
    active_po_id: str,
    personnel: dict,
    historical_rates: dict,
    transition_info: dict = None,
    workdays: int = 22,
) -> list[PersonMonthAllocation]:
    """
    将某个 Level 的月度总需求分配到每个人员。

    算法:
    1. 筛选该 Level 在目标月前已入职的人员
    2. 每人权重 = 个人历史月均 / 所有人的历史月均总和
    3. 每人分配天数 = monthly_demand * 权重
    4. 最后一人做舍入调整，确保总和 = monthly_demand
    5. 过渡月份：按权重从大到小，先用旧PO再用新PO

    Args:
        level: Level 名称
        month: 目标月 "2026-08"
        monthly_demand: 该 Level 当月总需求天数
        active_po_id: 当前活跃 PO
        personnel: 人员配置
        historical_rates: {name: avg_days} from compute_person_historical_rates()
        transition_info: {from_po, to_po, split_old, split_new} or None

    Returns:
        list[PersonMonthAllocation]
    """
    # 1. 筛选该 Level 的活跃人员
    active_people = []
    for name, info in personnel.items():
        if info.get("level") != level:
            continue
        join_month = info.get("join_month", "2026-01")
        if join_month <= month:
            active_people.append(name)

    if not active_people:
        return []

    # 2. 计算权重
    total_rate = sum(historical_rates.get(p, 0) for p in active_people)
    if total_rate <= 0:
        # 所有人权重相同
        weights = {p: 1.0 / len(active_people) for p in active_people}
    else:
        weights = {}
        for p in active_people:
            rate = historical_rates.get(p, 0)
            weights[p] = rate / total_rate if rate > 0 else 0.5 / len(active_people)

    # 3. 分配天数（不含过渡拆分）
    raw_allocations = {}
    remaining_demand = monthly_demand
    sorted_people = sorted(active_people, key=lambda p: weights.get(p, 0), reverse=True)

    for i, person in enumerate(sorted_people):
        if i == len(sorted_people) - 1:
            # 最后一人：直接给剩余天数
            days = max(0, round(remaining_demand, 1))
        else:
            days = round(monthly_demand * weights[person], 1)
            remaining_demand -= days
        raw_allocations[person] = max(0, days)

    # 4. 过渡月份拆分
    allocations = []
    for person in sorted_people:
        total = raw_allocations[person]
        po_days = {}

        if transition_info and transition_info.get("split_old", 0) > 0:
            # 需要拆分：按权重分配旧 PO 和新 PO 的天数
            split_old_total = transition_info["split_old"]
            split_new_total = transition_info.get("split_new", 0)
            from_po = transition_info["from_po"]
            to_po = transition_info["to_po"]

            # 旧 PO 份额 = 个人权重 * 旧 PO 总量
            old_days = round(split_old_total * weights[person], 1)
            # 新 PO 份额 = 总天数 - 旧 PO 天数
            new_days = round(total - old_days, 1)

            # 确保不超分
            if old_days > total:
                old_days = total
                new_days = 0.0

            if old_days > 0:
                po_days[from_po] = old_days
            if new_days > 0:
                po_days[to_po] = new_days

            allocations.append(PersonMonthAllocation(
                person_name=person, month=month, level=level,
                total_days=total, po_days=po_days, transition=True,
                historical_rate=historical_rates.get(person, 0),
            ))
        else:
            po_days[active_po_id] = total
            allocations.append(PersonMonthAllocation(
                person_name=person, month=month, level=level,
                total_days=total, po_days=po_days, transition=False,
                historical_rate=historical_rates.get(person, 0),
            ))

    return allocations


def build_person_forecast_from_projection(
    projections: dict[str, list[ForwardProjection]],
    inventory: dict,
    personnel: dict,
    historical_rates: dict,
) -> dict:
    """
    从正向推演结果构建人员级 Forecast。

    Args:
        projections: {level: [ForwardProjection]} from run_forward_projection()
        inventory: PO inventory
        personnel: 人员配置
        historical_rates: {name: avg_days}

    Returns:
        {name: {month_str: {po_id: days}}}
    """
    result = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

    for level, proj_list in projections.items():
        for fp in proj_list:
            month = fp.month
            demand = fp.monthly_demand
            active_po = fp.active_po_id

            if active_po == "NONE":
                continue

            transition_info = None
            if fp.transition:
                transition_info = {
                    "from_po": fp.transition_from_po,
                    "to_po": fp.transition_to_po,
                    "split_old": fp.split_days_old,
                    "split_new": fp.split_days_new,
                }

            allocations = allocate_level_to_persons(
                level=level, month=month, monthly_demand=demand,
                active_po_id=active_po, personnel=personnel,
                historical_rates=historical_rates,
                transition_info=transition_info,
            )

            for alloc in allocations:
                for po_id, days in alloc.po_days.items():
                    if days > 0:
                        result[alloc.person_name][month][po_id] += days

    return {k: dict(v) for k, v in result.items()}


# ============================================================
# Phase D4: 正向推演 (Forward Projection)
# ============================================================
def run_forward_projection(
    inventory: dict,
    consumption: dict[str, LevelConsumption],
    po_configs: list[POLevelConfig],
    personnel: dict,
    start_month: str = None,
    end_month: str = None,
) -> dict[str, list[ForwardProjection]]:
    """
    正向推演: 按 Level 独立模拟 PO 消耗。

    输入:
      start_month: 推演起始月 (默认当前月)
      end_month: 推演结束月 (默认 None = 持续到所有 PO 耗尽，上限 36 个月)
    """
    if start_month is None:
        now = datetime.now()
        start_month = f"{now.year}-{now.month:02d}"

    MAX_PROJECTION_MONTHS = 36

    # 按 Level 组织 PO 配置
    level_pos = defaultdict(list)
    for c in po_configs:
        level_pos[c.level].append(c)
    for lv in level_pos:
        level_pos[lv].sort(key=lambda c: c.priority)

    projections = {}

    # Phase 1: 构建每个 Level 的 PO 队列 + 月消耗率
    level_po_queues = {}
    level_monthly_rates = {}
    projection_start = start_month

    for level, pos in level_pos.items():
        if not pos:
            continue

        lc = consumption.get(level)
        monthly_rate = 0.0
        if lc and lc.latest_monthly_rate > 0:
            monthly_rate = lc.latest_monthly_rate
        elif lc and lc.monthly_avg > 0:
            monthly_rate = lc.monthly_avg
        else:
            people = [n for n, info in personnel.items() if info.get("level") == level]
            monthly_rate = len(people) * 22 * 0.85

        level_monthly_rates[level] = monthly_rate

        po_queue = []
        for c in pos:
            if c.end_month and c.end_month < projection_start:
                continue
            inv_info = inventory.get(c.po_id, {}).get(c.level, {})
            consumed = inv_info.get("consumed", 0.0)
            remaining = max(0, c.days_quota - consumed)
            po_queue.append({
                "po_id": c.po_id,
                "quota": c.days_quota,
                "consumed": consumed,
                "remaining": remaining,
                "priority": c.priority,
                "start_month": c.start_month,
                "end_month": c.end_month,
                "allocation_mode": getattr(c, 'allocation_mode', 'usage'),
            })

        level_po_queues[level] = po_queue

    # Phase 2: 确定推演结束月
    MAX_PROJECTION_MONTHS = 36
    if end_month is not None:
        ey, em = map(int, end_month.split('-'))
        max_end = (ey, em)
    else:
        max_additional = 6  # 默认至少推演 6 个月
        for level, po_queue in level_po_queues.items():
            total_remaining = sum(p["remaining"] for p in po_queue)
            rate = level_monthly_rates.get(level, 1)
            if rate > 0:
                months = math.ceil(total_remaining / rate)
                max_additional = max(max_additional, months)
        sy, sm = map(int, start_month.split('-'))
        total_m = sm + max_additional
        em = ((total_m - 1) % 12) + 1
        ey = sy + (total_m - 1) // 12
        max_end = (ey, em)

    # Phase 3: 生成共享月份列表
    sy, sm = map(int, start_month.split('-'))
    proj_months = []
    y, m = sy, sm
    count = 0
    while (y, m) <= max_end and count < MAX_PROJECTION_MONTHS:
        proj_months.append(f"{y}-{m:02d}")
        m += 1
        if m > 12:
            y += 1
            m = 1
        count += 1

    # Phase 4: 逐 Level 推演
    for level, pos in level_pos.items():
        if not pos:
            continue
        po_queue = level_po_queues.get(level, [])
        monthly_rate = level_monthly_rates.get(level, 0)

        level_projections = []
        current_po_idx = 0

        while current_po_idx < len(po_queue) and po_queue[current_po_idx]["remaining"] <= 0:
            current_po_idx += 1

        for month_str in proj_months:
            if current_po_idx >= len(po_queue):
                level_projections.append(ForwardProjection(
                    month=month_str, active_po_id="NONE",
                    monthly_demand=0, consumed_from_po=0, po_remaining_after=0,
                    alert="ALL PO EXHAUSTED - Need new PO!"
                ))
                continue

            current_po = po_queue[current_po_idx]
            demand = monthly_rate
            remaining = current_po["remaining"]
            is_quarterly = current_po.get("allocation_mode") == "quarterly"
            end_month = current_po.get("end_month", "")

            fp = ForwardProjection(
                month=month_str,
                active_po_id=current_po["po_id"],
                monthly_demand=round(demand, 1),
                consumed_from_po=0.0,
                po_remaining_after=0.0,
            )

            # 季度模式: PO 到期强制切出，剩余天数作废
            if is_quarterly and end_month and month_str > end_month:
                # 当前 PO 已过期，强制切换到下一个
                fp.transition = True
                fp.transition_from_po = current_po["po_id"]
                fp.consumed_from_po = 0
                fp.po_remaining_after = 0
                fp.split_days_old = 0
                fp.split_days_new = round(demand, 1)
                fp.alert = f"季度 PO {current_po['po_id']} {end_month} 到期自动退出"
                current_po["remaining"] = 0
                current_po_idx += 1
                if current_po_idx < len(po_queue):
                    next_po = po_queue[current_po_idx]
                    fp.transition_to_po = next_po["po_id"]
                    next_po["remaining"] -= demand
                    fp.po_remaining_after = round(next_po["remaining"], 1)
                else:
                    fp.transition_to_po = "NONE"
                    fp.alert += " — 需新增 PO！"

            elif remaining >= demand:
                # 够用
                current_po["remaining"] -= demand
                fp.consumed_from_po = round(demand, 1)
                fp.po_remaining_after = round(current_po["remaining"], 1)
            else:
                # 不够 → 过渡
                fp.consumed_from_po = round(remaining, 1)
                fp.po_remaining_after = 0.0
                fp.transition = True
                fp.transition_from_po = current_po["po_id"]
                fp.split_days_old = round(remaining, 1)
                fp.split_days_new = round(demand - remaining, 1)

                # 计算过渡日期
                month_num = int(month_str.split('-')[1])
                wd = WORKDAYS.get(month_num, 22)
                day_fraction = remaining / demand if demand > 0 else 0
                trans_day = max(1, min(wd, int(wd * day_fraction)))
                fp.transition_date = f"{month_str}-{trans_day:02d} (约)"

                # 切换到下一个 PO
                current_po["remaining"] = 0
                current_po_idx += 1
                if current_po_idx < len(po_queue):
                    next_po = po_queue[current_po_idx]
                    fp.transition_to_po = next_po["po_id"]
                    # 新 PO 消耗剩余天数
                    next_po["remaining"] -= (demand - remaining)
                    fp.alert = (f"PO {fp.transition_from_po} 耗尽 → "
                                f"切换到 {fp.transition_to_po}")
                    fp.po_remaining_after = round(next_po["remaining"], 1)
                else:
                    fp.transition_to_po = "NONE"
                    fp.alert = "WARNING: 所有 PO 已耗尽，需要创建新 PO！"

            level_projections.append(fp)

            # 更新 monthly_rate (考虑新人加入)
            # 简单处理: 保持当前速率不变，或微调

        projections[level] = level_projections

    return projections


# ============================================================
# Phase D5: 过渡检测
# ============================================================
def detect_transitions(
    projections: dict[str, list[ForwardProjection]],
    inventory: dict,
) -> list[dict]:
    """从正向推演中提取过渡提醒"""
    alerts = []

    for level, proj_list in projections.items():
        for fp in proj_list:
            if fp.transition:
                urgency = "PLANNED"
                month_num = int(fp.month.split('-')[1])
                now_month = datetime.now().month

                if fp.month <= f"2026-{now_month:02d}":
                    urgency = "URGENT"
                elif month_num - now_month <= 1:
                    urgency = "WARNING"

                alerts.append({
                    "level": level,
                    "transition_month": fp.month,
                    "transition_date": fp.transition_date,
                    "from_po": fp.transition_from_po,
                    "to_po": fp.transition_to_po,
                    "days_in_old_po": fp.split_days_old,
                    "days_in_new_po": fp.split_days_new,
                    "urgency": urgency,
                    "action": (f"确保 {fp.transition_to_po} 在 {fp.transition_date} 前生效"
                               if fp.transition_to_po != "NONE"
                               else f"需要在 {fp.month} 前创建新 PO！"),
                })

    return alerts


# ============================================================
# Phase E: 报告生成
# ============================================================
def generate_report_v3(
    merged_allocation: dict,
    inventory: dict,
    projections: dict[str, list[ForwardProjection]],
    transitions: list[dict],
    mc_warnings: list[dict],
    personnel: dict,
    po_configs: list[POLevelConfig],
    gap_alerts: list[str],
    output_path: str,
) -> None:
    """生成 5-Sheet 输出报告"""
    import xlsxwriter

    wb = xlsxwriter.Workbook(output_path)

    # --- 通用格式 ---
    fmt_header = wb.add_format({
        'bold': True, 'bg_color': '#2F5496', 'font_color': 'white',
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_size': 11, 'font_name': 'Microsoft YaHei', 'text_wrap': True,
    })
    fmt_cell = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_size': 10, 'font_name': 'Microsoft YaHei',
    })
    fmt_num = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_size': 10, 'num_format': '#,##0.0', 'font_name': 'Microsoft YaHei',
    })
    fmt_num2 = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_size': 10, 'num_format': '#,##0.00', 'font_name': 'Microsoft YaHei',
    })
    fmt_pct = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_size': 10, 'num_format': '0.0%', 'font_name': 'Microsoft YaHei',
    })
    fmt_green = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_size': 10, 'bg_color': '#C6EFCE', 'font_name': 'Microsoft YaHei',
    })
    fmt_yellow = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_size': 10, 'bg_color': '#FFEB9C', 'font_name': 'Microsoft YaHei',
    })
    fmt_red = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_size': 10, 'bg_color': '#FFC7CE', 'font_color': '#9C0006',
        'font_name': 'Microsoft YaHei',
    })
    fmt_gray = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_size': 10, 'bg_color': '#D9D9D9', 'font_name': 'Microsoft YaHei',
    })

    # 所有月份
    all_months = sorted(set(
        m for name_data in merged_allocation.values()
        for m in name_data.keys()
    ))
    # 加上投影月份
    for level_proj in projections.values():
        for fp in level_proj:
            if fp.month not in all_months:
                all_months.append(fp.month)
    all_months.sort()

    all_po_ids = sorted(set(c.po_id for c in po_configs))
    all_names = sorted(set(
        list(merged_allocation.keys()) + list(personnel.keys())
    ))

    # ============ Sheet 1: PO Consumption Status ============
    ws1 = wb.add_worksheet("PO_Consumption_Status")
    h1 = ["PO ID", "Level", "Priority", "Start", "End",
          "Quota (d)", "Consumed (d)", "Remaining (d)",
          "Usage %", "Monthly Burn (d)", "Est. Exhaust", "Status"]
    for i, h in enumerate(h1):
        ws1.write(0, i, h, fmt_header)

    row = 1
    for po_id in sorted(inventory.keys()):
        for level in sorted(inventory[po_id].keys()):
            info = inventory[po_id][level]
            quota = info["quota"]
            consumed = info["consumed"]
            remaining = info["remaining"]
            usage = info["usage_pct"] / 100 if quota > 0 else 0

            # 找月消耗率
            burn_rate = 0
            for lv, lc_data in projections.items():
                if lv == level:
                    # 从投影中取第一个月的 demand 作为 burn rate
                    proj_list = projections.get(lv, [])
                    if proj_list:
                        burn_rate = proj_list[0].monthly_demand
                    break

            # 预估耗尽日期
            exhaust_date = ""
            if remaining > 0 and burn_rate > 0:
                months_left = remaining / burn_rate
                exhaust_date = f"~{months_left:.1f} months"

            # 状态判断
            if consumed <= 0 and quota > 0:
                status = "Future"
                status_fmt = fmt_cell
            elif remaining <= 0:
                status = "EXHAUSTED"
                status_fmt = fmt_gray
            elif usage > 0.90:
                status = "CRITICAL"
                status_fmt = fmt_red
            elif usage > 0.70:
                status = "WARNING"
                status_fmt = fmt_yellow
            elif remaining < 10:
                status = "Low"
                status_fmt = fmt_yellow
            else:
                status = "OK"
                status_fmt = fmt_green

            ws1.write(row, 0, po_id, fmt_cell)
            ws1.write(row, 1, level, fmt_cell)
            ws1.write(row, 2, info["priority"], fmt_cell)
            ws1.write(row, 3, info.get("start_month", ""), fmt_cell)
            ws1.write(row, 4, info.get("end_month", ""), fmt_cell)
            ws1.write(row, 5, quota, fmt_num)
            ws1.write(row, 6, consumed, fmt_num)
            ws1.write(row, 7, remaining, fmt_num)
            ws1.write(row, 8, usage, fmt_pct)
            ws1.write(row, 9, round(burn_rate, 1) if burn_rate else "-", fmt_num if burn_rate else fmt_cell)
            ws1.write(row, 10, exhaust_date, fmt_cell)
            ws1.write(row, 11, status, status_fmt)
            row += 1

    ws1.set_column(0, 0, 14)
    ws1.set_column(1, 1, 30)
    ws1.set_column(2, 4, 8)
    ws1.set_column(5, 7, 14)
    ws1.set_column(8, 8, 10)
    ws1.set_column(9, 11, 16)
    ws1.freeze_panes(1, 0)

    # ============ Sheet 2: Person-Month-PO Matrix ============
    ws2 = wb.add_worksheet("Person_Month_PO_Matrix")
    ws2.write(0, 0, "Name", fmt_header)
    ws2.write(0, 1, "Level", fmt_header)
    col = 2
    for month in all_months:
        ws2.write(0, col, f"{month}\nPO", fmt_header)
        ws2.write(0, col + 1, f"{month}\nDays", fmt_header)
        col += 2

    for r, name in enumerate(all_names, 1):
        ws2.write(r, 0, name, fmt_cell)
        ws2.write(r, 1, personnel.get(name, {}).get("level", ""), fmt_cell)
        col = 2
        for month in all_months:
            po_days = merged_allocation.get(name, {}).get(month, {})
            if po_days:
                po_str = ", ".join(po_days.keys())
                days_str = ", ".join(f"{v:.1f}" for v in po_days.values())
                ws2.write(r, col, po_str, fmt_cell)
                ws2.write(r, col + 1, days_str, fmt_num)
            col += 2

    ws2.set_column(0, 0, 14)
    ws2.set_column(1, 1, 28)
    ws2.freeze_panes(1, 2)

    # ============ Sheet 3: Monthly Level Summary ============
    ws3 = wb.add_worksheet("Monthly_Level_Summary")
    # 按 Level 分组
    all_levels = sorted(set(c.level for c in po_configs))

    ws3.write(0, 0, "Month", fmt_header)
    ws3.write(0, 1, "Level", fmt_header)
    col = 2
    for po_id in all_po_ids:
        ws3.write(0, col, f"{po_id[-8:]} (d)", fmt_header)
        col += 1
    ws3.write(0, col, "Total (d)", fmt_header)
    ws3.write(0, col + 1, "Workdays", fmt_header)
    ws3.write(0, col + 2, "Utilization %", fmt_header)

    row = 1
    for month in all_months:
        month_num = int(month.split('-')[1])
        for level in all_levels:
            ws3.write(row, 0, month, fmt_cell)
            ws3.write(row, 1, level, fmt_cell)
            col = 2
            total = 0.0
            for po_id in all_po_ids:
                # 汇总该月该Level在该PO的天数
                po_total = 0.0
                for name in all_names:
                    if personnel.get(name, {}).get("level") == level:
                        po_total += merged_allocation.get(name, {}).get(month, {}).get(po_id, 0)
                ws3.write(row, col, round(po_total, 1) if po_total > 0 else "-",
                         fmt_num if po_total > 0 else fmt_cell)
                total += po_total
                col += 1

            ws3.write(row, col, round(total, 1), fmt_num)
            ws3.write(row, col + 1, WORKDAYS.get(month_num, 22), fmt_cell)

            # Utilization: total / (people_count * workdays)
            people_count = len([n for n, info in personnel.items()
                               if info.get("level") == level])
            if people_count > 0 and WORKDAYS.get(month_num, 22) > 0:
                util = total / (people_count * WORKDAYS.get(month_num, 22))
            else:
                util = 0
            util_fmt = fmt_pct if util < 1 else fmt_red
            ws3.write(row, col + 2, min(util, 1.0), util_fmt)
            row += 1

    ws3.set_column(0, 1, 12)
    ws3.set_column(2, col, 12)
    ws3.freeze_panes(1, 0)

    # ============ Sheet 4: Forward Projection ============
    ws4 = wb.add_worksheet("Forward_Projection")
    h4 = ["Level", "Month", "Active PO", "Monthly Demand (d)",
          "Consumed from PO (d)", "PO Remaining (d)",
          "Transition?", "From/To PO", "Transition Date", "Alert"]
    for i, h in enumerate(h4):
        ws4.write(0, i, h, fmt_header)

    row = 1
    for level in sorted(projections.keys()):
        for fp in projections[level]:
            t_fmt = fmt_yellow if fp.transition else fmt_cell
            ws4.write(row, 0, level, fmt_cell)
            ws4.write(row, 1, fp.month, fmt_cell)
            ws4.write(row, 2, fp.active_po_id, t_fmt)
            ws4.write(row, 3, fp.monthly_demand, fmt_num)
            ws4.write(row, 4, fp.consumed_from_po, fmt_num)
            ws4.write(row, 5, fp.po_remaining_after, fmt_num)
            ws4.write(row, 6, "YES" if fp.transition else "-",
                     fmt_yellow if fp.transition else fmt_cell)
            ws4.write(row, 7,
                     f"{fp.transition_from_po} → {fp.transition_to_po}" if fp.transition else "-",
                     t_fmt)
            ws4.write(row, 8, fp.transition_date if fp.transition else "-", t_fmt)
            ws4.write(row, 9, fp.alert, fmt_red if fp.alert else fmt_cell)
            row += 1

    ws4.set_column(0, 0, 30)
    ws4.set_column(1, 1, 10)
    ws4.set_column(2, 2, 14)
    ws4.set_column(3, 5, 16)
    ws4.set_column(6, 6, 10)
    ws4.set_column(7, 7, 24)
    ws4.set_column(8, 9, 18)
    ws4.freeze_panes(1, 0)

    # ============ Sheet 5: Alerts ============
    ws5 = wb.add_worksheet("Alerts")

    # 5A: Transition Alerts
    ws5.merge_range(0, 0, 0, 6, "PO Transition Alerts", fmt_header)
    ta_headers = ["Level", "Month", "Transition Date", "From PO", "To PO",
                  "Days in Old", "Days in New", "Urgency", "Action"]
    for i, h in enumerate(ta_headers):
        ws5.write(1, i, h, fmt_header)

    row = 2
    for t in sorted(transitions, key=lambda x: x["transition_month"]):
        u_fmt = fmt_red if t["urgency"] == "URGENT" else (
            fmt_yellow if t["urgency"] == "WARNING" else fmt_cell)
        ws5.write(row, 0, t["level"], fmt_cell)
        ws5.write(row, 1, t["transition_month"], fmt_cell)
        ws5.write(row, 2, t["transition_date"], fmt_cell)
        ws5.write(row, 3, t["from_po"], fmt_cell)
        ws5.write(row, 4, t["to_po"], fmt_cell)
        ws5.write(row, 5, t["days_in_old_po"], fmt_num)
        ws5.write(row, 6, t["days_in_new_po"], fmt_num)
        ws5.write(row, 7, t["urgency"], u_fmt)
        ws5.write(row, 8, t["action"], fmt_cell)
        row += 1

    # 5B: PO Config Gap Alerts
    if gap_alerts:
        row += 1
        ws5.merge_range(row, 0, row, 6, "PO Configuration Warnings", fmt_header)
        row += 1
        for alert in gap_alerts:
            ws5.write(row, 0, alert, fmt_yellow)
            row += 1

    # 5C: MC Verification Warnings
    if mc_warnings:
        row += 1
        ws5.merge_range(row, 0, row, 6, "Monthly Closing Verification Warnings", fmt_header)
        row += 1
        mc_headers = ["Person", "Month", "Allocated Days", "MC Days", "Diff %", "Severity"]
        for i, h in enumerate(mc_headers):
            ws5.write(row, i, h, fmt_header)
        row += 1

        for w in sorted(mc_warnings, key=lambda x: -x["diff_pct"]):
            s_fmt = fmt_red if w["severity"] == "SEVERE" else fmt_yellow
            ws5.write(row, 0, w["name"], fmt_cell)
            ws5.write(row, 1, w["month"], fmt_cell)
            ws5.write(row, 2, w["merged_days"], fmt_num)
            ws5.write(row, 3, w["mc_days"], fmt_num)
            ws5.write(row, 4, w["diff_pct"] / 100, fmt_pct)
            ws5.write(row, 5, w["severity"], s_fmt)
            row += 1

    ws5.set_column(0, 8, 18)
    ws5.freeze_panes(2, 0)

    wb.close()
    print(f"\n  Report saved: {output_path}")


# ============================================================
# Phase F: Main
# ============================================================

MONTH_COL_MAP_2026 = {
    1: "W", 2: "X", 3: "Y", 4: "Z", 5: "AA", 6: "AB",
    7: "AC", 8: "AD", 9: "AE", 10: "AF", 11: "AG", 12: "AH",
}


def build_budget_report_updates(
    sheet_name: str,
    merged_allocation: dict,         # {name: {month: {po_id: days}}}
    person_forecast: dict,           # {name: {month: {po_id: days}}}
    inventory: dict,
    personnel: dict,
    fc_person_rows: dict,            # {name: forecast_row_num}
    ac_person_rows: dict,            # {name: actual_row_num}
) -> dict:
    """
    构建 Budget Report 的单元格更新映射。

    Returns:
        {sheet_name: {cell_ref: new_value}}
    """
    updates = {}

    if not fc_person_rows or not ac_person_rows:
        return updates

    # ---- Forecast 区域更新 ----
    fc_updates = {}
    for name, months in person_forecast.items():
        if name not in fc_person_rows:
            continue
        row = fc_person_rows[name]
        for month_str, po_days in months.items():
            try:
                month_num = int(month_str.split('-')[1])
            except (ValueError, IndexError):
                continue
            col_letter = MONTH_COL_MAP_2026.get(month_num)
            if col_letter:
                total_days = round(sum(po_days.values()), 1)
                cell_ref = f"{col_letter}{row}"
                fc_updates[cell_ref] = total_days if total_days > 0 else 0.0

    if fc_updates:
        updates[sheet_name] = fc_updates

    # ---- Actual 区域更新 ----
    ac_updates = {}
    for name, months in merged_allocation.items():
        if name not in ac_person_rows:
            continue
        row = ac_person_rows[name]
        for month_str, po_days in months.items():
            try:
                month_num = int(month_str.split('-')[1])
            except (ValueError, IndexError):
                continue
            col_letter = MONTH_COL_MAP_2026.get(month_num)
            if col_letter:
                total_days = round(sum(po_days.values()), 1)
                cell_ref = f"{col_letter}{row}"
                ac_updates[cell_ref] = total_days if total_days > 0 else 0.0

    if ac_updates:
        updates[sheet_name] = {**updates.get(sheet_name, {}), **ac_updates}

    return updates


def run_with_writeback(
    project: str,
    as_of: str = None,
    mode: str = "preview",
    generate_usage_status: bool = False,
    usage_status_output: str = None,
):
    """
    运行分析并通过 budget_report_writer 写入 Budget Report。
    """
    from budget_report_writer import (
        BudgetReportReader, BudgetReportWriter,
        backup_budget_report, print_diff_summary, save_diff_excel, confirm_write,
    )
    from po_usage_status_generator import generate_po_usage_status

    setup_utf8_stdout()
    print("=" * 70)
    print(f"  PO Hour Allocator V2 — {project} [MODE: {mode}]")
    print(f"  Run time: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 70)

    # 1. 加载人员配置
    personnel_map = {"CNTS-25165": config.CNTS25165_PERSONNEL}
    personnel = personnel_map.get(project, {})
    if not personnel:
        print(f"  [ERROR] No personnel config for project: {project}")
        return
    print(f"\n  Personnel: {len(personnel)} people")

    # 2. 加载 PO Level 配置
    po_configs = load_po_level_config(project)
    print(f"  PO Level Configs: {len(po_configs)} entries")

    po_configs, gap_alerts = resolve_po_start_end_gaps(po_configs)
    for c in po_configs:
        print(f"    {c.po_id} | {c.level} | {c.days_quota:.0f}d x "
              f"{c.unit_price:.0f}/h | {c.start_month}~{c.end_month} | P{c.priority}")

    # 3. 加载数据源
    print("\n--- Data Loading ---")
    gongshi_records = load_gongshi_qingdan(personnel=personnel)
    gs_people = len(set(r.person_name for r in gongshi_records))
    print(f"  Priority 1a (工时清单表): {len(gongshi_records)} records, {gs_people} people")

    waiyuan_records = load_waiyuan_qingkuang(personnel=personnel)
    wy_people = len(set(r.person_name for r in waiyuan_records))
    print(f"  Priority 1b (外援情况): {len(waiyuan_records)} records, {wy_people} people")

    from project_registry import build_profile
    project_profile = build_profile(project)
    project_config = config.PROJECTS.get(project, {})
    project_dir = project_profile.budget_report_dir
    budget_actuals, budget_forecast = load_budget_report_actuals_and_forecast(
        project_dir, personnel, project_id=project
    )

    timesheet_days = load_timesheet_days(project_dir, personnel)
    print(f"  Priority 3 (Timesheet): {len(timesheet_days)} people")

    # 4. 合并
    print("\n--- Data Merging ---")
    merged = merge_ground_truth_allocation(
        gongshi_records, waiyuan_records, budget_actuals, timesheet_days,
        personnel, po_configs, project_id=project,
    )

    # 5. MC 验证
    print("\n--- Verification ---")
    mc_hours = load_monthly_closing_for_verification(project, personnel)
    mc_warnings = verify_with_monthly_closing(merged, mc_hours, personnel)

    # 6. 库存快照
    print("\n--- PO Inventory ---")
    inventory = build_po_level_inventory(po_configs, merged, personnel)
    for po_id in sorted(inventory.keys()):
        for level in sorted(inventory[po_id].keys()):
            info = inventory[po_id][level]
            status_icon = "✓" if info["remaining"] > 0 else "✗"
            print(f"  {status_icon} {po_id} | {level}: "
                  f"{info['consumed']:.1f}/{info['quota']:.0f}d used "
                  f"({info['usage_pct']:.1f}%), {info['remaining']:.1f}d remaining")

    # 7. 消耗率
    print("\n--- Consumption Rates ---")
    consumption = calculate_level_consumption_rates(merged, personnel, inventory)
    for level, lc in consumption.items():
        print(f"  {level}: avg {lc.monthly_avg:.1f} d/mo "
              f"(recent {LOOKBACK}mo: {lc.latest_monthly_rate:.1f}), "
              f"{len(lc.active_people)} people")

    # 8. 正向推演
    print("\n--- Forward Projection ---")
    start = as_of or datetime.now().strftime("%Y-%m")
    projections = run_forward_projection(
        inventory, consumption, po_configs, personnel, start_month=start
    )
    for level, proj_list in projections.items():
        tc = sum(1 for fp in proj_list if fp.transition)
        print(f"  {level}: {len(proj_list)} months projected, {tc} transitions")

    # 9. 过渡检测
    transitions = detect_transitions(projections, inventory)
    for t in transitions:
        print(f"  [{t['urgency']}] {t['level']}: {t['transition_month']} "
              f"{t['from_po']} → {t['to_po']} "
              f"({t['days_in_old_po']:.1f}d + {t['days_in_new_po']:.1f}d)")

    # 10. 人员级分配
    print("\n--- Person-Level Allocation ---")
    historical_rates = compute_person_historical_rates(merged, personnel)
    person_forecast = build_person_forecast_from_projection(
        projections, inventory, personnel, historical_rates
    )
    fc_people = len(person_forecast)
    fc_entries = sum(
        sum(len(po_days) for po_days in months.values())
        for months in person_forecast.values()
    )
    print(f"  Person forecast: {fc_people} people, {fc_entries} person-month-PO entries")

    # 11. 查找 Budget Report 模板
    budget_path = _latest_active_budget_report(project, project_dir)
    if not budget_path:
        print(f"[ERROR] No active Budget Report found in {project_dir}; template files are not write targets")
        return

    if mode == "usage-status":
        print(f"\n  Budget Report data source (read-only): {os.path.basename(budget_path)}")
    else:
        print(f"\n  Active Budget Report write target: {os.path.basename(budget_path)}")

    def write_usage_status():
        timestamp = datetime.now().strftime("%y%m%d_%H%M%S")
        file_name = f"PO_Usage_Status_{timestamp}.xlsx"
        po_usage_dir = project_config.get("po_usage_dir", project_dir)
        output_path = usage_status_output or os.path.join(
            po_usage_dir, file_name
        )

        def generate_to(path):
            os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
            generate_po_usage_status(
                budget_path=budget_path,
                output_path=path,
                personnel=personnel,
                merged_allocation=merged,
                inventory=inventory,
                forward_projection=projections,
                person_forecast=person_forecast,
                as_of_month=start,
                template_path=project_profile.po_usage_template_path or None,
            )

        try:
            generate_to(output_path)
            return output_path, None
        except PermissionError as target_error:
            if usage_status_output:
                raise
            fallback_path = os.path.join(config.OUTPUT_DIR, file_name)
            generate_to(fallback_path)
            return fallback_path, str(target_error)

    if mode == "usage-status":
        print("\n--- PO Usage Status Generation (Budget Report unchanged) ---")
        output_path, fallback_reason = write_usage_status()
        if fallback_reason:
            print(f"  [WARN] Project folder write failed: {fallback_reason}")
            print(f"  PO Usage Status generated locally; copy it to the project folder: {output_path}")
        return

    # 12. 读取模板 → 获取人员行映射
    reader = BudgetReportReader(budget_path, personnel)
    if not reader.check_template_version():
        print("  [WARN] Template not V2 — using legacy row detection...")
        # 尝试使用旧版方法
        fc_rows = _detect_person_rows_legacy(budget_path, "FORECAST", personnel)
        ac_rows = _detect_person_rows_legacy(budget_path, "ACTUAL", personnel)
    else:
        fc_rows = reader.read_person_row_mapping()
        ac_rows = reader.read_actual_person_row_mapping()

    print(f"  Forecast person→row: {len(fc_rows)} mapped")
    print(f"  Actual person→row: {len(ac_rows)} mapped")

    # 13. 构建更新
    sheet_name, _ = reader._get_anchors()
    if not sheet_name:
        print("[ERROR] Cannot find active Budget_FC sheet")
        return

    updates = build_budget_report_updates(
        sheet_name, merged, person_forecast, inventory,
        personnel, fc_rows, ac_rows,
    )
    print(f"\n  Built {sum(len(v) for v in updates.values())} cell updates in '{sheet_name}'")

    # 14. Preview / Write
    writer = BudgetReportWriter(budget_path, project)
    diff = writer.write_updates(updates, dry_run=True)
    print_diff_summary(diff)

    if mode == "write" and diff.total_changes > 0:
        if confirm_write():
            diff = writer.write_updates(updates, dry_run=False)
            budget_path = writer.filepath
            print(f"  Budget Report updated: {budget_path}")

            # 生成 PO Usage Status
            if generate_usage_status:
                print("\n--- PO Usage Status Generation ---")
                try:
                    po_usage_path, fallback_reason = write_usage_status()
                    if fallback_reason:
                        print(f"  [WARN] Project folder write failed: {fallback_reason}")
                        print(f"  PO Usage Status generated locally; copy it to the project folder: {po_usage_path}")
                    else:
                        print(f"  PO Usage Status updated: {po_usage_path}")
                except Exception as e:
                    print(f"  [WARN] PO Usage Status generation failed: {e}")
                    print(f"  Budget Report write was successful — PO Usage Status can be regenerated later.")
        else:
            print("  Write cancelled.")
    elif diff.total_changes == 0:
        print("  No changes to apply.")

    # 同时生成独立分析报告
    print("\n--- Report Generation ---")
    timestamp = datetime.now().strftime("%y%m%d_%H%M")
    output_path = os.path.join(config.OUTPUT_DIR, f"PO_Plan_V2_{timestamp}.xlsx")
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    generate_report_v3(
        merged, inventory, projections, transitions,
        mc_warnings, personnel, po_configs, gap_alerts, output_path,
    )

    print(f"\n{'=' * 70}")
    print(f"  Done! Report: {output_path}")
    print(f"{'=' * 70}")


def _detect_person_rows_legacy(budget_path: str, section: str, personnel: dict) -> dict:
    """Legacy fallback: detect person rows without anchor markers."""
    import openpyxl
    name_lookup = config.build_name_lookup(personnel)

    wb = openpyxl.load_workbook(budget_path, data_only=True)
    for sn in wb.sheetnames:
        if 'Internal Budget' in sn and '2026' in sn:
            ws = wb[sn]
            # 搜索可能的数据区域
            mapping = {}
            section_found = False
            for row_idx in range(1, ws.max_row + 1):
                j_val = ws.cell(row=row_idx, column=10).value
                if not j_val:
                    continue
                j_str = str(j_val).strip()
                # Check if matches start of section
                if section == "FORECAST" and j_str == "Forecast":
                    section_found = True
                    continue
                if section == "ACTUAL" and j_str == "Actual":
                    section_found = True
                    continue
                if not section_found:
                    continue

                # Try to normalize
                n = j_str
                if ',' in n:
                    parts = n.split(',', 1)
                    n = f"{parts[1].strip()} {parts[0].strip()}"
                canonical = name_lookup.get(n.lower())
                if not canonical:
                    for alias, cname in name_lookup.items():
                        if alias in n.lower() or n.lower() in alias:
                            canonical = cname
                            break
                if canonical and canonical in personnel:
                    mapping[canonical] = row_idx

            wb.close()
            return mapping

    wb.close()
    return {}


def main(project: str = "CNTS-25165", as_of: str = None, mode: str = "report",
         generate_usage_status: bool = False, usage_status_output: str = None):
    """PO 工时规划器入口。

    Args:
        project: 项目代码
        as_of: 推演起点月
        mode: "report" (默认，独立报告) | "preview" (差异预览) | "write" (回写) | "usage-status" (仅生成 PO Usage Status)
        generate_usage_status: 是否同时生成 PO Usage Status
    """
    if mode in ("preview", "write", "usage-status"):
        run_with_writeback(project, as_of, mode, generate_usage_status, usage_status_output)
        return

    # === 原有 report 模式（保持不变）===
    setup_utf8_stdout()
    print("=" * 70)
    print(f"  PO Hour Allocator V2 — {project}")
    print(f"  Run time: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 70)

    # 1. 加载人员配置
    personnel_map = {"CNTS-25165": config.CNTS25165_PERSONNEL}
    personnel = personnel_map.get(project, {})
    if not personnel:
        print(f"  [ERROR] No personnel config for project: {project}")
        return
    print(f"\n  Personnel: {len(personnel)} people")

    # 2. 加载 PO Level 配置
    po_configs = load_po_level_config(project)
    print(f"  PO Level Configs: {len(po_configs)} entries")

    # 2b. 填补缺失日期
    po_configs, gap_alerts = resolve_po_start_end_gaps(po_configs)
    for c in po_configs:
        print(f"    {c.po_id} | {c.level} | {c.days_quota:.0f}d x "
              f"{c.unit_price:.0f}/h | {c.start_month}~{c.end_month} | P{c.priority}")

    # 3. 加载数据源 (按优先级)
    print("\n--- Data Loading ---")

    # P1a: 工时清单表
    gongshi_records = load_gongshi_qingdan(personnel=personnel)
    gs_people = len(set(r.person_name for r in gongshi_records))
    print(f"  Priority 1a (工时清单表): {len(gongshi_records)} records, {gs_people} people")

    # P1b: 外援情况
    waiyuan_records = load_waiyuan_qingkuang(personnel=personnel)
    wy_people = len(set(r.person_name for r in waiyuan_records))
    print(f"  Priority 1b (外援情况): {len(waiyuan_records)} records, {wy_people} people")

    # P2: Budget Report
    from project_registry import build_profile
    project_profile = build_profile(project)
    project_config = config.PROJECTS.get(project, {})
    project_dir = project_profile.budget_report_dir
    budget_actuals, budget_forecast = load_budget_report_actuals_and_forecast(
        project_dir, personnel, project_id=project
    )

    # P3: Timesheet PDFs
    timesheet_days = load_timesheet_days(project_dir, personnel)
    ts_people = len(timesheet_days)
    print(f"  Priority 3 (Timesheet): {ts_people} people")

    # 4. 合并
    print("\n--- Data Merging ---")
    merged = merge_ground_truth_allocation(
        gongshi_records, waiyuan_records,
        budget_actuals, timesheet_days,
        personnel, po_configs, project_id=project,
    )

    # 5. MC 验证
    print("\n--- Verification ---")
    project_wbs = project
    mc_hours = load_monthly_closing_for_verification(project_wbs, personnel)
    mc_warnings = verify_with_monthly_closing(merged, mc_hours, personnel)

    # 6. 库存快照
    print("\n--- PO Inventory ---")
    inventory = build_po_level_inventory(po_configs, merged, personnel)
    for po_id in sorted(inventory.keys()):
        for level in sorted(inventory[po_id].keys()):
            info = inventory[po_id][level]
            status_icon = "✓" if info["remaining"] > 0 else "✗"
            print(f"  {status_icon} {po_id} | {level}: "
                  f"{info['consumed']:.1f}/{info['quota']:.0f}d used "
                  f"({info['usage_pct']:.1f}%), {info['remaining']:.1f}d remaining")

    # 7. 消耗率
    print("\n--- Consumption Rates ---")
    consumption = calculate_level_consumption_rates(merged, personnel, inventory)
    for level, lc in consumption.items():
        print(f"  {level}: avg {lc.monthly_avg:.1f} d/mo "
              f"(recent {LOOKBACK}mo: {lc.latest_monthly_rate:.1f}), "
              f"{len(lc.active_people)} people")

    # 8. 正向推演
    print("\n--- Forward Projection ---")
    start = as_of or datetime.now().strftime("%Y-%m")
    projections = run_forward_projection(
        inventory, consumption, po_configs, personnel, start_month=start
    )
    for level, proj_list in projections.items():
        transitions_count = sum(1 for fp in proj_list if fp.transition)
        print(f"  {level}: {len(proj_list)} months projected, "
              f"{transitions_count} transitions")

    # 9. 过渡检测
    transitions = detect_transitions(projections, inventory)
    for t in transitions:
        print(f"  [{t['urgency']}] {t['level']}: {t['transition_month']} "
              f"{t['from_po']} → {t['to_po']} "
              f"({t['days_in_old_po']:.1f}d + {t['days_in_new_po']:.1f}d)")

    # 10. 生成报告
    print("\n--- Report Generation ---")
    timestamp = datetime.now().strftime("%y%m%d_%H%M")
    output_path = os.path.join(config.OUTPUT_DIR, f"PO_Plan_V2_{timestamp}.xlsx")
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)

    generate_report_v3(
        merged, inventory, projections, transitions,
        mc_warnings, personnel, po_configs, gap_alerts, output_path,
    )

    print(f"\n{'=' * 70}")
    print(f"  Done! Output: {output_path}")
    print(f"{'=' * 70}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="PO Hour Allocator V2")
    parser.add_argument("--project", default="CNTS-25165", help="Project code")
    parser.add_argument("--as-of", default=None, help="Projection start month (YYYY-MM)")
    parser.add_argument("--mode", default="report",
                       choices=["report", "preview", "write", "usage-status"],
                       help="report: standalone report (default); "
                            "preview: diff preview only; "
                            "write: backup + write to Budget Report; "
                            "usage-status: update PO Usage Status only")
    parser.add_argument("--generate-usage-status", action="store_true",
                       help="Also generate PO Usage Status after Budget Report write")
    parser.add_argument("--usage-status-output", default=None,
                       help="Optional output path for PO Usage Status")
    args = parser.parse_args()
    main(args.project, args.as_of, args.mode, args.generate_usage_status,
         args.usage_status_output)
