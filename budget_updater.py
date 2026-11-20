# -*- coding: utf-8 -*-
"""
IAV Budget Automation - Tool B: Budget Report Monthly Updater

功能:
1. 从 Monthly Closing 解析各项目工时（按项目号+成本中心汇总）
2. 从 03_Hose & Tengqing 汇总月度差旅费用
3. 生成综合 Budget Summary Excel（工时+差旅+PO状态）
"""

import sys
import io
import os
import re
import glob
import math
import sqlite3
import warnings
from datetime import datetime
from collections import defaultdict

import config
from config import *
from utils import *

# ============================================================
# Module B1: 工时汇总
# ============================================================

def _monthly_closing_header_row(ws):
    """Find the row containing Monthly Closing column labels."""
    for row_idx in range(1, min(ws.max_row, 6) + 1):
        headers = [safe_str(value).strip().lower() for value in next(
            ws.iter_rows(min_row=row_idx, max_row=row_idx, values_only=True)
        )]
        if any("repdate" in value for value in headers):
            return row_idx, headers
    return None, []


def _monthly_closing_columns(headers, month_name):
    """Return (total quantity column, project number column, month hours column)."""
    total_col = next((i for i, value in enumerate(headers)
                      if "total quantity" in value), 9)
    project_col = next((
        i for i, value in enumerate(headers)
        if ("project number" in value or "project no" in value)
        and "wbs" not in value
    ), None)
    month_aliases = {
        "Jan": {"jan", "january"}, "Feb": {"feb", "february"},
        "Mar": {"mar", "march"}, "Apr": {"apr", "april"},
        "May": {"may"}, "June": {"jun", "june"},
        "July": {"jul", "july"}, "Aug": {"aug", "august"},
        "Sep": {"sep", "sept", "september"}, "Oct": {"oct", "october"},
        "Nov": {"nov", "november"}, "Dec": {"dec", "december"},
    }
    month_col = None
    prior_col = None
    aliases = month_aliases.get(month_name, set())
    for i, value in enumerate(headers):
        if any(term in value for term in ("until end", "last month", "accumulated", "prior month")):
            prior_col = i
            continue
        tokens = set(re.findall(r"[a-z]+", value))
        if value in {"this month", "current month"} or tokens.intersection(aliases):
            month_col = i
            break
    return total_col, project_col, month_col, prior_col


def _read_monthly_closing_data():
    """Read project totals and per-person hours by month from the latest Monthly Closing."""
    latest_mc = find_latest_file(TMS_DIR, "Monthly Closing_Data_Base_*.xlsx")
    if not latest_mc:
        print("[ERROR] No Monthly Closing file found!")
        return {}, {}

    print(f"[INFO] Reading: {os.path.basename(latest_mc)}")
    month_names = {
        "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "June": 6,
        "July": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
    }
    # openpyxl cannot preserve this source workbook's extended validation
    # metadata. We only read Monthly Closing and never save it, so suppress the
    # resulting warning that otherwise looks like a failed write-back.
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message="Data Validation extension is not supported and will be removed",
            category=UserWarning,
        )
        wb = safe_load_workbook(latest_mc)
    if wb is None:
        return {}, {}

    totals = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    people = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    for sheet_name, month_num in month_names.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header_row, headers = _monthly_closing_header_row(ws)
        if header_row is None:
            continue
        total_col, project_col, month_col, prior_col = _monthly_closing_columns(headers, sheet_name)
        if project_col is None:
            project_col = 10 if sheet_name == "Jan" else 12
        # January has no separate current-month column. Other sheets expose it
        # explicitly; if not, use Total Quantity minus the previous cumulative column.
        if month_col is None and sheet_name == "Jan":
            month_col = total_col

        for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=True):
            if len(row) <= max(total_col, project_col):
                continue
            proj_num = safe_str(row[project_col]).strip() if row[project_col] else ""
            if not proj_num:
                continue
            rd = safe_str(row[2]) if len(row) > 2 else ""
            year_match = re.search(r"20\d{2}", rd)
            year = int(year_match.group(0)) if year_match else 2026

            qty = row[month_col] if month_col is not None and month_col < len(row) else None
            if not isinstance(qty, (int, float)) and prior_col is not None:
                total_qty = row[total_col] if total_col < len(row) else None
                prior_qty = row[prior_col] if prior_col < len(row) else None
                if isinstance(total_qty, (int, float)) and isinstance(prior_qty, (int, float)):
                    qty = total_qty - prior_qty
            if not isinstance(qty, (int, float)) or qty == 0:
                continue

            proj_key = extract_project_key(proj_num) or proj_num
            month_str = f"{year}-{month_num:02d}"
            cc_parts = proj_num.rsplit("-", 1)
            cost_center = cc_parts[1] if len(cc_parts) > 1 else "00000"
            hours = round(float(qty), 2)
            totals[proj_key][cost_center][month_str] += hours

            last = safe_str(row[0]) if row[0] else ""
            first = safe_str(row[1]) if row[1] else ""
            person_name = f"{first} {last}".strip()
            if person_name:
                people[proj_key][person_name][month_str] += hours

    wb.close()
    return dict(totals), dict(people)


def aggregate_hours_by_project():
    """Return Monthly Closing hours by project, cost center, and calendar month."""
    totals, _ = _read_monthly_closing_data()
    return totals


def _normalize_target_period(target_month):
    if target_month in (None, ""):
        return None
    value = str(target_month).strip()
    if re.fullmatch(r"\d{6}", value):
        year, month = value[:4], int(value[4:6])
        return f"{year}-{month:02d}" if 1 <= month <= 12 else None
    if re.fullmatch(r"\d{4}-\d{2}", value):
        year, month = value.split("-", 1)
        return value if 1 <= int(month) <= 12 else None
    if re.fullmatch(r"\d{4}", value):
        month = int(value[2:])
        return f"20{value[:2]}-{month:02d}" if 1 <= month <= 12 else None
    return None


def _read_approved_timesheet_hours(project):
    """Read current, approved PDF facts without modifying the fact database."""
    from fact_store import load_approved_timesheet_facts

    facts = load_approved_timesheet_facts(project)
    result = defaultdict(lambda: defaultdict(float))
    total_hours = 0.0
    for fact in facts:
        person = str(fact.get("canonical_name") or fact.get("person_key") or "").strip()
        month = str(fact.get("month") or "").strip()
        try:
            hours = float(fact.get("hours") or 0)
        except (TypeError, ValueError):
            continue
        if person and month and math.isfinite(hours) and hours > 0:
            result[person][month] += hours
            total_hours += hours
    print(
        f"  [Review Center] Loaded {len(facts)} current approved Timesheet facts "
        f"({total_hours:.1f}h) from {config.FACT_STORE_PATH}"
        if facts else
        f"  [Review Center] No current approved Timesheet facts for {project}; "
        f"Monthly Closing fallback will be used ({config.FACT_STORE_PATH})"
    )
    return dict(result)


# ============================================================
# Module B2: 差旅汇总
# ============================================================

def aggregate_travel_costs(target_month=None):
    """
    从 03_Hose & Tengqing 汇总差旅费用
    Args:
        target_month: "YYYYMM" (e.g. "202608") or travel-folder "YYMM" (e.g. "2608"). None = latest
    Returns:
        list of {project, category, description, amount, month}
    """
    if target_month is None:
        # Find latest month folder
        subdirs = sorted([d for d in os.listdir(TRAVEL_DIR)
                         if os.path.isdir(os.path.join(TRAVEL_DIR, d)) and d.isdigit()],
                        reverse=True)
        target_month = subdirs[0] if subdirs else None

    if not target_month:
        print("[WARNING] No travel data folders found")
        return []

    # The workbench accepts YYYYMM (for example, 202608), while travel
    # folders use YYMM (2608).
    target_month = str(target_month).strip()
    if len(target_month) == 6 and target_month.isdigit():
        target_month = target_month[2:]

    travel_path = os.path.join(TRAVEL_DIR, target_month)
    if not os.path.exists(travel_path):
        print(f"[WARNING] Travel folder not found: {travel_path}")
        return []

    print(f"[INFO] Scanning travel data: {target_month}")
    items = []

    for fname in os.listdir(travel_path):
        fpath = os.path.join(travel_path, fname)
        if os.path.isdir(fpath):
            continue

        # Skip non-data files
        if fname.startswith("~$"):
            continue

        ext = os.path.splitext(fname)[1].lower()

        # Process Excel files (IAV China reports, 合思 reconciliation, ticket booking)
        if ext in (".xlsx", ".xls"):
            items.extend(_parse_travel_excel(fpath, target_month))

        # Process PDF payment requests
        elif ext == ".pdf":
            # Extract basic info from filename
            # e.g. "Payment+Request(Supplier)-B25001435.pdf"
            if "Payment" in fname or "Trip" in fname:
                items.append({
                    "project": "",
                    "category": "Business Travel",
                    "description": fname[:80],
                    "amount": 0,  # Need to read PDF content
                    "month": target_month,
                })

    return items


def _parse_travel_excel(filepath, month_str):
    """Parse a travel-related Excel file"""
    items = []
    try:
        wb = safe_load_workbook(filepath)
        if wb is None:
            return items

        for sn in wb.sheetnames[:3]:  # Usually 1-2 sheets
            headers, data = read_sheet(filepath, sn)
            if not data:
                continue

            # Try to identify columns
            proj_col = None
            amount_col = None
            desc_col = None
            cat_col = None

            for i, h in enumerate(headers):
                h_l = h.lower()
                if any(k in h_l for k in ["项目", "project", "wbs", "成本中心"]):
                    proj_col = i
                if any(k in h_l for k in ["金额", "amount", "借方本币", "cost", "sum"]):
                    amount_col = i
                if any(k in h_l for k in ["摘要", "description", "用途", "科目名称"]):
                    desc_col = i
                if any(k in h_l for k in ["科目", "category", "类型"]):
                    cat_col = i

            for row in data:
                proj = safe_str(row[proj_col]) if proj_col is not None and proj_col < len(row) else ""
                desc = safe_str(row[desc_col]) if desc_col is not None and desc_col < len(row) else ""
                cat = safe_str(row[cat_col]) if cat_col is not None and cat_col < len(row) else "Travel"

                amount = 0
                if amount_col is not None and amount_col < len(row):
                    try:
                        amount = float(row[amount_col]) if row[amount_col] else 0
                    except (ValueError, TypeError):
                        pass

                # If no specific amount column found, try to find numeric values
                if amount == 0 and amount_col is None:
                    for v in row:
                        try:
                            amount = float(v) if v and float(v) > 100 else amount
                        except (ValueError, TypeError):
                            pass

                if amount > 0:
                    # Extract project code from description or project column
                    proj_key = extract_project_key(proj) or extract_project_key(desc) or ""
                    items.append({
                        "project": proj_key,
                        "category": cat,
                        "description": desc[:100],
                        "amount": round(amount, 2),
                        "month": month_str,
                    })

        wb.close()
    except Exception as e:
        print(f"  [WARNING] Could not parse {os.path.basename(filepath)}: {e}")

    return items


# ============================================================
# Module B3: PO Status
# ============================================================

def get_po_status():
    """
    从 PO_Record.xlsx 读取各项目 PO 状态
    Returns: list of PO status dicts
    """
    if not os.path.exists(PO_RECORD_FILE):
        print("[WARNING] PO Record file not found")
        return []

    headers, data = read_sheet(PO_RECORD_FILE, "POtable")
    result = []
    for row in data:
        if len(row) < 4:
            continue
        top_wbs = safe_str(row[0])
        po_num = safe_str(row[1])
        value = row[2] if isinstance(row[2], (int, float)) else 0
        date_recv = row[3] if isinstance(row[3], datetime) else safe_str(row[3])
        note = safe_str(row[4]) if len(row) > 4 else ""

        proj_key = extract_project_key(top_wbs) or top_wbs
        proj_name = WBS_TO_PROJECT.get(proj_key, proj_key)

        result.append({
            "Project": proj_key,
            "Project Name": proj_name,
            "WBS": top_wbs,
            "PO Number": po_num,
            "Value": value,
            "Date Received": date_recv,
            "Note": note,
        })

    return result


# ============================================================
# Module B4: 综合输出
# ============================================================

def generate_budget_summary(output_path, target_month=None):
    """
    生成月度 Budget Summary Excel
    """
    setup_utf8_stdout()
    print("=" * 60)
    print("  IAV Budget Report Monthly Updater")
    print("=" * 60)

    # B1: Hours aggregation
    print("\n[B1] Aggregating hours by project...")
    hours_data = aggregate_hours_by_project()
    print(f"  Found {len(hours_data)} projects with hours")

    # B2: Travel aggregation
    print("\n[B2] Aggregating travel costs...")
    travel_items = aggregate_travel_costs(target_month)
    print(f"  Found {len(travel_items)} travel cost items")

    # B3: PO Status
    print("\n[B3] Reading PO status...")
    po_status = get_po_status()
    print(f"  Found {len(po_status)} POs")

    # ---- Build Sheet 1: Hours Summary ----
    h_headers = ["Project", "Project Name", "Cost Center", "Month", "Hours (h)", "Days"]
    h_rows = []
    for proj_key, ccs in _normalize_hours_data(hours_data):
        proj_name = WBS_TO_PROJECT.get(proj_key, proj_key)
        for cc, month, hours_val in ccs:
            days = round(hours_val / HOURS_PER_DAY, 2)
            h_rows.append([proj_key, proj_name, cc, month, round(hours_val, 2), days])

    # ---- Build Sheet 2: Travel Summary ----
    t_headers = ["Project", "Category", "Description", "Amount (RMB)", "Month"]
    t_rows = [[i["project"], i["category"], i["description"], i["amount"], i["month"]]
              for i in travel_items]

    # ---- Build Sheet 3: PO Status ----
    if po_status:
        p_headers = list(po_status[0].keys())
        p_rows = [list(p.values()) for p in po_status]
    else:
        p_headers = ["Status"]
        p_rows = [["No PO data found"]]

    # ---- Build Sheet 4: Project Overview ----
    o_headers, o_rows = _build_overview(hours_data, travel_items, po_status)

    sheet_data = {
        "Hours_Summary": (h_headers, h_rows),
        "Travel_Summary": (t_headers, t_rows),
        "PO_Status": (p_headers, p_rows),
        "Project_Overview": (o_headers, o_rows),
    }

    write_excel(output_path, sheet_data)
    print(f"\n{'='*60}")
    print(f"  Budget Summary generated!")
    print(f"  {'='*60}")
    for row in o_rows:
        print(f"  {row[1]:20s} | {row[2]:8.1f}h | {row[3]:6.1f}d | {row[4]:8.0f} RMB | {row[5]} POs")

    return hours_data, travel_items, po_status


def _normalize_hours_data(hours_data: dict) -> list:
    """将 {proj: {cc: {month: hours}}} 展平"""
    result = []
    for proj_key in sorted(hours_data.keys()):
        ccs = hours_data[proj_key]
        items = []
        for cc, months in sorted(ccs.items()):
            for month, hours_val in sorted(months.items()):
                items.append((cc, month, hours_val))
        result.append((proj_key, items))
    return result


def _build_overview(hours_data: dict, travel_items: list, po_status: list):
    """构建项目概览"""
    o_headers = ["Project", "Project Name", "Total Hours (h)", "Total Days",
                 "Total Travel (RMB)", "PO Count", "Active POs"]
    o_rows = []
    for proj_key in sorted(set(list(hours_data.keys()) +
                               [t["project"] for t in travel_items if t["project"]])):
        proj_name = WBS_TO_PROJECT.get(proj_key, proj_key)
        total_h = sum(sum(months.values()) for months in hours_data.get(proj_key, {}).values())
        total_d = round(total_h / HOURS_PER_DAY, 2)
        total_t = sum(t["amount"] for t in travel_items if t["project"] == proj_key)
        pos = [p for p in po_status if p["Project"] == proj_key]
        po_count = len(pos)
        active_pos = sum(1 for p in pos if isinstance(p.get("Date Received"), datetime)
                        and p["Date Received"] > datetime(2026, 1, 1))
        o_rows.append([proj_key, proj_name, round(total_h, 2), total_d,
                       round(total_t, 2), po_count, active_pos])
    return o_headers, o_rows


# ============================================================
# Module B5: Budget Report Write-Back (月度 Actual + Travel)
# ============================================================

def write_actuals_to_budget_report(project: str, target_month: str = None):
    """
    将 MC 工时 + 差旅数据写入 Budget Report。
    三层安全: 备份 → 预览 → 确认
    """
    import openpyxl
    from openpyxl.utils import column_index_from_string, get_column_letter

    setup_utf8_stdout()
    print("=" * 60)
    print(f"  Budget Monthly Write-Back — {project}")
    print("=" * 60)

    target_period = _normalize_target_period(target_month)
    if target_month not in (None, "") and not target_period:
        print(f"  [ERROR] Invalid target month: {target_month!r}; expected YYYYMM or YYMM")
        return

    # 1. 加载数据
    print("\n[Step 1] Loading data...")
    # 工时
    hours_data, mc_person_hours = _read_monthly_closing_data()
    proj_hours = hours_data.get(project, {})
    mc_by_person = defaultdict(lambda: defaultdict(float))
    approved_by_person = defaultdict(lambda: defaultdict(float))
    # Actual 工时按人员/月取数：已批准 Timesheet 是三方人员的正式工时；
    # 对没有已批准工时表的人员，再使用 Monthly Closing 当月列补齐。
    for raw_name, month_values in mc_person_hours.get(project, {}).items():
        for month_str, hours in month_values.items():
            mc_by_person[raw_name][month_str] += hours
    approved_timesheets = _read_approved_timesheet_hours(project)
    for raw_name, month_values in approved_timesheets.items():
        for month_str, hours in month_values.items():
            approved_by_person[raw_name][month_str] += hours

    if not proj_hours:
        print(f"  [WARN] No hours found for {project}")

    if target_period:
        mc_period_hours = sum(
            month_values.get(target_period, 0.0)
            for month_values in proj_hours.values()
        )
        timesheet_period_hours = sum(
            month_values.get(target_period, 0.0)
            for month_values in approved_by_person.values()
        )
        print(f"  Source month: {target_period}")
        print(f"  Monthly Closing: {mc_period_hours:.1f}h; approved Timesheets: {timesheet_period_hours:.1f}h")

    # 差旅
    travel_items = aggregate_travel_costs(target_month)
    proj_travel = [t for t in travel_items if t["project"] == project]
    if not target_period:
        total_hours = sum(sum(m.values()) for m in proj_hours.values())
        approved_total = sum(sum(m.values()) for m in approved_by_person.values())
        print(f"  Hours: {total_hours:.1f}h across {len(proj_hours)} cost centers")
        print(f"  Approved Timesheets: {approved_total:.1f}h")
    print(f"  Travel: {len(proj_travel)} items, ¥{sum(t['amount'] for t in proj_travel):,.0f}")

    # 2. 定位 Budget Report
    print("\n[Step 2] Locating Budget Report...")
    from project_registry import build_profile, latest_budget_report
    profile = build_profile(project)
    project_dir = profile.budget_report_dir
    if not project_dir or not os.path.isdir(project_dir):
        print(f"  [ERROR] Budget Report dir not found for {project}")
        return

    active_report = latest_budget_report(profile)
    if active_report is None:
        print(f"  [ERROR] No Budget Report found in {project_dir}")
        return
    budget_path = str(active_report)
    print(f"  Found: {os.path.basename(budget_path)}")

    # 3. 读取 Budget Report — 定位 Actual 区域
    print("\n[Step 3] Reading Budget Report...")
    from budget_report_writer import (
        BudgetReportReader, BudgetReportWriter,
        backup_budget_report, print_diff_summary, confirm_write,
    )

    reader = BudgetReportReader(budget_path, config.CNTS25165_PERSONNEL if project == "CNTS-25165" else {})
    sheet_name, anchors = reader._get_anchors()
    if not sheet_name:
        print("  [ERROR] Cannot find active Budget_FC sheet")
        return

    # 获取 Actual 区域人员行映射
    if anchors:
        ac_rows = reader.read_actual_person_row_mapping(sheet_name)
    else:
        ac_rows = _detect_person_rows_legacy(budget_path, "ACTUAL")

    row_by_person = {}
    for display_name, row_num in ac_rows.items():
        canonical = reader._normalize_name(display_name)
        if not canonical and project != "CNTS-25165":
            canonical = display_name.strip()
        if not canonical:
            continue
        row_by_person[canonical] = row_num
    print(f"  Actual person→row: {len(row_by_person)} mapped")

    # 4. 构建 Actual 区域更新
    print("\n[Step 4] Building Actual updates...")
    # 月份列映射 (与 po_allocator_v2 一致)
    MONTH_COL_MAP = {
        1: "W", 2: "X", 3: "Y", 4: "Z", 5: "AA", 6: "AB",
        7: "AC", 8: "AD", 9: "AE", 10: "AF", 11: "AG", 12: "AH",
    }

    updates = {}
    # 统一姓名别名，避免预算表、Monthly Closing、Timesheet 使用不同姓名顺序时漏写。
    mc_hours_by_canonical = defaultdict(lambda: defaultdict(float))
    for raw_name, month_values in mc_by_person.items():
        canonical = reader._normalize_name(raw_name)
        if not canonical and project != "CNTS-25165":
            canonical = raw_name.strip()
        if canonical:
            for month_str, hours in month_values.items():
                mc_hours_by_canonical[canonical][month_str] += hours

    timesheet_hours_by_canonical = defaultdict(lambda: defaultdict(float))
    for raw_name, month_values in approved_by_person.items():
        canonical = reader._normalize_name(raw_name)
        if not canonical and project != "CNTS-25165":
            canonical = raw_name.strip()
        if canonical:
            for month_str, hours in month_values.items():
                timesheet_hours_by_canonical[canonical][month_str] += hours

    unmapped_timesheet_people = sorted(
        person for person, month_values in timesheet_hours_by_canonical.items()
        if (not target_period or target_period in month_values)
        if person not in row_by_person
    )
    if unmapped_timesheet_people:
        print(f"  [WARN] Approved Timesheets have no Budget Actual row: {', '.join(unmapped_timesheet_people)}")

    all_periods = set()
    for person_months in mc_hours_by_canonical.values():
        all_periods.update(person_months)
    for person_months in timesheet_hours_by_canonical.values():
        all_periods.update(person_months)

    # Actual 工时更新：若该人该月有已批准工时表，以工时表为准；否则取 MC。
    for canonical, row_num in row_by_person.items():
        periods = {target_period} if target_period else all_periods
        for month_str in periods:
            if not month_str:
                continue
            if month_str in timesheet_hours_by_canonical.get(canonical, {}):
                hours_val = timesheet_hours_by_canonical[canonical][month_str]
            elif month_str in mc_hours_by_canonical.get(canonical, {}):
                hours_val = mc_hours_by_canonical[canonical][month_str]
            else:
                continue
            try:
                month_num = int(month_str.split('-')[1])
            except (ValueError, IndexError):
                continue
            col = MONTH_COL_MAP.get(month_num)
            if col:
                days = round(hours_val / HOURS_PER_DAY, 2)
                ref = f"{col}{row_num}"
                updates[ref] = days

    # 5. Travel Cost sheet 更新
    travel_updates = {}
    if proj_travel:
        # 定位 Travel Cost sheet
        wb = safe_load_workbook(budget_path)
        travel_sheet_name = None
        for sn in wb.sheetnames:
            if 'Travel Cost' in sn.lower().replace(' ', ''):
                travel_sheet_name = sn
                break
        if not travel_sheet_name:
            # 搜索含 "Travel" 的 sheet
            for sn in wb.sheetnames:
                if 'travel' in sn.lower():
                    travel_sheet_name = sn
                    break
        wb.close()

        if travel_sheet_name:
            print(f"  Travel Cost sheet: {travel_sheet_name}")
            # 按类别汇总
            cat_totals = defaultdict(float)
            for t in proj_travel:
                cat_totals[t.get("category", "Travel")] += t.get("amount", 0)
            # 此处需要根据实际模板确定写入位置，暂时输出摘要
            print(f"  Travel categories: {dict(cat_totals)}")
            # 注意: Travel Cost sheet 的具体写入逻辑待查看模板后细化
            # 当前先写入 Budget Report 的 Actual 工时，Travel 记录在日志中
        else:
            print("  [WARN] No Travel Cost sheet found in Budget Report")

    # 合并更新
    ac_updates = {sheet_name: updates} if updates else {}
    if travel_updates:
        ac_updates.update(travel_updates)

    if not ac_updates:
        print("  No updates to apply.")
        return

    # 6. 预览 & 写入
    writer = BudgetReportWriter(budget_path, project)
    diff = writer.write_updates(ac_updates, dry_run=True)
    print_diff_summary(diff)

    if diff.total_changes == 0:
        print("  No changes to apply.")
        return

    if confirm_write():
        diff = writer.write_updates(ac_updates, dry_run=False)
        print(f"  Monthly Actuals updated in: {os.path.basename(writer.filepath)}")
    else:
        print("  Write cancelled.")


def _detect_person_rows_legacy(budget_path, section):
    """Legacy person row detection for non-V2 templates."""
    import openpyxl as _xl
    wb = _xl.load_workbook(budget_path, data_only=True)
    mapping = {}
    try:
        for sn in wb.sheetnames:
            if 'Internal Budget' in sn and '2026' in sn:
                ws = wb[sn]
                section_found = False
                for row_idx in range(1, ws.max_row + 1):
                    j_val = ws.cell(row=row_idx, column=10).value
                    if not j_val:
                        continue
                    j_str = str(j_val).strip()
                    if section == "FORECAST" and j_str == "Forecast":
                        section_found = True
                        continue
                    if section == "ACTUAL" and j_str == "Actual":
                        section_found = True
                        continue
                    if not section_found:
                        continue
                    n = j_str
                    if ',' in n:
                        parts = n.split(',', 1)
                        n = f"{parts[1].strip()} {parts[0].strip()}"
                    mapping[n] = row_idx
    finally:
        wb.close()
    return mapping


# ============================================================
# 部门看板数据聚合
# ============================================================

def get_dept_dashboard_data(dept: str = None, months: list = None):
    """
    按部门聚合 Budget 数据用于看板展示。
    从 Mapping.xlsx 动态读取项目→Leader→部门映射。

    Returns:
        {dept, projects: [{code, name, hours, days, travel, po_count, po_active}]}
    """
    import openpyxl

    # 1. 从 Mapping.xlsx 读取项目→Leader→部门映射
    project_dept_map = {}
    mapping_path = MAPPING_FILE
    if os.path.exists(mapping_path):
        try:
            wb = openpyxl.load_workbook(mapping_path, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
                if len(row) < 6:
                    continue
                leader = str(row[1]).strip() if row[1] else ""
                proj_def = str(row[2]).strip() if row[2] else ""
                if not proj_def or not leader:
                    continue
                dept_code = PROJECT_LEADER_TO_DEPT.get(leader, "TV")
                if dept:
                    if dept_code == dept:
                        project_dept_map[proj_def] = dept_code
                else:
                    project_dept_map[proj_def] = dept_code
            wb.close()
        except Exception as e:
            print(f"  [WARN] Could not read Mapping.xlsx: {e}")

    if dept:
        project_dept_map = {k: v for k, v in project_dept_map.items() if v == dept}

    # 2. 聚合工时
    hours_data = aggregate_hours_by_project()

    # 3. 聚合差旅
    travel_items = aggregate_travel_costs()

    # 4. PO 状态
    po_status = get_po_status()

    # 5. 构建看板数据
    projects = []
    for proj_key in sorted(project_dept_map.keys()):
        proj_name = WBS_TO_PROJECT.get(proj_key, proj_key)
        total_h = sum(sum(months.values()) for months in hours_data.get(proj_key, {}).values())
        total_d = round(total_h / HOURS_PER_DAY, 2)
        total_t = sum(t["amount"] for t in travel_items if t["project"] == proj_key)
        pos = [p for p in po_status if p["Project"] == proj_key]
        po_count = len(pos)
        active_pos = sum(1 for p in pos if isinstance(p.get("Date Received"), datetime)
                        and p["Date Received"] > datetime(2026, 1, 1))
        projects.append({
            "code": proj_key,
            "name": proj_name,
            "hours": round(total_h, 1),
            "days": total_d,
            "travel": round(total_t, 2),
            "po_count": po_count,
            "po_active": active_pos,
        })

    return {
        "dept": dept or "All",
        "projects": projects,
    }


# ============================================================
# Main
# ============================================================

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Budget Report Monthly Updater")
    parser.add_argument("--project", default=None, help="Project code filter")
    parser.add_argument("--month", default=None, help="Target month (YYYYMM)")
    parser.add_argument("--mode", default="report",
                       choices=["report", "write"],
                       help="report: standalone summary (default); write: write to Budget Report")
    parser.add_argument("--dept", default=None, help="Department filter for dashboard data")
    args = parser.parse_args()

    setup_utf8_stdout()

    if args.mode == "write":
        if not args.project:
            print("[ERROR] --project is required for write mode")
            sys.exit(1)
        write_actuals_to_budget_report(args.project, args.month)
        return

    if args.dept:
        # Dashboard JSON mode
        data = get_dept_dashboard_data(args.dept, None)
        d_projects = len(data.get("projects", []))
        print(f"Department: {data['dept']}, Projects: {d_projects}")
        for p in data["projects"]:
            print(f"  {p['code']} {p['name']}: {p['hours']}h, ¥{p['travel']}, {p['po_active']}/{p['po_count']} POs")
        return

    # Default: report mode
    ts = datetime.now().strftime("%y%m%d_%H%M")
    output = os.path.join(OUTPUT_DIR, f"Budget_Summary_{ts}.xlsx")
    generate_budget_summary(output, args.month)


if __name__ == "__main__":
    main()
