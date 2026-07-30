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
from datetime import datetime
from collections import defaultdict

from config import *
from utils import *

# ============================================================
# Module B1: 工时汇总
# ============================================================

def aggregate_hours_by_project():
    """
    从 Monthly Closing 读取全部项目工时，按项目号汇总
    Returns: dict {project_key: {cost_center: {month: total_hours}}}
    """
    latest_mc = find_latest_file(TMS_DIR, "Monthly Closing_Data_Base_*.xlsx")
    if not latest_mc:
        print("[ERROR] No Monthly Closing file found!")
        return {}

    print(f"[INFO] Reading: {os.path.basename(latest_mc)}")
    month_names = {
        "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "June": 6,
        "July": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
    }

    wb = safe_load_workbook(latest_mc)
    if wb is None:
        return {}

    # Structure: {project_key: {cost_center: {month: hours}}}
    result = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    # Also track per-person for detail
    person_detail = defaultdict(list)

    for sn in wb.sheetnames:
        if sn not in month_names:
            continue
        mn = month_names[sn]
        ws = wb[sn]

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            if len(row) < 13:
                continue
            last = safe_str(row[0]) if row[0] else ""
            first = safe_str(row[1]) if row[1] else ""
            name = f"{last}, {first}".strip(", ")
            proj_num = safe_str(row[12]) if len(row) > 12 else ""
            qty = row[9] if isinstance(row[9], (int, float)) else 0
            if not proj_num or qty <= 0:
                continue

            rd = safe_str(row[2]) if len(row) > 2 else ""
            yr = int(rd.split("_")[0]) if "_" in rd else 2026
            month_str = f"{yr}-{mn:02d}"

            # Extract project key and cost center
            proj_key = extract_project_key(proj_num) or proj_num
            # Cost center is typically the last segment: CN03-22052-10000
            parts = proj_num.rsplit("-", 1)
            cc = parts[1] if len(parts) > 1 else "00000"

            hours = round(qty, 2)
            result[proj_key][cc][month_str] += hours
            person_detail[proj_key].append({
                "name": name,
                "cost_center": cc,
                "month": month_str,
                "hours": hours,
            })

    wb.close()
    return dict(result)


# ============================================================
# Module B2: 差旅汇总
# ============================================================

def aggregate_travel_costs(target_month=None):
    """
    从 03_Hose & Tengqing 汇总差旅费用
    Args:
        target_month: "YYYYMM" format, e.g. "2605". None = latest
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
    for proj_key in sorted(hours_data.keys()):
        proj_name = WBS_TO_PROJECT.get(proj_key, proj_key)
        ccs = hours_data[proj_key]
        for cc, months in sorted(ccs.items()):
            for month, hours_val in sorted(months.items()):
                days = round(hours_val / HOURS_PER_DAY, 2)
                h_rows.append([proj_key, proj_name, cc, month, round(hours_val, 2), days])

    # ---- Build Sheet 2: Travel Summary ----
    t_headers = ["Project", "Category", "Description", "Amount (RMB)", "Month"]
    t_rows = []
    for item in travel_items:
        t_rows.append([
            item["project"],
            item["category"],
            item["description"],
            item["amount"],
            item["month"],
        ])

    # ---- Build Sheet 3: PO Status ----
    if po_status:
        p_headers = list(po_status[0].keys())
        p_rows = [list(p.values()) for p in po_status]
    else:
        p_headers = ["Status"]
        p_rows = [["No PO data found"]]

    # ---- Build Sheet 4: Project Overview ----
    # Summarize per project: total hours, total travel, PO count
    o_headers = ["Project", "Project Name", "Total Hours (h)", "Total Days",
                 "Total Travel (RMB)", "PO Count", "Active POs"]
    o_rows = []
    for proj_key in sorted(set(list(hours_data.keys()) +
                               [t["project"] for t in travel_items if t["project"]])):
        proj_name = WBS_TO_PROJECT.get(proj_key, proj_key)

        # Total hours
        total_h = sum(sum(months.values()) for months in hours_data.get(proj_key, {}).values())
        total_d = round(total_h / HOURS_PER_DAY, 2)

        # Total travel
        total_t = sum(t["amount"] for t in travel_items if t["project"] == proj_key)

        # PO count
        pos = [p for p in po_status if p["Project"] == proj_key]
        po_count = len(pos)
        # Active = POs with recent dates
        active_pos = sum(1 for p in pos if isinstance(p.get("Date Received"), datetime)
                        and p["Date Received"] > datetime(2026, 1, 1))

        o_rows.append([proj_key, proj_name, round(total_h, 2), total_d,
                       round(total_t, 2), po_count, active_pos])

    # ---- Write output ----
    sheet_data = {
        "Hours_Summary": (h_headers, h_rows),
        "Travel_Summary": (t_headers, t_rows),
        "PO_Status": (p_headers, p_rows),
        "Project_Overview": (o_headers, o_rows),
    }

    write_excel(output_path, sheet_data)

    # Print summary
    print(f"\n{'='*60}")
    print(f"  Budget Summary generated!")
    print(f"  {'='*60}")
    for row in o_rows:
        print(f"  {row[1]:20s} | {row[2]:8.1f}h | {row[3]:6.1f}d | {row[4]:8.0f} RMB | {row[5]} POs")


# ============================================================
# Main
# ============================================================

def main():
    setup_utf8_stdout()
    ts = datetime.now().strftime("%y%m%d_%H%M")
    output = os.path.join(OUTPUT_DIR, f"Budget_Summary_{ts}.xlsx")
    generate_budget_summary(output)


if __name__ == "__main__":
    main()
