# -*- coding: utf-8 -*-
"""
IAV Budget Automation - Tool A: PO Hour Allocator (PO工时智能分配器)

功能:
1. 双数据源自动合并: Monthly Closing (IAV员工) + Budget Report (分包商人员)
2. 按 PO 优先级自动分配每人每月工时到每张 PO
3. 过渡月份拆分计算
4. 添加新 PO 后自动重算 Use Untill 日期
5. 输出: 每人每月每天数（独立列）+ PO消耗进度 + Use Untill
"""

import sys
import io
import os
import re
import glob
from datetime import datetime
from collections import defaultdict, OrderedDict

from config import *
from utils import *

# ============================================================
# 月份工具
# ============================================================

def parse_month(month_str):
    y, m = str(month_str).replace("_", "-").split("-")
    return (int(y), int(m))

def format_month(y, m):
    return f"{y}-{m:02d}"

def month_range(start, end):
    sy, sm = parse_month(start)
    ey, em = parse_month(end)
    result = []
    y, m = sy, sm
    while (y, m) <= (ey, em):
        result.append(format_month(y, m))
        m += 1
        if m > 12:
            y += 1
            m = 1
    return result


# ============================================================
# 双数据源: Monthly Closing + Budget Report
# ============================================================

def load_all_hours(project_wbs="CNTS-25165", personnel=None):
    """
    合并双数据源: Monthly Closing (IAV内部) + Budget Report (含分包商)
    """
    all_hours = defaultdict(lambda: defaultdict(float))
    sources_found = []

    # Source 1: Monthly Closing
    latest_mc = find_latest_file(TMS_DIR, "Monthly Closing_Data_Base_*.xlsx")
    if latest_mc:
        mc_hours = _read_monthly_closing(latest_mc, project_wbs)
        if mc_hours:
            print(f"[INFO] Source 1 (Monthly Closing): {len(mc_hours)} people")
            sources_found.append("Monthly Closing")
            for name, months in mc_hours.items():
                for m, d in months.items():
                    all_hours[name][m] += d

    # Source 2: Budget Report (for subcontractor hours)
    budget_dir = PROJECTS.get(project_wbs, {}).get("budget_report_dir", "")
    if budget_dir and os.path.exists(budget_dir):
        budget_files = glob.glob(os.path.join(budget_dir, "IAV_Budget_Report_*.xlsx"))
        if budget_files:
            latest_br = max(budget_files, key=os.path.getmtime)
            br_hours = _read_budget_report_hours(latest_br, personnel)
            if br_hours:
                print(f"[INFO] Source 2 (Budget Report): {len(br_hours)} people")
                sources_found.append("Budget Report")
                for name, months in br_hours.items():
                    for m, d in months.items():
                        # Only add Budget Report data if NOT already in Monthly Closing
                        # (Monthly Closing is more accurate for IAV employees)
                        if name not in all_hours or not all_hours.get(name):
                            all_hours[name][m] = d

    if not sources_found:
        print("[ERROR] No hours data found from any source!")
        return {}

    print(f"[OK] Combined {len(all_hours)} people from: {', '.join(sources_found)}")
    return dict(all_hours)


def _read_monthly_closing(filepath, project_wbs):
    """读 Monthly Closing 工时"""
    month_names = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "June": 6,
                   "July": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}
    wb = safe_load_workbook(filepath)
    if wb is None:
        return {}
    person_hours = defaultdict(lambda: defaultdict(float))
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
            proj = safe_str(row[12]) if len(row) > 12 else ""
            qty = row[9] if isinstance(row[9], (int, float)) else 0
            if project_wbs in proj and qty > 0:
                rd = safe_str(row[2]) if len(row) > 2 else ""
                yr = int(rd.split("_")[0]) if "_" in rd else 2026
                days = round(qty / HOURS_PER_DAY, 5)
                person_hours[name][f"{yr}-{mn:02d}"] += days
    wb.close()
    return dict(person_hours)


def _read_budget_report_hours(filepath, personnel):
    """
    从 Budget Report 的 Internal Budget_FC sheet 读取实际工时
    特别用于获取分包商人员工时（这些不在 Monthly Closing 中）

    从实际数据观察:
    - 2026 Q2 sheet: 月份列从 col M (12) 开始: Jan=col M... Dec=col X
      但 Budget Report 有 gaps（如 start from Mar），需要从 row 9/20 等判断
    - "Actual" 区域: 人员名在第 J 或 K 列，每月天数从 L 列向右延伸
    """
    if personnel is None:
        personnel = {}

    wb = safe_load_workbook(filepath)
    if wb is None:
        return {}
    person_hours = defaultdict(lambda: defaultdict(float))

    for sn in wb.sheetnames:
        if "Budget_FC" not in sn and "Internal Budget_FC" not in sn:
            continue
        if "2027" in sn:
            continue
        yr = 2026
        if "2025" in sn:
            yr = 2025
        elif "2024" in sn:
            yr = 2024

        ws = wb[sn]
        # Find month headers row - look for cells with short month names
        month_map = {}  # {col_idx: (year, month_num)}
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
            vals = [safe_str(c) if c else "" for c in row]
            month_count = 0
            for ci, v in enumerate(vals):
                if v and len(v) == 3 and v[:3] in ["Jan","Feb","Mar","Apr","May","Jun",
                                                      "Jul","Aug","Sep","Oct","Nov","Dec"]:
                    month_num = ["Jan","Feb","Mar","Apr","May","Jun",
                                "Jul","Aug","Sep","Oct","Nov","Dec"].index(v[:3]) + 1
                    month_map[ci] = (yr, month_num)
                    month_count += 1
            if month_count >= 3:
                break

        if not month_map:
            continue

        # Now scan all rows for person data in the "Actual" section
        in_actual = False
        for row_idx, row in enumerate(ws.iter_rows(min_row=10, max_row=ws.max_row, values_only=True), 10):
            vals = [safe_str(c) if c else "" for c in row]

            # Detect "Actual" section header
            if vals[10] == "Actual" or "Actual" in safe_str(vals[10]):
                in_actual = True
                continue
            if vals[10] == "Forecast" or vals[10] == "":
                # End of Actual section
                if in_actual:
                    break
                continue

            if not in_actual:
                continue

            # Person name in col K (10)
            pname = vals[10] if len(vals) > 10 else ""
            if not pname or pname in ["Forecast", "Actual", ""]:
                continue

            # Read monthly days from mapped columns
            for ci, (cyr, cmn) in month_map.items():
                if ci >= len(vals):
                    continue
                v = vals[ci]
                if v and v.replace(".", "").replace("-", "").replace(" ", "").isdigit():
                    try:
                        d = float(v)
                    except ValueError:
                        continue
                    if d == 0:
                        continue
                    person_hours[pname][f"{cyr}-{cmn:02d}"] += d

    wb.close()
    return dict(person_hours)


# ============================================================
# PO 配置读取 (支持 input/po_config.xlsx 覆盖)
# ============================================================

def load_po_config():
    """优先从 input/po_config.xlsx 读取，否则用 config.py 默认值"""
    input_cfg = os.path.join(INPUT_DIR, "po_config.xlsx")
    if os.path.exists(input_cfg):
        print(f"[INFO] Using PO config file: po_config.xlsx")
        headers, data = read_sheet(input_cfg, "PO_Config")
        if data:
            po_list = []
            for row in data:
                if len(row) < 6:
                    continue
                po_id = safe_str(row[0]).strip()
                if not po_id:  # Skip empty rows
                    continue
                po_list.append({
                    "po_id": safe_str(row[0]),
                    "value": float(row[1]) if row[1] else 0,
                    "people": safe_str(row[2]),
                    "start_month": safe_str(row[3]),
                    "end_month": safe_str(row[4]),
                    "priority": int(float(row[5])) if row[5] else 99,
                })
            if po_list:
                print(f"  Loaded {len(po_list)} POs from config file")
                return po_list

    print(f"[INFO] Using default PO config (5 POs)")
    po_list = []
    for po_id, value, people, start, end, priority in CNTS25165_PO_CONFIG:
        po_list.append({
            "po_id": po_id,
            "value": float(value),
            "people": people if isinstance(people, str) else ", ".join(people),
            "start_month": start,
            "end_month": end,
            "priority": priority,
        })
    return po_list


def create_po_config_template():
    """生成 PO 配置模板供用户填写"""
    output_path = os.path.join(INPUT_DIR, "po_config.xlsx")
    if os.path.exists(output_path):
        print(f"[INFO] po_config.xlsx already exists at {output_path}")
        return output_path

    headers = ["PO_ID", "Value_RMB", "People", "Start_Month", "End_Month", "Priority"]
    rows = []
    for po_id, value, people, start, end, priority in CNTS25165_PO_CONFIG:
        rows.append([
            po_id,
            value,
            people if isinstance(people, str) else ", ".join(people),
            start,
            end,
            priority,
        ])
    from utils import write_excel
    write_excel(output_path, {"PO_Config": (headers, rows)})
    print(f"[OK] PO config template created: {output_path}")
    print(f"     To add a new PO: add a row in this file, save, re-run")
    print(f"     Priority 1 = first to use, higher numbers = later POs")
    return output_path


# ============================================================
# 核心分配算法
# ============================================================

def calculate_po_allocation(personnel, po_config, actual_hours, vctc_rates):
    """
    按 PO 优先级分配每人的每月工时
    Returns:
        allocation: {po_id: {name: {month: days}}}  # 每PO每人每月天数
        po_summary: 每张 PO 的汇总
        use_untill: 每人当前 PO 预计耗尽日期
        transition_detail: 过渡月份拆分详情
    """
    sorted_pos = sorted(po_config, key=lambda p: p["priority"])
    allocation = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    po_remaining = {}
    person_po_idx = defaultdict(int)

    # Initialize PO data
    for po in sorted_pos:
        po_people = (list(personnel.keys())
                     if po["people"].lower() == "all"
                     else [p.strip() for p in po["people"].split(",")])
        total_dr = sum(vctc_rates.get(personnel.get(p, {}).get("level", ""), 0)
                      for p in po_people)
        po["_people"] = po_people
        po["_daily_rate_sum"] = total_dr
        po["_max_days"] = round(po["value"] / total_dr, 2) if total_dr > 0 else 0
        po_remaining[po["po_id"]] = po["value"]

    # Collect all months
    all_months = set()
    for months in actual_hours.values():
        all_months.update(months.keys())
    sorted_months = sorted(all_months)

    # Track transitions
    transition_detail = []

    # Process each person-month
    for month in sorted_months:
        for name in personnel:
            if name not in actual_hours or month not in actual_hours[name]:
                continue
            days_todo = actual_hours[name][month]

            while days_todo > 0.001 and person_po_idx[name] < len(sorted_pos):
                po = sorted_pos[person_po_idx[name]]
                po_id = po["po_id"]

                if name not in po["_people"]:
                    person_po_idx[name] += 1
                    continue
                if month < po["start_month"]:
                    person_po_idx[name] += 1
                    continue

                prate = vctc_rates.get(personnel.get(name, {}).get("level", ""), 0)
                if prate == 0:
                    person_po_idx[name] += 1
                    continue

                prem_days = round(po_remaining[po_id] / prate, 5)

                if prem_days >= days_todo:
                    allocation[po_id][name][month] += days_todo
                    cost = round(days_todo * prate, 2)
                    po_remaining[po_id] = round(po_remaining[po_id] - cost, 2)
                    days_todo = 0
                else:
                    if prem_days > 0:
                        allocation[po_id][name][month] += prem_days
                        cost = round(prem_days * prate, 2)
                        po_remaining[po_id] = round(po_remaining[po_id] - cost, 2)

                        # Record transition
                        next_po_id = sorted_pos[person_po_idx[name] + 1]["po_id"] \
                                     if person_po_idx[name] + 1 < len(sorted_pos) else "NONE"
                        transition_detail.append({
                            "Person": name,
                            "Month": month,
                            "Old PO": po_id,
                            "New PO": next_po_id,
                            "Days in Old PO": round(prem_days, 2),
                            "Days in New PO": round(days_todo - prem_days, 2),
                            "Total Days": round(actual_hours[name][month], 2),
                        })
                        days_todo = round(days_todo - prem_days, 5)

                    person_po_idx[name] += 1

    # Calculate Use Untill per PO (not per person — PO is shared)
    # Logic: PO remaining value / total daily rate of ALL people on this PO
    use_untill = {}
    latest_month = sorted_months[-1] if sorted_months else "2026-01"

    # First, compute per-PO Use Untill
    po_use_untill = {}
    for po in sorted_pos:
        po_id = po["po_id"]
        rem_val = po_remaining[po_id]
        if rem_val <= 0:
            continue
        # Total daily burn rate for this PO = sum of daily rates of ALL people on it
        total_daily_burn = po.get("_daily_rate_sum", 0)
        if total_daily_burn <= 0:
            continue
        po_remaining_days = round(rem_val / total_daily_burn, 2)

        # How many calendar months based on ~21 workdays/month
        months_left = po_remaining_days / 21
        y, m = parse_month(latest_month)
        extra_months = max(1, int(months_left) + (1 if (months_left % 1) > 0.2 else 0))
        for _ in range(extra_months):
            m += 1
            if m > 12:
                y += 1
                m = 1
        po_use_untill[po_id] = {
            "remaining_value": rem_val,
            "remaining_team_days": po_remaining_days,
            "estimated_until": format_month(y, m),
        }

    # Then assign to each person based on their active PO
    for name in personnel:
        idx = person_po_idx.get(name, 0)
        if idx < len(sorted_pos):
            po = sorted_pos[idx]
            po_id = po["po_id"]
            if name in po["_people"] and po_id in po_use_untill:
                info = po_use_untill[po_id]
                prate = vctc_rates.get(personnel.get(name, {}).get("level", ""), 0)
                per_person_days = round(po_remaining[po_id] / prate, 2) if prate > 0 else 0
                use_untill[name] = {
                    "current_po": po_id,
                    "remaining_team_days": info["remaining_team_days"],
                    "remaining_person_days": per_person_days,
                    "estimated_until": info["estimated_until"],
                    "note": f"Shared with {len(po['_people'])} people"
                }

    # PO Summary
    po_summary = []
    for po in sorted_pos:
        rem = po_remaining[po["po_id"]]
        used = round(po["value"] - rem, 2)
        pct = round(used / po["value"] * 100, 1) if po["value"] > 0 else 0
        po_summary.append({
            "PO ID": po["po_id"],
            "Value (RMB)": po["value"],
            "Used (RMB)": used,
            "Remaining (RMB)": rem,
            "Usage %": pct,
            "Max Days": po["_max_days"],
            "People": po["people"],
            "Start": po["start_month"],
            "End": po["end_month"],
            "Priority": po["priority"],
        })

    return dict(allocation), po_summary, use_untill, transition_detail


# ============================================================
# 输出生成 (每月独立显示，不累计)
# ============================================================

def generate_report(allocation, po_summary, use_untill, personnel, actual_hours,
                    transition_detail, output_path):
    setup_utf8_stdout()

    # Collect all months
    all_m = set()
    for months in actual_hours.values():
        all_m.update(months.keys())
    sorted_m = sorted(all_m)

    # ======== Sheet 1: 每人每月每PO天数矩阵 (每月独立列) ========
    # Format: Person | Level | PO_ID | Jan | Feb | Mar | ... | Total
    mh = ["Person", "Level", "PO ID"] + sorted_m + ["Total"]
    mr = []

    for po_id in sorted(allocation.keys()):
        for name in sorted(allocation[po_id].keys()):
            months_data = allocation[po_id][name]
            row = [name, personnel[name]["level"], po_id]
            total = 0
            for m in sorted_m:
                d = round(months_data.get(m, 0), 2)
                row.append(d if d > 0 else "")
                total += d
            row.append(round(total, 2))
            mr.append(row)

    # Also add people with no data (future/not started yet)
    added_people = set()
    for row in mr:
        added_people.add((row[0], row[2]))

    # ======== Sheet 2: PO Summary ========
    sh = list(po_summary[0].keys())
    sr = [list(p.values()) for p in po_summary]

    # ======== Sheet 3: Use Untill (per person, with remaining days) ========
    uh = ["Person", "Level", "Current PO", "Team Remaining Days",
          "Person Remaining Days", "Use Untill (est.)", "Note"]
    ur = []
    for name in sorted(personnel.keys()):
        info = use_untill.get(name, {})
        ur.append([
            name,
            personnel[name]["level"],
            info.get("current_po", "N/A"),
            info.get("remaining_team_days", "N/A"),
            info.get("remaining_person_days", "N/A"),
            info.get("estimated_until", "N/A"),
            info.get("note", ""),
        ])

    # ======== Sheet 4: Transition Detail (过渡月份拆分) ========
    th = ["Person", "Month", "Old PO", "New PO", "Days in Old PO",
          "Days in New PO", "Total Days"]
    tr = []
    for t in transition_detail:
        tr.append([t["Person"], t["Month"], t["Old PO"], t["New PO"],
                   t["Days in Old PO"], t["Days in New PO"], t["Total Days"]])

    # ======== Sheet 5: Monthly Overview (all people, all months) ========
    oh = ["Person", "Level"] + sorted_m + ["Total"]
    o_r = []
    for name in sorted(personnel.keys()):
        row = [name, personnel[name]["level"]]
        total = 0
        for m in sorted_m:
            d = round(actual_hours.get(name, {}).get(m, 0), 2)
            row.append(d if d > 0 else "")
            total += d
        row.append(round(total, 2))
        o_r.append(row)

    sheet_data = OrderedDict([
        ("1_PO_Days_Per_Month", (mh, mr)),
        ("2_PO_Summary", (sh, sr)),
        ("3_Use_Untill", (uh, ur)),
        ("4_Transition_Detail", (th, tr)),
        ("5_Monthly_Overview", (oh, o_r)),
    ])

    write_excel(output_path, sheet_data)
    print(f"\n[OK] Report: {output_path}")


# ============================================================
# 名称模糊匹配
# ============================================================

def match_personnel(actual_hours, personnel):
    """将实际工时中的人名匹配到 personnel dict 的 key

    匹配策略:
    1. 精确匹配优先（如 "Hou Jing" == "Hou Jing"）
    2. 模糊匹配: 姓和名的集合完全一致（忽略顺序），
       例如 "Hou Jing" <-> "Jing, Hou"（集合都是 {hou, jing}）
    3. 每个原始数据源名称只能匹配一次（防止一个数据匹配给多人）
    """
    matched = {}
    available_hours = dict(actual_hours)  # 可匹配池，匹配后即移除

    for cfg_name in personnel:
        found = False

        # Step 1: Exact match
        if cfg_name in available_hours:
            matched[cfg_name] = available_hours.pop(cfg_name)
            found = True
        else:
            # Step 2: Fuzzy match - 要求名字部分集合完全一致
            cfg_parts = set(p.strip().lower() for p in cfg_name.replace(",", " ").split() if p.strip())

            best_match = None
            for ah_name, months in available_hours.items():
                ah_parts = set(p.strip().lower() for p in ah_name.replace(",", " ").split() if p.strip())

                # 必须两个集合完全相等（忽略顺序），才能匹配
                if cfg_parts == ah_parts:
                    best_match = ah_name
                    break

            if best_match:
                matched[cfg_name] = available_hours.pop(best_match)
                found = True

        if not found:
            matched[cfg_name] = {}

    found_count = sum(1 for v in matched.values() if v)
    print(f"  Matched: {found_count}/{len(personnel)} people")
    for name in sorted(personnel.keys()):
        if not matched.get(name):
            print(f"    [NO DATA] {name} (subcontractor - will use 0 until Budget Report data available)")
    return matched


# ============================================================
# Main
# ============================================================

def main():
    setup_utf8_stdout()
    print("=" * 60)
    print("  PO Hour Allocator - CNTS-25165 VCTC AI")
    print("=" * 60)

    # Ensure PO config template exists
    create_po_config_template()

    # 1. Load rates
    print("\n[1/5] Loading rates...")
    rates = VCTC_RATES

    # 2. Load actual hours (dual source)
    print("\n[2/5] Loading actual hours (Monthly Closing + Budget Report)...")
    actual_hours_raw = load_all_hours("CNTS-25165", CNTS25165_PERSONNEL)
    actual_hours = match_personnel(actual_hours_raw, CNTS25165_PERSONNEL)

    # 3. Load PO config
    print("\n[3/5] Loading PO configuration...")
    po_config = load_po_config()
    print(f"  {len(po_config)} POs loaded")

    # 4. Calculate allocation
    print("\n[4/5] Calculating allocation...")
    allocation, po_summary, use_untill, transitions = calculate_po_allocation(
        CNTS25165_PERSONNEL, po_config, actual_hours, rates
    )

    # Print summary
    print(f"\n  {'PO ID':>15s} | {'Value':>10s} | {'Used':>10s} | {'Remaining':>10s} | {'%':>5s}")
    print(f"  {'-'*15}|{'-'*11}|{'-'*11}|{'-'*11}|{'-'*6}")
    for po in po_summary:
        print(f"  {po['PO ID']:>15s} | {po['Value (RMB)']:>10,.0f} | "
              f"{po['Used (RMB)']:>10,.0f} | {po['Remaining (RMB)']:>10,.0f} | "
              f"{po['Usage %']:>4.0f}%")

    print(f"\n  --- Use Untill (estimated PO exhaustion date) ---")
    for name in sorted(use_untill.keys()):
        info = use_untill[name]
        print(f"  {name:20s} | PO: {info['current_po']:15s} | "
              f"Team remaining: {info['remaining_team_days']:6.1f}d | Until: {info['estimated_until']}"
              f" {info.get('note', '')}")

    if transitions:
        print(f"\n  --- Transition Detail (PO split months) ---")
        for t in transitions[:10]:
            print(f"  {t['Person']:20s} | {t['Month']} | Old: {t['Old PO']} ({t['Days in Old PO']}d)"
                  f" | New: {t['New PO']} ({t['Days in New PO']}d)")

    # Show allocation per month per person
    print(f"\n  --- Per-Month Allocation (recent 4 months) ---")
    all_m = sorted(set(m for months in actual_hours.values() for m in months))
    for name in sorted(CNTS25165_PERSONNEL.keys()):
        totals = defaultdict(float)
        for po_id, persons in allocation.items():
            if name in persons:
                for m, d in persons[name].items():
                    totals[m] += d
        recent = {m: d for m, d in sorted(totals.items()) if m in all_m[-4:]}
        if recent:
            detail = " | ".join(f"{m}:{d:.1f}d" for m, d in sorted(recent.items()))
            print(f"  {name:20s} | {detail}")

    # 5. Generate output
    print("\n[5/5] Generating report...")
    ts = datetime.now().strftime("%y%m%d_%H%M")
    rpt = os.path.join(OUTPUT_DIR, f"PO_Plan_{ts}.xlsx")
    generate_report(allocation, po_summary, use_untill, CNTS25165_PERSONNEL,
                    actual_hours, transitions, rpt)

    print("\n" + "=" * 60)
    print(f"  Done! Output: output/PO_Plan_{ts}.xlsx")
    print(f"  To add new POs, edit: input/po_config.xlsx")
    print("=" * 60)


if __name__ == "__main__":
    main()
