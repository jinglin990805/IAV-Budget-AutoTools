# -*- coding: utf-8 -*-
"""
IAV Budget Automation - Tool C: Travel Cost Summary (差旅费用汇总)

功能:
1. 扫描 03_Hose & Tengqing 全部月度差旅数据
2. 解析: 合思对账单 (xlsx)、IAV China 月度报表 (xlsx)、机票对账单 (xls)
3. 按项目人员匹配归类 -> 输出汇总 Excel
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
# 合思对账单解析 (平台消费: 机票/酒店/火车/用车)
# ============================================================

def parse_hesi(filepath, month_str):
    """
    解析合思对账单
    Sheet结构: 企业对账单, 酒店, 火车&飞机, 汇总, 交易记录, ...
    关键字段: 产品类型, 交易金额, 申请单名 (含人名)
    """
    items = []
    try:
        wb = safe_load_workbook(filepath)
        if wb is None:
            return items

        # 汇总 sheet: 总金额按类型
        has_summary = "汇总" in wb.sheetnames

        # 火车&飞机 sheet: 最全的信息 (含人名+金额)
        target_sheets = ["火车&飞机", "火车", "机票", "酒店"]
        for sn in wb.sheetnames:
            if sn not in target_sheets:
                continue

            ws = wb[sn]
            # Find header row
            headers = None
            header_row_idx = None
            for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), 1):
                vals = [safe_str(c) if c else "" for c in row]
                has_keys = sum(1 for v in vals if any(k in v for k in ["产品类型", "订单", "申请单", "金额"]))
                if has_keys >= 2:
                    headers = vals
                    header_row_idx = row_idx
                    break

            if headers is None:
                continue

            # Find columns
            type_col = next((i for i, h in enumerate(headers) if "产品类型" in h), 0)
            app_name_col = next((i for i, h in enumerate(headers)
                                 if "申请单名" in h or "申请单标题" in h), None)
            amount_col = None
            for i, h in enumerate(headers):
                if any(k in h for k in ["交易金额", "实付金额", "支付金额", "票价", "票面价"]):
                    amount_col = i
                    break
            dept_col = next((i for i, h in enumerate(headers) if "Department" in h or "部门" in h), None)

            for row in ws.iter_rows(min_row=header_row_idx + 1, max_row=ws.max_row, values_only=True):
                vals = [safe_str(c) if c else "" for c in row]
                ptype = vals[type_col] if type_col < len(vals) else ""
                app_name = vals[app_name_col] if app_name_col is not None and app_name_col < len(vals) else ""

                amount = 0
                if amount_col is not None and amount_col < len(vals):
                    try:
                        amount = abs(float(vals[amount_col]))
                    except (ValueError, TypeError):
                        pass

                dept = vals[dept_col] if dept_col is not None and dept_col < len(vals) else ""

                if amount > 0 or ptype:
                    # Extract person name from application title
                    # e.g. "郑山河（ZhengShanhe）2025-08-23Trip" -> Zheng Shanhe
                    person = ""
                    parenthetical = re.search(r'（([^）]+)）', app_name)
                    if parenthetical:
                        person = parenthetical.group(1)
                    elif "-" in app_name:
                        person = app_name.split("-")[0].strip()
                    else:
                        person = app_name[:20].strip()

                    # Map person name to project
                    proj_key = _match_person_to_project(person)
                    # Also check department
                    if not proj_key and dept:
                        proj_key = _match_dept_to_project(dept)

                    items.append({
                        "month": month_str,
                        "source": "合思",
                        "type": ptype or sn,
                        "person": person,
                        "app_name": app_name[:60],
                        "amount": round(amount, 2),
                        "project": proj_key or "",
                        "dept": dept[:30],
                    })

        wb.close()
    except Exception as e:
        print(f"  [WARN] 合思 parse error ({os.path.basename(filepath)}): {e}")

    return items


# ============================================================
# IAV China 月度报表 (服务费账单, 按成本中心)
# ============================================================

def parse_iav_china_report(filepath, month_str):
    """解析 IAV China 月度报表 (消费汇总)"""
    items = []
    try:
        sheets = get_sheet_names(filepath)
        target = None
        for sn in sheets:
            if "消费汇总" in sn or "汇总" in sn:
                target = sn
                break
        if target is None:
            target = sheets[0]

        headers, data = read_sheet(filepath, target)
        if not data:
            return items

        # Find category/amount rows: typically "类别" col + "企业支付金额" col
        cat_col, amount_col = 0, 2
        for row in data:
            if len(row) < 3:
                continue
            cat = safe_str(row[cat_col])
            amt = 0
            try:
                amt = float(row[amount_col]) if row[amount_col] else 0
            except (ValueError, TypeError):
                pass
            if amt > 0 and cat:
                items.append({
                    "month": month_str,
                    "source": "IAV China Report",
                    "type": cat,
                    "person": "",
                    "app_name": f"IAV China {cat} total",
                    "amount": round(amt, 2),
                    "project": "",
                    "dept": "IAV China",
                })
    except Exception as e:
        print(f"  [WARN] IAV China report parse error: {e}")
    return items


# ============================================================
# 机票对账单 (ticketsbuchung, xls 格式)
# ============================================================

def parse_ticketsbuchung(filepath, month_str):
    """
    解析机票对账单 (ticketsbuchung der flüge)
    这个文件包含项目号列
    """
    items = []
    try:
        # Monthly statements are normally old-style .xls, but some months were
        # saved as .xlsx. xlrd refuses those ("Excel xlsx file; not supported"),
        # so pick the reader from the real file format instead of assuming.
        with open(filepath, "rb") as fh:
            is_xlsx = fh.read(2) == b"PK"
        sheets = []
        if is_xlsx:
            from openpyxl import load_workbook
            wb = load_workbook(filepath, read_only=True, data_only=True)
            for ws in wb.worksheets:
                sheets.append([list(row) for row in ws.iter_rows(values_only=True)])
            wb.close()
        else:
            import xlrd
            xls_wb = xlrd.open_workbook(filepath)
            for sn in xls_wb.sheet_names():
                ws = xls_wb.sheet_by_name(sn)
                sheets.append([[ws.cell_value(r, c) for c in range(ws.ncols)]
                               for r in range(ws.nrows)])

        for rows in sheets:
            if len(rows) < 2:
                continue
            ncols = max(len(row) for row in rows)

            def cell(r, c):
                if r >= len(rows) or c >= len(rows[r]):
                    return ""
                value = rows[r][c]
                return "" if value is None else value

            # Find header row
            headers = [str(cell(0, c)).strip() for c in range(ncols)]
            header_row = 0

            # Columns
            name_col = next((i for i, h in enumerate(headers) if "Name" in h or "姓名" in h), 0)
            proj_col = next((i for i, h in enumerate(headers)
                            if "项目号" in h or "Project" in h or "Kst" in h), None)
            amount_col = next((i for i, h in enumerate(headers)
                              if "金额" in h or "Betrag" in h or "Amount" in h or "Sum" in h), None)
            date_col = next((i for i, h in enumerate(headers)
                            if "日期" in h or "Datum" in h or "Date" in h), None)

            for r in range(1, len(rows)):
                name = str(cell(r, name_col)).strip() if name_col < ncols else ""
                proj = str(cell(r, proj_col)).strip() if proj_col is not None and proj_col < ncols else ""
                amt = 0
                if amount_col is not None and amount_col < ncols:
                    try:
                        amt = abs(float(cell(r, amount_col)))
                    except (ValueError, TypeError):
                        pass
                if amt > 0:
                    proj_key = extract_project_key(proj) or _match_person_to_project(name) or ""
                    items.append({
                        "month": month_str,
                        "source": "机票对账单",
                        "type": "机票",
                        "person": name,
                        "app_name": f"Flight - {name}",
                        "amount": round(amt, 2),
                        "project": proj_key,
                        "dept": proj[:30],
                    })
    except ImportError:
        print(f"  [SKIP] xlrd not available, cannot read .xls: {os.path.basename(filepath)}")
    except Exception as e:
        print(f"  [WARN] Ticketsbuchung parse error ({os.path.basename(filepath)}): {e}")
    return items


# ============================================================
# 人员->项目 匹配 (基于已知的人员名单)
# ============================================================

def _match_person_to_project(person_name):
    """根据人名匹配到项目"""
    if not person_name:
        return ""
    pn = person_name.lower().replace(" ", "").replace(",", "")
    for proj_key, proj_info in PROJECTS.items():
        # Check if this person is in the project's known personnel
        if proj_key == "CNTS-25165":
            for cfg_name in CNTS25165_PERSONNEL:
                cfg_normalized = cfg_name.lower().replace(" ", "").replace(",", "")
                if cfg_normalized in pn or pn in cfg_normalized:
                    return proj_key
    return ""


def _match_dept_to_project(dept):
    """根据部门/Cost Center 匹配项目"""
    if not dept:
        return ""
    for proj_key, proj_info in PROJECTS.items():
        for cc in proj_info.get("cost_centers", []):
            if cc in dept:
                return proj_key
    return ""


# ============================================================
# 主扫描函数
# ============================================================

def scan_all_travel(months=None):
    """
    扫描 03_Hose & Tengqing 全部月份差旅数据
    Args:
        months: list of month folders, e.g. ["2601", "2602"]. None = scan all
    Returns: list of travel items
    """
    if months is None:
        months = sorted([d for d in os.listdir(TRAVEL_DIR)
                        if os.path.isdir(os.path.join(TRAVEL_DIR, d)) and d.isdigit()])

    all_items = []
    for month_str in months:
        folder = os.path.join(TRAVEL_DIR, month_str)
        if not os.path.exists(folder):
            continue

        files = os.listdir(folder)
        if not files:
            print(f"  {month_str}: empty folder, skipping")
            continue

        print(f"  {month_str}: {len(files)} files")

        for fname in files:
            if fname.startswith("~$"):
                continue
            fpath = os.path.join(folder, fname)
            ext = os.path.splitext(fname)[1].lower()

            # 合思对账单 (always .xlsx)
            if "合思" in fname or "对账" in fname:
                if ext in (".xlsx",):
                    items = parse_hesi(fpath, month_str)
                    all_items.extend(items)
                    if items:
                        print(f"    合思: {len(items)} items")

            # IAV China 月度报表
            elif "IAV" in fname and "China" in fname and ext in (".xlsx",):
                items = parse_iav_china_report(fpath, month_str)
                all_items.extend(items)
                if items:
                    print(f"    IAV China Report: {len(items)} items")

            # 机票对账单 (ticketsbuchung)
            elif "ticketsbuchung" in fname.lower() or "机票对账单" in fname:
                if ext in (".xls", ".xlsx"):
                    items = parse_ticketsbuchung(fpath, month_str)
                    all_items.extend(items)
                    if items:
                        print(f"    机票对账单: {len(items)} items")

    return all_items


# ============================================================
# 输出汇总
# ============================================================

def generate_travel_summary(items, output_path):
    setup_utf8_stdout()

    if not items:
        print("[WARN] No travel items found!")
        # Still create empty report
        items = []

    # Sheet 1: 明细
    d_headers = ["Month", "Source", "Type", "Person", "Project", "Amount", "Description"]
    d_rows = []
    for it in items:
        d_rows.append([it["month"], it["source"], it["type"], it["person"],
                       it["project"], it["amount"], it["app_name"]])

    # Sheet 2: 按项目汇总
    proj_totals = defaultdict(lambda: defaultdict(float))
    for it in items:
        key = it["project"] or "Unclassified"
        proj_totals[key][it["type"]] += it["amount"]

    p_headers = ["Project", "Type", "Total Amount"]
    p_rows = []
    for proj in sorted(proj_totals.keys()):
        for ptype, amt in sorted(proj_totals[proj].items()):
            p_rows.append([proj, ptype, round(amt, 2)])

    # Sheet 3: 按月汇总
    month_totals = defaultdict(lambda: defaultdict(float))
    for it in items:
        month_totals[it["month"]][it["project"] or "Unclassified"] += it["amount"]

    m_headers = ["Month", "Project", "Amount"]
    m_rows = []
    for month in sorted(month_totals.keys()):
        for proj, amt in sorted(month_totals[month].items()):
            m_rows.append([month, proj, round(amt, 2)])

    # Sheet 4: 按人员汇总
    person_totals = defaultdict(float)
    person_projs = defaultdict(set)
    for it in items:
        if it["person"]:
            person_totals[it["person"]] += it["amount"]
            if it["project"]:
                person_projs[it["person"]].add(it["project"])

    pe_headers = ["Person", "Project(s)", "Total Amount"]
    pe_rows = []
    for name in sorted(person_totals.keys(), key=lambda n: -person_totals[n]):
        pe_rows.append([name, ", ".join(person_projs.get(name, set())),
                        round(person_totals[name], 2)])

    sheet_data = OrderedDict([
        ("Detail", (d_headers, d_rows)),
        ("By_Project", (p_headers, p_rows)),
        ("By_Month", (m_headers, m_rows)),
        ("By_Person", (pe_headers, pe_rows)),
    ])

    write_excel(output_path, sheet_data)

    # Print summary
    print(f"\n  Travel Cost Summary:")
    print(f"  {'Project':25s} | {'Amount (RMB)':>12s}")
    print(f"  {'-'*25}|{'-'*13}")
    for proj in sorted(proj_totals.keys()):
        total = sum(proj_totals[proj].values())
        print(f"  {proj:25s} | {total:>12,.0f}")


# ============================================================
# Main
# ============================================================

def main():
    setup_utf8_stdout()
    print("=" * 60)
    print("  Travel Cost Summary - All Projects")
    print("=" * 60)

    print("\n[1/2] Scanning travel data folders...")
    items = scan_all_travel()  # Scan all months
    print(f"\n  Total: {len(items)} travel cost items found")

    print("\n[2/2] Generating summary report...")
    ts = datetime.now().strftime("%y%m%d_%H%M")
    rpt = os.path.join(OUTPUT_DIR, f"Travel_Summary_{ts}.xlsx")
    generate_travel_summary(items, rpt)

    print(f"\n{'='*60}")
    print(f"  Done! Output: output/Travel_Summary_{ts}.xlsx")
    print("=" * 60)


if __name__ == "__main__":
    main()
