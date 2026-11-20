# -*- coding: utf-8 -*-
"""
IAV Budget Automation Tools — 共用工具函数
"""

import os
import io
import sys
import glob
import openpyxl
import xlsxwriter
from datetime import datetime


# ============================================================
# 编码处理
# ============================================================

def setup_utf8_stdout():
    """确保控制台支持 UTF-8 输出"""
    if hasattr(sys.stdout, 'buffer') and sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        except Exception:
            pass  # 静默失败，允许在一些环境下跳过


def safe_str(value):
    """安全地将值转换为字符串（处理 None、数字、中文等）"""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value)
    return str(value)


# ============================================================
# Excel 读取
# ============================================================

def safe_load_workbook(path, data_only=True):
    """
    安全打开 Excel 文件
    Returns: openpyxl Workbook or None
    """
    if not os.path.exists(path):
        print(f"[WARNING] File not found: {path}")
        return None
    try:
        return openpyxl.load_workbook(path, data_only=data_only)
    except Exception as e:
        print(f"[ERROR] Failed to open {path}: {e}")
        return None


def get_sheet_names(path):
    """获取 Excel 文件的所有 sheet 名称"""
    wb = safe_load_workbook(path)
    if wb is None:
        return []
    names = wb.sheetnames
    wb.close()
    return names


def read_sheet(path, sheet_name=None):
    """
    读取一个 sheet 的全部数据
    Returns: (headers list, rows list of lists)
    """
    wb = safe_load_workbook(path)
    if wb is None:
        return [], []

    if sheet_name is None:
        sheet_name = wb.sheetnames[0]

    if sheet_name not in wb.sheetnames:
        print(f"[WARNING] Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
        wb.close()
        return [], []

    ws = wb[sheet_name]
    rows = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
        rows.append([safe_str(c) for c in row])

    wb.close()
    headers = rows[0] if rows else []
    data = rows[1:] if len(rows) > 1 else []
    return headers, data


def find_row_range(ws, start_row, col, match_value):
    """在指定列中查找匹配值的行号范围"""
    result_rows = []
    for row in range(start_row, ws.max_row + 1):
        cell_val = safe_str(ws.cell(row=row, column=col).value)
        if match_value.lower() in cell_val.lower():
            result_rows.append(row)
    return result_rows


# ============================================================
# Excel 写入
# ============================================================

def write_excel(output_path, sheet_data, formats=None):
    """
    写入格式化 Excel 文件
    Args:
        output_path: 输出文件路径
        sheet_data: dict {sheet_name: [header_list, rows_list_of_lists]}
        formats: dict (reserved for future formatting)
    """
    wb = xlsxwriter.Workbook(output_path)

    # 定义常用格式
    fmt_header = wb.add_format({
        'bold': True,
        'bg_color': '#4472C4',
        'font_color': 'white',
        'border': 1,
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter',
    })
    fmt_cell = wb.add_format({
        'border': 1,
        'text_wrap': True,
        'valign': 'vcenter',
    })
    fmt_number = wb.add_format({
        'border': 1,
        'num_format': '#,##0.00',
        'valign': 'vcenter',
    })
    fmt_date = wb.add_format({
        'border': 1,
        'num_format': 'yyyy-mm-dd',
        'valign': 'vcenter',
    })
    fmt_warn = wb.add_format({
        'border': 1,
        'bg_color': '#FFF2CC',
        'valign': 'vcenter',
    })

    for sheet_name, (headers, rows) in sheet_data.items():
        # Sheet 名称最多 31 字符
        safe_name = sheet_name[:31]
        ws = wb.add_worksheet(safe_name)

        # 写入表头
        for col, header in enumerate(headers):
            ws.write(0, col, safe_str(header), fmt_header)

        # 写入数据行
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                ws.write(r + 1, c, val if val is not None else "", fmt_cell)

        # 自动调整列宽（粗略估算）
        for col, header in enumerate(headers):
            max_len = len(safe_str(header))
            for row in rows[:100]:  # 采样前100行
                cell_len = len(safe_str(row[col])) if col < len(row) else 0
                max_len = max(max_len, cell_len)
            ws.set_column(col, col, min(max_len + 2, 40))

        # 冻结首行
        ws.freeze_panes(1, 0)

    wb.close()
    print(f"[OK] Output written: {output_path}")


# ============================================================
# 文件查找
# ============================================================

def find_latest_file(directory, pattern):
    """
    在目录中按修改时间查找最新文件
    Args:
        directory: 搜索目录
        pattern: glob 模式，如 "Monthly Closing_Data_Base_*.xlsx"
    Returns: 最新文件的完整路径或 None
    """
    full_pattern = os.path.join(directory, pattern) if not os.path.isabs(pattern) else pattern
    files = glob.glob(full_pattern)
    if not files:
        print(f"[WARNING] No files found matching: {full_pattern}")
        return None
    latest = max(files, key=os.path.getmtime)
    return latest


# ============================================================
# 项目号处理
# ============================================================

def extract_project_key(wbs_element, wbs_map=None):
    """
    从 WBS element 提取项目号前缀
    e.g. "CNTS-25165-11000" → "CNTS-25165"
         "CN03-22052-10026" → "CN03-22052"
    """
    if not wbs_element:
        return None
    wbs = safe_str(wbs_element).strip()
    # Match pattern: CNxx-xxxxx (2 letters, 2 digits, dash, 5 digits)
    import re
    match = re.match(r'(CN[A-Z0-9]{2,4}-\d{4,6})', wbs)
    if match:
        return match.group(1)
    return wbs


def read_sheet_raw(path, sheet_name=None):
    """
    读取一个 sheet 的全部数据，保留原始值（不转为字符串）
    Returns: (headers list, rows list of lists with raw values)
    """
    wb = safe_load_workbook(path)
    if wb is None:
        return [], []

    if sheet_name is None:
        sheet_name = wb.sheetnames[0]

    if sheet_name not in wb.sheetnames:
        print(f"[WARNING] Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
        wb.close()
        return [], []

    ws = wb[sheet_name]
    rows = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
        rows.append(list(row))

    wb.close()
    headers = rows[0] if rows else []
    data = rows[1:] if len(rows) > 1 else []
    return headers, data


# ============================================================
# 名称标准化
# ============================================================

def normalize_name_last_first(raw_name):
    """
    将 "Last, First" 格式转为 "First Last" 格式
    处理边界情况：
      "Tang, Qiwei"  → "Tang Qiwei"
      "Li, Xiaoxuan" → "Li Xiaoxuan"
      "Tang Qiwei"   → "Tang Qiwei" (不变)
      中文名原样保留
    """
    if not raw_name:
        return ""
    name = str(raw_name).strip()
    # "Last, First" 格式
    if ',' in name:
        parts = name.split(',', 1)
        last = parts[0].strip()
        first = parts[1].strip() if len(parts) > 1 else ""
        return f"{first} {last}".strip()
    return name


# ============================================================
# 日期/月份解析
# ============================================================

def parse_date_month(date_value):
    """
    从各种日期格式中提取 "YYYY-MM" 字符串
    支持:
      - datetime.datetime(2026, 1, 26) → "2026-01"
      - "2026-01-26 00:00:00" → "2026-01"
      - "2026-Jan" → "2026-01"
      - 2026-01-26 (date类型) → "2026-01"
      - "2026-01" → "2026-01" (不变)
    Returns: "YYYY-MM" str or None
    """
    import re
    if date_value is None:
        return None

    # datetime 对象
    if isinstance(date_value, datetime):
        return date_value.strftime("%Y-%m")

    s = str(date_value).strip()
    if not s:
        return None

    # "2026-01-26" 或 "2026-01-26 00:00:00"
    match = re.match(r'(\d{4})-(\d{2})-\d{2}', s)
    if match:
        return f"{match.group(1)}-{match.group(2)}"

    # "2026-Jan"
    month_names = {
        "jan": "01", "feb": "02", "mar": "03", "apr": "04",
        "may": "05", "jun": "06", "jul": "07", "aug": "08",
        "sep": "09", "oct": "10", "nov": "11", "dec": "12",
        "january": "01", "february": "02", "march": "03", "april": "04",
        "june": "06", "july": "07", "august": "08",
        "september": "09", "october": "10", "november": "11", "december": "12",
    }
    match = re.match(r'(\d{4})-([A-Za-z]+)', s)
    if match:
        year = match.group(1)
        month_name = match.group(2).lower()
        if month_name in month_names:
            return f"{year}-{month_names[month_name]}"

    # 已经是 "YYYY-MM"
    if re.match(r'^\d{4}-\d{2}$', s):
        return s

    return None


def infer_level_from_rate(rate, rate_map=None):
    """
    从日费率推断 Level
    2460 → "Testing Engineer-Expert"
    2260 → "Testing Engineer-Senior"
    2100 → "Testing Engineer-Experienced"
    1600 → "Testing Engineer-Junior"

    使用最近值匹配（允许小偏差）
    """
    if rate_map is None:
        rate_map = {
            2460: "Testing Engineer-Expert",
            2260: "Testing Engineer-Senior",
            2100: "Testing Engineer-Experienced",
            1600: "Testing Engineer-Junior",
        }
    try:
        r = float(rate)
    except (TypeError, ValueError):
        return None

    # 精确匹配
    if r in rate_map:
        return rate_map[r]

    # 最近值匹配 (容差 5%)
    closest = min(rate_map.keys(), key=lambda k: abs(k - r))
    if abs(closest - r) / closest < 0.05:
        return rate_map[closest]

    return None


def normalize_wbs(wbs_element):
    """规范化 WBS 格式，去除全角字符等"""
    if not wbs_element:
        return ""
    s = safe_str(wbs_element).strip()
    # 全角英文字符替换（SAP 系统中常见）
    fullwidth_map = {
        'Ｍ': 'M', 'Ａ': 'A', 'Ｐ': 'P', 'Ｎ': 'N', 'Ｇ': 'G',
        'Ｉ': 'I', '＆': '&', '０': '0', '１': '1', '２': '2',
        '３': '3', '４': '4', '５': '5', '６': '6', '７': '7',
        '８': '8', '９': '9', '－': '-',
    }
    for fw, hw in fullwidth_map.items():
        s = s.replace(fw, hw)
    return s
