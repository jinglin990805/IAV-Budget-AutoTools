# -*- coding: utf-8 -*-
"""
第三方人员工时 PDF 解析器

从 CNTS-25165/01_IAV TMS & WP/ 下的工时 PDF 中提取:
- 人员姓名 (中英文)
- 月份、年份
- PO 编号
- 当月总工时

支持两种来源:
1. 文本型 PDF (pdfplumber 直接提取)
2. 文件名解析 (YYYYMM姓名IAV.pdf，作为备用/补充)
"""

import os
import re
import pdfplumber
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class TimesheetRecord:
    """单条工时记录"""
    name_cn: str           # 中文名
    name_en: str           # 英文名
    year: int
    month: int
    po_number: str         # PO 编号
    total_hours: float     # 当月总工时
    source: str            # 数据来源: "pdf_text" | "filename_only" | "manual"
    filepath: str          # 源文件路径
    has_text: bool         # PDF 是否有可提取文本


def parse_filename(filepath: str) -> Optional[dict]:
    """
    从文件名解析: YYYYMM姓名IAV.pdf
    例: 202604侯景IAV.pdf → name_cn=侯景, year=2026, month=4
    """
    basename = os.path.basename(filepath)
    match = re.match(r'(\d{4})(\d{2})(.+?)IAV.*?\.pdf', basename)
    if match:
        return {
            "year": int(match.group(1)),
            "month": int(match.group(2)),
            "name_cn": match.group(3).strip(),
        }
    return None


def parse_timesheet_pdf(filepath: str) -> Optional[TimesheetRecord]:
    """
    解析单个工时 PDF，提取工时信息

    支持的文本格式 (来自 VCTC-AI5 工时报告):
    - 姓名 Name: <Name>
    - 要货单号 Purchase Order No.:<PO Number>
    - 月/年 Month / Year: <MM> / <YYYY>
    - 本月实际工作时间 Exact Work-hour of this month : <Hours>Hours
    """
    # 先从文件名获取基本信息
    filename_info = parse_filename(filepath)
    if not filename_info:
        return None

    record = TimesheetRecord(
        name_cn=filename_info["name_cn"],
        name_en="",
        year=filename_info["year"],
        month=filename_info["month"],
        po_number="",
        total_hours=0.0,
        source="filename_only",
        filepath=filepath,
        has_text=False,
    )

    # 尝试用 pdfplumber 提取文本
    try:
        pdf = pdfplumber.open(filepath)
    except Exception:
        return record  # 无法打开，返回文件名解析结果

    full_text = ""
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            full_text += text + "\n"
    pdf.close()

    if not full_text.strip():
        return record  # 扫描件，无文本

    record.has_text = True
    record.source = "pdf_text"

    # === 提取姓名 ===
    # 格式: "姓名 Name: Sun Shumin" 或 "姓名 Name: YinZicheng"
    name_match = re.search(r'姓名\s*Name[:\s]+([A-Za-z\s]+?)(?:\n|要货|公司|中/)', full_text)
    if name_match:
        record.name_en = name_match.group(1).strip()
    else:
        name_match = re.search(r'Name[:\s]+([A-Z][a-z]+\s*[A-Z]?[a-z]*)', full_text)
        if name_match:
            record.name_en = name_match.group(1).strip()

    # === 提取 PO 编号 ===
    po_match = re.search(r'(?:Purchase Order No\.?|要货单号)[:\s]*(PO\d{8,})', full_text, re.IGNORECASE)
    if not po_match:
        po_match = re.search(r'(?:PO|Purchase Order)[:\s]*(\d{8,})', full_text, re.IGNORECASE)
    if po_match:
        record.po_number = po_match.group(1)

    # === 提取月份/年份 (作为交叉验证) ===
    month_match = re.search(r'(?:Month\s*/\s*Year|月\s*/\s*年)[:\s]*(\d{1,2})\s*/\s*(\d{4})', full_text, re.IGNORECASE)
    if month_match:
        pdf_month = int(month_match.group(1))
        pdf_year = int(month_match.group(2))
        # 用 PDF 内容覆盖文件名解析结果 (PDF 内容更可靠)
        record.year = pdf_year
        record.month = pdf_month

    # === 提取总工时 ===
    # 格式: "本月实际工作时间 Exact Work-hour of this month : 168Hours"
    # 注意: 有些 PDF 可能写作 "Exacta Work-hour" (typo) 或格式略有差异
    hours_match = re.search(
        r'(?:Exacta?\s*Work-hour|实际工作时间|Total\s*[Hh]ours?)[^:]*[:\s]+([\d.]+)\s*(?:Hours?|H|h)',
        full_text, re.IGNORECASE
    )
    if hours_match:
        try:
            record.total_hours = float(hours_match.group(1))
        except ValueError:
            pass  # 保持 0.0
    else:
        # 备用: 更宽泛的 "XXX Hours" 匹配
        # 找最后一个出现在文本中的 "数字Hours" 模式
        all_hours = re.findall(r'(\d{2,4})\s*Hours?\b', full_text, re.IGNORECASE)
        if all_hours:
            # 取最大的数字作为总工时 (通常总工时最大)
            record.total_hours = float(max(all_hours, key=int))

    return record


def scan_timesheet_directory(base_dir: str, months: Optional[list[str]] = None) -> list[TimesheetRecord]:
    """
    扫描工时目录，解析所有 PDF

    Args:
        base_dir: 工时根目录 (如 CNTS-25165/01_IAV TMS & WP/)
        months: 要扫描的月份子目录列表 (如 ["2604", "2605"])，None 表示全部

    Returns:
        解析出的所有工时记录
    """
    records = []

    if months is None:
        # 自动扫描所有 YYMM 格式的子目录
        subdirs = []
        for entry in os.listdir(base_dir):
            entry_path = os.path.join(base_dir, entry)
            if os.path.isdir(entry_path) and re.match(r'\d{4}$', entry):
                subdirs.append(entry)
        months = sorted(subdirs)

    for month_dir in months:
        dir_path = os.path.join(base_dir, month_dir)
        if not os.path.isdir(dir_path):
            continue

        for filename in os.listdir(dir_path):
            if not filename.endswith('.pdf'):
                continue

            # 跳过 archive/resign 等特殊目录中的文件
            filepath = os.path.join(dir_path, filename)
            if 'archive' in filepath.lower() or 'resign' in filepath.lower():
                continue

            record = parse_timesheet_pdf(filepath)
            if record:
                records.append(record)

    return records


def records_to_summary(records: list[TimesheetRecord]) -> dict:
    """
    将工时记录汇总为 {姓名: {YYYY-MM: hours}} 格式
    用于分配算法
    """
    summary = {}
    for r in records:
        key = r.name_cn if r.name_cn else r.name_en
        if key not in summary:
            summary[key] = {}
        month_key = f"{r.year}-{r.month:02d}"
        # 同名同月可能有多份 PDF（(001), (002)），取合计
        if month_key in summary[key]:
            summary[key][month_key] += r.total_hours
        else:
            summary[key][month_key] = r.total_hours
    return summary


def print_records(records: list[TimesheetRecord]) -> None:
    """打印工时记录摘要"""
    import sys
    import io
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    print(f"\n{'='*80}")
    print(f"Total records: {len(records)}")
    text_count = sum(1 for r in records if r.has_text)
    scan_count = sum(1 for r in records if not r.has_text)
    print(f"  Text-based (pdfplumber): {text_count}")
    print(f"  Scanned (filename only): {scan_count}")
    print(f"{'='*80}")

    # 按月份分组
    by_month = {}
    for r in records:
        mk = f"{r.year}-{r.month:02d}"
        by_month.setdefault(mk, []).append(r)

    for month in sorted(by_month.keys()):
        month_records = by_month[month]
        print(f"\n--- {month} ({len(month_records)} records) ---")
        for r in sorted(month_records, key=lambda x: x.name_cn):
            source_icon = "📄" if r.has_text else "🖼️"
            print(f"  {source_icon} {r.name_cn} ({r.name_en}): "
                  f"{r.total_hours:.0f}h | PO: {r.po_number}")


if __name__ == "__main__":
    import sys
    import io
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    base = (
        r"C:\Users\LinJing(CN-TV-F)\Desktop\Jing\00_IAV_TV"
        r"\02_Budget Report\CNTS-25165\01_IAV TMS & WP"
    )

    # 测试: 只扫描 2026年4月和5月
    print("=== Timesheet PDF Parser Test ===\n")
    records = scan_timesheet_directory(base, months=["2604", "2605"])
    print_records(records)

    print("\n\n--- Summary (for allocation algorithm) ---")
    summary = records_to_summary(records)
    for name in sorted(summary.keys()):
        months_data = summary[name]
        total = sum(months_data.values())
        print(f"  {name}: {total:.0f}h total | " + " | ".join(
            f"{m}: {h:.0f}h" for m, h in sorted(months_data.items())
        ))
