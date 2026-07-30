# -*- coding: utf-8 -*-
"""
PO PDF 解析器 — 从客户 PO PDF 中提取 Level 配额信息

支持的 PO 格式: VCTC Call-off Purchase Order
提取内容: PO编号、日期、各Level的天数配额和单价、总金额、交货日期
"""

import re
import pdfplumber
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class POLevelQuota:
    """单个 Level 的配额"""
    item: int
    qty_days: float
    unit: str
    description: str
    unit_price: float
    total_rmb: float


@dataclass
class POInfo:
    """PO 完整信息"""
    po_number: str
    po_date: str
    sap_po_number: str = ""
    contract_number: str = ""
    total_excl_vat: float = 0.0
    vat_rate: float = 0.0
    vat_amount: float = 0.0
    total_incl_vat: float = 0.0
    delivery_date: str = ""
    level_quotas: list[POLevelQuota] = field(default_factory=list)


# Level 描述关键词 → VCTC level 名称映射
LEVEL_KEYWORD_MAP = {
    "senior engineer": "Testing Engineer-Senior",
    "expert engineer": "Testing Engineer-Expert",
    "experienced engineer": "Testing Engineer-Experienced",
    "experience engineer": "Testing Engineer-Experienced",
    "junior engineer": "Testing Engineer-Junior",
    "senior": "Testing Engineer-Senior",
    "expert": "Testing Engineer-Expert",
    "experienced": "Testing Engineer-Experienced",
    "junior": "Testing Engineer-Junior",
    "engineer-senior": "Testing Engineer-Senior",
    "engineer-expert": "Testing Engineer-Expert",
    "engineer-experienced": "Testing Engineer-Experienced",
    "engineer-junior": "Testing Engineer-Junior",
}

# 已知的 PO Description 精确映射
KNOWN_DESCRIPTION_MAP = {
    "WP2 Experience Engineer": "Testing Engineer-Experienced",
    "WP1/WP2 Senior Engineer": "Testing Engineer-Senior",
    "WP1/WP2 Expert Engineer": "Testing Engineer-Expert",
    "WP2 Junior Engineer": "Testing Engineer-Junior",
}


def match_level(description: str) -> Optional[str]:
    """将 PO 中的描述文本匹配到标准 Level 名称"""
    desc_lower = description.strip().lower()

    # 先精确匹配
    for exact, level in KNOWN_DESCRIPTION_MAP.items():
        if exact.lower() == desc_lower:
            return level

    # 关键词模糊匹配
    for keyword, level in LEVEL_KEYWORD_MAP.items():
        if keyword in desc_lower:
            return level

    return None


def parse_po_pdf(filepath: str) -> Optional[POInfo]:
    """
    解析 VCTC Call-off Purchase Order PDF

    提取:
    - PO number (e.g., PO-VCTC00-2607000312)
    - 各 Level 的 Days × Unit Price = Total
    - 总金额、交货日期
    """
    try:
        pdf = pdfplumber.open(filepath)
    except Exception as e:
        print(f"  [ERROR] Cannot open PDF: {filepath} — {e}")
        return None

    info = POInfo(po_number="", po_date="")
    full_text = ""

    for page in pdf.pages:
        text = page.extract_text()
        if text:
            full_text += text + "\n"

        # 也尝试从表格提取
        tables = page.extract_tables()
        for table in tables:
            for row in table:
                if row:
                    row_text = " | ".join([str(c) if c else "" for c in row])
                    full_text += row_text + "\n"

    pdf.close()

    # === 提取 PO 编号 ===
    po_match = re.search(r'PO[-\s]?number[^P]*?(PO-?\w*?\d{8,})', full_text, re.IGNORECASE)
    if not po_match:
        po_match = re.search(r'(PO-VCTC\d{2}-\d{8,})', full_text)
    if po_match:
        info.po_number = po_match.group(1)

    # === 提取 PO 日期 ===
    date_match = re.search(r'(?:PO.*?date.*?|日期.*?)[:\s]*([A-Z][a-z]+ \d{1,2},?\s*\d{4})', full_text, re.IGNORECASE)
    if not date_match:
        date_match = re.search(r'(?:Date:|日期:)\s*([A-Z][a-z]+ \d{1,2},?\s*\d{4})', full_text)
    if date_match:
        info.po_date = date_match.group(1)

    # === 提取 SAP PO 编号 ===
    sap_match = re.search(r'SAP\s*(?:PO|订单)\s*(?:number|编号)[:\s]*(\d{8,})', full_text, re.IGNORECASE)
    if sap_match:
        info.sap_po_number = sap_match.group(1)

    # 预处理: 修复 level 描述跨行问题
    # PO PDF 中 "Engineer" 单独一行，需要合并到前一行描述中（在价格数字之前）
    # 例: "1 180.000 Days WP2 Experience 2,100.00 378,000.00\nEngineer"
    #   → "1 180.000 Days WP2 Experience Engineer 2,100.00 378,000.00"
    full_text = re.sub(
        r'(\d+\s+[\d.]+\s+Days\s+.+?)(\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2})\n(Engineer|engineer|Senior|Expert|Experienced|Junior)',
        r'\1 \3\2',
        full_text
    )
    # 也处理没有 "Days" 关键字的情况 (如 "Senior\nEngineer" → "Senior Engineer")
    full_text = re.sub(r'\n(Senior|Expert|Experienced|Junior)\s*\n', r' \1\n', full_text)

    # === 提取表格行: Item | QTY | Unit | Description | Price/Unit | Total ===
    # 现在 Description 应该在同一行了
    item_pattern = re.compile(
        r'(\d{1,3})\s+([\d,]+\.?\d*)\s+(Days?)\s+'
        r'([A-Za-z0-9/\s]+?\s*(?:Engineer|Senior|Expert|Experienced|Junior))'
        r'\s+([\d,]+\.?\d{2})\s+([\d,]+\.?\d{2})'
    )
    for match in item_pattern.finditer(full_text):
        item = int(match.group(1))
        qty = float(match.group(2).replace(",", ""))
        unit = match.group(3)
        desc = match.group(4).strip()
        price = float(match.group(5).replace(",", ""))
        total = float(match.group(6).replace(",", ""))

        quota = POLevelQuota(
            item=item,
            qty_days=qty,
            unit=unit,
            description=desc,
            unit_price=price,
            total_rmb=total,
        )
        info.level_quotas.append(quota)

    # === 提取总金额 ===
    total_match = re.search(
        r'Total\s*value\s*excluding\s*VAT[^R]*RMB\s*([\d,]+\.?\d{0,2})',
        full_text, re.IGNORECASE
    )
    if total_match:
        info.total_excl_vat = float(total_match.group(1).replace(",", ""))

    total_incl_match = re.search(
        r'Total\s*Purchase\s*Order\s*Amount\s*incl[^R]*RMB\s*([\d,]+\.?\d{0,2})',
        full_text, re.IGNORECASE
    )
    if total_incl_match:
        info.total_incl_vat = float(total_incl_match.group(1).replace(",", ""))

    vat_match = re.search(r'VAT\s*Amount[^R]*RMB\s*([\d,]+\.?\d{0,2})', full_text, re.IGNORECASE)
    if vat_match:
        info.vat_amount = float(vat_match.group(1).replace(",", ""))

    vat_rate_match = re.search(r'VAT\s*%[^:]*:\s*([\d.]+)\s*%', full_text, re.IGNORECASE)
    if vat_rate_match:
        info.vat_rate = float(vat_rate_match.group(1))

    # === 提取交货日期 ===
    delivery_match = re.search(
        r'Delivery\s*Date[^:]*:\s*([A-Z][a-z]+ \d{1,2},?\s*\d{4})',
        full_text, re.IGNORECASE
    )
    if delivery_match:
        info.delivery_date = delivery_match.group(1)

    return info


def po_to_config_rows(po_info: POInfo) -> list[dict]:
    """
    将 POInfo 转换为 po_config_v2 的行格式
    每行: PO_ID | Value_RMB | Level | Unit_Price | Days_Quota
    """
    rows = []
    for quota in po_info.level_quotas:
        level = match_level(quota.description)
        rows.append({
            "PO_ID": po_info.po_number,
            "Value_RMB": quota.total_rmb,
            "Level": level or quota.description,
            "Level_Description": quota.description,
            "Unit_Price": quota.unit_price,
            "Days_Quota": quota.qty_days,
            "Delivery_Date": po_info.delivery_date,
        })
    return rows


def print_po_info(po_info: POInfo) -> None:
    """打印 PO 信息（用于验证）"""
    import sys
    import io
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    print(f"PO Number:      {po_info.po_number}")
    print(f"PO Date:        {po_info.po_date}")
    print(f"SAP PO Number:  {po_info.sap_po_number}")
    print(f"Total (excl.VAT): RMB {po_info.total_excl_vat:,.2f}")
    print(f"Delivery Date:  {po_info.delivery_date}")
    print(f"Level Quotas ({len(po_info.level_quotas)} items):")
    for q in po_info.level_quotas:
        level = match_level(q.description)
        print(f"  [{q.item}] {q.qty_days:,.0f} {q.unit} × ¥{q.unit_price:,.0f}/天 = ¥{q.total_rmb:,.0f}  "
              f"({q.description}) → {level or 'UNKNOWN'}")


if __name__ == "__main__":
    import sys
    import io
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    # 测试: 解析 2026 Q3 PO
    test_path = (
        r"C:\Users\LinJing(CN-TV-F)\Desktop\Jing\00_IAV_TV"
        r"\02_Budget Report\CNTS-25165\00_PO & CBD\26Q3"
        r"\PO-VCTC00-2607000312.pdf"
    )
    print("=== PO PDF Parser Test ===\n")
    info = parse_po_pdf(test_path)
    if info:
        print_po_info(info)
        print("\n--- Config Rows ---")
        for row in po_to_config_rows(info):
            print(f"  {row}")
    else:
        print("Failed to parse PO PDF.")
