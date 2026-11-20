# -*- coding: utf-8 -*-
"""
PO PDF parser - extracts Level quota info from customer PO PDFs.

Supported formats:
  - VCTC Call-off (PO-VCTC00-xxxxxxxx)
  - CARIAD CN Call-off (PO-CARIADCN-xxxxxxxx)
  - VWGC (66xxxxxxx, PO-VCTC00-xxxxxxxx)
  - BMW (4700xxxxxx)
  - CARIAD Contract Call Off (7300xxxxxx)
"""

import re
import pdfplumber
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class POLevelQuota:
    item: int
    qty_days: float
    unit: str
    description: str
    unit_price: float
    total_rmb: float


@dataclass
class POInfo:
    po_number: str
    po_date: str
    sap_po_number: str = ""
    contract_number: str = ""
    total_excl_vat: float = 0.0
    vat_rate: float = 0.0
    vat_amount: float = 0.0
    total_incl_vat: float = 0.0
    delivery_date: str = ""
    level_quotas: list = field(default_factory=list)


LEVEL_KEYWORDS = [
    ("senior engineer", "Testing Engineer-Senior"),
    ("expert engineer", "Testing Engineer-Expert"),
    ("experienced engineer", "Testing Engineer-Experienced"),
    ("experience engineer", "Testing Engineer-Experienced"),
    ("junior engineer", "Testing Engineer-Junior"),
    ("senior", "Testing Engineer-Senior"),
    ("expert", "Testing Engineer-Expert"),
    ("experienced", "Testing Engineer-Experienced"),
    ("junior", "Testing Engineer-Junior"),
    ("engineer-senior", "Testing Engineer-Senior"),
    ("engineer-expert", "Testing Engineer-Expert"),
    ("engineer-experienced", "Testing Engineer-Experienced"),
    ("engineer-junior", "Testing Engineer-Junior"),
    ("fse (cn)", "FSE (CN)"),
    ("fse (ger)", "FSE (GER)"),
    ("tf4 local", "TF4 Local (SP)"),
    ("tf3 local", "TF3 Local (SP)"),
    ("tf1 local", "TF1 Local (SP)"),
    ("te2 local", "TE2 Local (SP)"),
    ("te3 local", "TE3 Local (SP)"),
    ("te local", "TE Local (SP)"),
    ("ticket s", "Ticket S"),
    ("ticket m", "Ticket M"),
    ("ticket l", "Ticket L"),
    ("testing engineer", "Testing Engineer"),
]

KNOWN_DESCRIPTIONS = {
    "WP2 Experience Engineer": "Testing Engineer-Experienced",
    "WP1/WP2 Senior Engineer": "Testing Engineer-Senior",
    "WP1/WP2 Expert Engineer": "Testing Engineer-Expert",
    "WP2 Junior Engineer": "Testing Engineer-Junior",
}


def match_level(description):
    desc = description.strip()
    dlower = desc.lower()
    for exact, level in KNOWN_DESCRIPTIONS.items():
        if exact.lower() == dlower:
            return level
    for keyword, level in LEVEL_KEYWORDS:
        if keyword in dlower:
            return level
    return None


def _extract_text(pdf):
    parts = []
    for page in pdf.pages:
        t = page.extract_text()
        if t:
            parts.append(t)
        for table in (page.extract_tables() or []):
            for row in table:
                if row:
                    parts.append(" | ".join([str(c) if c else "" for c in row]))
    return "\n".join(parts)


def _extract_ocr_text(pdf):
    """Extract text from image-only PO pages with the existing Tesseract stack."""
    try:
        import config
        from document_processing.tesseract_provider import TesseractProvider

        provider = TesseractProvider(config.TESSERACT_EXE, languages="eng", timeout=90)
        if not provider.available:
            return ""
        parts = []
        for page_number, page in enumerate(pdf.pages, 1):
            image = page.to_image(resolution=getattr(config, "OCR_DPI", 300)).original
            result = provider.recognize(image, psm="6", page=page_number)
            if result.text.strip():
                parts.append(result.text)
        return "\n".join(parts)
    except Exception:
        # OCR is a best-effort fallback; callers retain the normal review path
        # when the rendering or OCR stack is unavailable.
        return ""


def _clean_number(s):
    if not s:
        return 0.0
    s = str(s).strip()
    s = re.sub(r"(?:CNY|RMB).*$", "", s, flags=re.I).strip()
    last_dot = s.rfind(".")
    last_comma = s.rfind(",")
    decimal_pos = max(last_dot, last_comma)
    decimal_sep = None
    if decimal_pos >= 0:
        after = s[decimal_pos + 1:]
        if len(after) <= 2:
            decimal_sep = s[decimal_pos]
        elif last_comma > last_dot:
            decimal_sep = ","
    if decimal_sep == ",":
        s = s.replace(".", "").replace(",", ".")
    elif decimal_sep == ".":
        s = s.replace(",", "")
    else:
        s = s.replace(",", "")
    return float(s)


def parse_po_pdf(filepath):
    try:
        pdf = pdfplumber.open(filepath)
    except Exception as e:
        print("  [ERROR] Cannot open PDF: {} - {}".format(filepath, e))
        return None

    full_text = _extract_text(pdf)
    if not full_text.strip():
        full_text = _extract_ocr_text(pdf)
    pdf.close()
    if not full_text.strip():
        return None

    # Fix cross-line level names (e.g. "WP2 Experience\nEngineer")
    full_text = re.sub(r"\n(Engineer|engineer)", r" \1", full_text)

    info = POInfo(po_number="", po_date="")

    # ---- PO Number ----
    for pat in [
        r"PO-CARIADCN-(\d{8,})",
        r"PO-VCTC\d{2}-(\d{8,})",
        r"\b(66\d{8})\b",
        r"PO\s*(?:Number|编号)[^\d]{0,30}?(\d{10})",
        r"(?:Doc(?:ument)?\s*Number|文档号码)\s*(\d{10})",
        r"SAP\s*(?:PO|订单)\s*(?:number|编号)[:\s]*(\d{8,})",
        r"PO[- ]?(?:Number|编号)?[:\s/]*([A-Z]{2,4}[- ]?\d{8,})",
        r"订单编号[:\s/]*(\d{8,})",
    ]:
        m = re.search(pat, full_text, re.I)
        if m:
            digits = re.search(r"(\d{7,})", m.group(1).replace("-", "").replace(" ", ""))
            info.po_number = digits.group(1) if digits else m.group(1).strip()
            break

    # ---- SAP PO Number ----
    sap_m = re.search(r"SAP\s*(?:PO|订单)\s*(?:number|编号)[:\s]*(\d{8,})", full_text, re.I)
    if not sap_m:
        sap_m = re.search(r"8500\d{6}", full_text)
    if sap_m:
        info.sap_po_number = sap_m.group(1) if sap_m.lastindex else sap_m.group(0)

    # ---- Date ----
    for pat in [
        r"(?:Date|日期)[\s:/]*\s*([A-Z][a-z]{2,8}\s+\d{1,2},?\s*\d{4})",
        r"(?:Doc\.?Date|文档日期)\s*(\d{2}\.\d{2}\.\d{4})",
        r"(\d{2}\.\d{2}\.\d{4})",
    ]:
        m = re.search(pat, full_text, re.I)
        if m:
            info.po_date = m.group(1)
            break

    # ---- Total (excl VAT) ----
    for pat in [
        r"Total\s*value\s*excluding\s*VAT[^R]*RMB\s*(\d[\d,.]*)",
        r"Net\s+amount[^\n]{0,100}?CNY\s*(\d[\d,.]*)",
        r"未税总金额[^R]*RMB\s*(\d[\d,.]*)",
        r"Net\s*amount不含税总价\s*CNY\s*(\d[\d,.]*)",
        r"Gross\s*Value\s*in\s*CNY\s*(\d[\d,.]*)",
        r"总额[：:]\s*CNY\s*(\d[\d,.]*)",
        r"Total\s*[计計]?\s*CNY\s*(\d[\d,.]*)",
    ]:
        m = re.search(pat, full_text, re.I)
        if m:
            info.total_excl_vat = _clean_number(m.group(1))
            break

    # ---- Line Items ----
    level_quotas = []

    # A: VCTC table rows
    for m in re.finditer(
        r"(\d{1,3})\s+([\d,.]+)\s+(Days?)\s+"
        r"([A-Za-z0-9/\s]+?\s*(?:Engineer|Senior|Expert|Experienced|Junior|Experience))\s+"
        r"([\d,.]+)\s+([\d,.]+)",
        full_text
    ):
        level_quotas.append(POLevelQuota(
            item=int(m.group(1)), qty_days=_clean_number(m.group(2)),
            unit=m.group(3), description=m.group(4).strip(),
            unit_price=_clean_number(m.group(5)), total_rmb=_clean_number(m.group(6)),
        ))

    # B: CARIAD / VWGC
    if not level_quotas:
        for m in re.finditer(
            r"(\d{1,5})\s+([\d,.]+)\s+(each|EA|Day)\s+(.+?)\s+([\d,.]+)\s+([\d,.]+)\s*$",
            full_text, re.M
        ):
            qty = _clean_number(m.group(2))
            desc = re.sub(r"\s*(?:Material\s*Number|Delivery\s*Date|Accounting|CC-).*$", "",
                          m.group(4), flags=re.I).strip()
            price = _clean_number(m.group(5))
            total = _clean_number(m.group(6))
            if qty < 1 or total < 100:
                continue
            level_quotas.append(POLevelQuota(
                item=int(m.group(1)), qty_days=qty,
                unit=m.group(3), description=desc,
                unit_price=price, total_rmb=total,
            ))

    # C: BMW
    if not level_quotas:
        for m in re.finditer(
            r"(\d{5})\s+(.+?)\s+([\d,.]+)\s+(EA|ea)\s+([\d,.]+)\s+([\d,.]+)",
            full_text
        ):
            desc = re.sub(r"\s*合同.*$", "", m.group(2).strip(), flags=re.I).strip()
            qty = _clean_number(m.group(3))
            price = _clean_number(m.group(5))
            total = _clean_number(m.group(6))
            level_quotas.append(POLevelQuota(
                item=int(m.group(1)), qty_days=qty,
                unit=m.group(4), description=desc,
                unit_price=price, total_rmb=total,
            ))

    # D: CARIAD Contract Call Off lump-sum
    if not level_quotas:
        for m in re.finditer(
            r"(\d{1,3})\s+([\d,.]+)\s+(EA|ea)\s+(.+?)\s+([\d,.]+)\s+CNY",
            full_text
        ):
            total_val = _clean_number(m.group(2))
            unit_price = _clean_number(m.group(5))
            desc = m.group(4).strip()
            if unit_price <= 2 and abs(total_val - 1) > 0.01 and total_val > 100:
                level_quotas.append(POLevelQuota(
                    item=int(m.group(1)), qty_days=total_val,
                    unit=m.group(3), description=desc,
                    unit_price=1.0, total_rmb=total_val,
                ))
            elif total_val > 100:
                level_quotas.append(POLevelQuota(
                    item=int(m.group(1)), qty_days=1.0,
                    unit=m.group(3), description=desc,
                    unit_price=total_val, total_rmb=total_val,
                ))

    # E: VWGC with CNY/1 EA (last resort)
    if not level_quotas:
        for m in re.finditer(
            r"(\d{5})\s+([\d,.]+)\s+(EA|ea)\s+(.+?)\s+([\d,.]+)\s+CNY(?:\s*/\s*1\s*(?:EA|ea))?",
            full_text
        ):
            qty = _clean_number(m.group(2))
            desc = re.sub(r"\s*(?:Material\s*Number|Delivery\s*Date|Accounting|CC-).*$", "",
                          m.group(4), flags=re.I).strip()
            price = _clean_number(m.group(5))
            level_quotas.append(POLevelQuota(
                item=int(m.group(1)), qty_days=qty,
                unit=m.group(3), description=desc,
                unit_price=price, total_rmb=qty * price,
            ))

    # Dedup by item number + total
    seen = set()
    deduped = []
    for q in level_quotas:
        key = (q.item, round(q.total_rmb, 2))
        if key not in seen:
            seen.add(key)
            deduped.append(q)
    info.level_quotas = deduped

    # Fallback total from line items (threshold 15% for VAT-inclusive totals)
    line_sum = sum(q.total_rmb for q in info.level_quotas)
    if line_sum > 0 and (info.total_excl_vat <= 0 or abs(info.total_excl_vat - line_sum) / line_sum > 0.15):
        info.total_excl_vat = line_sum

    # ---- Additional metadata ----
    total_incl = re.search(
        r"(?:Total\s*amount\s*incl|含税总价|Total\s*Purchase\s*Order\s*Amount\s*incl)[^\d]*?(?:RMB|CNY)?\s*(\d[\d,.]*)",
        full_text, re.I)
    if total_incl:
        info.total_incl_vat = _clean_number(total_incl.group(1))
    else:
        total_incl = re.search(
            r"Total\s+amount[^\n]{0,100}?CNY\s*(\d[\d,.]*)",
            full_text, re.I)
        if total_incl:
            info.total_incl_vat = _clean_number(total_incl.group(1))

    vat_amt = re.search(r"VAT\s*Amount[^R]*RMB\s*(\d[\d,.]*)", full_text, re.I)
    if vat_amt:
        info.vat_amount = _clean_number(vat_amt.group(1))

    vat_pct = re.search(r"VAT\s*%\s*(?:\([^)]+\))?\s*[:\s]*\s*([\d.]+)\s*%", full_text, re.I)
    if vat_pct:
        info.vat_rate = float(vat_pct.group(1))

    delivery = re.search(r"Delivery\s*Date[^:]*:\s*(\d{2}\.\d{2}\.\d{4})", full_text, re.I)
    if not delivery:
        delivery = re.search(r"Delivery\s*Date[^:]*:\s*([A-Za-z0-9 ,/]+)", full_text, re.I)
    if not delivery:
        delivery = re.search(r"交货日期.*?:\s*([^\n]{5,40})", full_text)
    if delivery:
        info.delivery_date = delivery.group(1).strip()

    return info if (info.po_number or info.level_quotas or info.total_excl_vat > 0) else None


def po_to_config_rows(po_info):
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


def print_po_info(po_info):
    import sys, io
    if sys.stdout.encoding != "utf-8":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    print("PO Number:      {}".format(po_info.po_number))
    print("PO Date:        {}".format(po_info.po_date))
    print("SAP PO Number:  {}".format(po_info.sap_po_number))
    print("Total (excl.VAT): {:,.2f}".format(po_info.total_excl_vat))
    print("Delivery Date:  {}".format(po_info.delivery_date))
    print("Level Quotas ({} items):".format(len(po_info.level_quotas)))
    for q in po_info.level_quotas:
        level = match_level(q.description)
        print("  [{}] {:,.0f} {} x {:,.0f} = {:,.0f}  ({}) -> {}".format(
            q.item, q.qty_days, q.unit, q.unit_price, q.total_rmb,
            q.description, level or "???"
        ))


if __name__ == "__main__":
    import sys, io, os
    if sys.stdout.encoding != "utf-8":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

    tests = [
        (r"C:\Users\LinJing(CN-TV-F)\Desktop\Jing\00_IAV_TV\02_Budget Report\CNTS-25165\00_PO & CBD\26Q3\PO-VCTC00-2607000312.pdf", "VCTC Call-off"),
        (r"C:\Users\LinJing(CN-TV-F)\Desktop\Jing\00_IAV_TV\02_Budget Report\CN03-22039\03_PO\CN03-22039-12000\PO-CARIADCN-2604001064.pdf", "CARIAD CN"),
        (r"C:\Users\LinJing(CN-TV-F)\Desktop\Jing\00_IAV_TV\02_Budget Report\CN03-22052\03_PO\2026\TF4\Q3\PO_7300066245.PDF", "CARIAD Contract Call Off"),
        (r"C:\Users\LinJing(CN-TV-F)\Desktop\Jing\00_IAV_TV\02_Budget Report\TVA_Budget\02_PO\CNVA-25150\PO-VCTC00-2604000671.pdf", "VWGC (VCTC)"),
        (r"C:\Users\LinJing(CN-TV-F)\Desktop\Jing\00_IAV_TV\02_Budget Report\CN03-22060\1.3_PO\2026\4700659863_Apr-Sep_2026.PDF", "BMW"),
        (r"C:\Users\LinJing(CN-TV-F)\Desktop\Jing\00_IAV_TV\02_Budget Report\CNTS-25165\03_GR Documents\2025\Invoices\XuanLiangliang&LiuXu&SunShumin_PO6600017845_GR5000172422\PO6600017845.pdf", "VWGC (66xxx)"),
    ]

    for fpath, label in tests:
        if not os.path.exists(fpath):
            print("\nSKIP {} (not found): {}".format(label, os.path.basename(fpath)))
            continue
        print("\n" + "=" * 70)
        print("{}: {}".format(label, os.path.basename(fpath)))
        print("=" * 70)
        info = parse_po_pdf(fpath)
        if info:
            print_po_info(info)
        else:
            print("FAILED to parse")
