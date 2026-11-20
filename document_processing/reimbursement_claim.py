"""Read an expense-claim export into a normalised claim index.

The exported "Reimbursement form" sheet is the reimbursement baseline used to
check travel invoices.  It is an index, not a per-invoice amount source: most
categories contain one aggregate line per claim.  This module only reads and
normalises the workbook; it never writes to it and never decides whether an
invoice is payable.

Duplicate header names are common in these exports (for example two
``消费起止日期`` columns with data in only one of them).  Every accessor
therefore reads the first non-empty value across *all* columns carrying that
header instead of trusting a single header position.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Optional

from openpyxl import load_workbook


DEFAULT_SHEET = "Reimbursement form"
HEADER_SCAN_ROWS = 12

# Claim category buckets match the expense categories used by the travel
# reviewer because they are reconciled directly against those expenses.
CATEGORY_ACCOMMODATION = "accommodation"
CATEGORY_TICKET = "ticket"
CATEGORY_TRAFFIC = "traffic"
CATEGORY_ALLOWANCE = "allowance"
CATEGORY_OTHER = "other"

CLAIM_CATEGORY_LABELS = {
    CATEGORY_ACCOMMODATION: "Hotel",
    CATEGORY_TICKET: "Flight & Train",
    CATEGORY_TRAFFIC: "Traffic Fee",
    CATEGORY_ALLOWANCE: "BT allowance",
    CATEGORY_OTHER: "其他",
}

# Expense categories written by the travel reviewer mapped to claim buckets.
EXPENSE_CATEGORY_BUCKETS = {
    "accommodation": CATEGORY_ACCOMMODATION,
    "ticket": CATEGORY_TICKET,
    "taxi": CATEGORY_TRAFFIC,
    "special": CATEGORY_TRAFFIC,
}

_CATEGORY_PATTERNS: tuple[tuple[str, tuple[str, ...]], ...] = (
    (CATEGORY_ACCOMMODATION, ("hotel", "住宿", "宾馆", "酒店")),
    (CATEGORY_TICKET, ("flight", "train", "机票", "火车", "车票", "航空", "铁路")),
    (
        CATEGORY_TRAFFIC,
        (
            "traffic fee", "traffic", "市内交通", "交通费", "打车", "出租车",
            "网约车", "滴滴", "通行费", "停车费", "过路费", "加油费",
        ),
    ),
    (CATEGORY_ALLOWANCE, ("allowance", "补贴", "补助", "餐补")),
)

_DATE_RE = re.compile(
    r"(?<!\d)(?:20\d{2})\s*[./年-]\s*\d{1,2}\s*[./月-]\s*\d{1,2}(?!\d)"
)
_NUMBER_RE = re.compile(r"-?\d+(?:,\d{3})*(?:\.\d+)?")


def _text(value: Any) -> str:
    if value is None:
        return ""
    return unicodedata.normalize("NFKC", str(value)).strip()


def _number(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    match = _NUMBER_RE.search(_text(value).replace(" ", ""))
    if not match:
        return None
    try:
        return float(match.group(0).replace(",", ""))
    except ValueError:
        return None


def parse_date(value: Any) -> Optional[date]:
    """Parse a claim date cell or the first date inside a free-text value."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    match = _DATE_RE.search(_text(value))
    if not match:
        return None
    token = re.sub(r"\s+", "", match.group(0)).replace("月", "-").replace("日", "")
    token = token.replace("年", "-").replace("/", "-").replace(".", "-")
    parts = [int(part) for part in token.split("-") if part]
    if len(parts) != 3:
        return None
    try:
        return date(parts[0], parts[1], parts[2])
    except ValueError:
        return None


def parse_period(value: Any) -> tuple[Optional[date], Optional[date]]:
    """Return the first and last date found in a period cell."""
    if value is None:
        return None, None
    if isinstance(value, (date, datetime)):
        parsed = parse_date(value)
        return parsed, parsed
    found = [parse_date(token) for token in _DATE_RE.findall(_text(value))]
    found = [item for item in found if item]
    if not found:
        return None, None
    return min(found), max(found)


def normalize_category(type_name: Any) -> str:
    """Map a claim expense type such as ``.../Hotel`` to an expense bucket."""
    raw = _text(type_name)
    segment = re.split(r"[/／>＞]", raw)[-1].strip().lower()
    for bucket, patterns in _CATEGORY_PATTERNS:
        if any(pattern in segment for pattern in patterns):
            return bucket
    # Fall back to the whole string when the last segment is empty.
    lowered = raw.lower()
    for bucket, patterns in _CATEGORY_PATTERNS:
        if any(pattern in lowered for pattern in patterns):
            return bucket
    return CATEGORY_OTHER


def _alias_key(value: str) -> str:
    """Keep only letters, digits and CJK characters for name comparison."""
    return re.sub(r"[^0-9A-Za-z一-鿿]", "", value).lower()


def normalize_person(name: Any) -> tuple[str, set[str]]:
    """Return a primary display name plus matching aliases.

    ``邱栋（QiuDong）`` becomes ``邱栋`` with aliases ``邱栋`` and ``qiudong``.
    Punctuation is removed from aliases so ``Sun, Yatao`` matches
    ``SunYatao`` from a scanner filename.
    """
    raw = _text(name)
    if not raw:
        return "", set()
    aliases: set[str] = set()
    for inner in re.findall(r"[（(]([^)）]+)[)）]", raw):
        key = _alias_key(inner)
        if key:
            aliases.add(key)
    primary = re.sub(r"[（(][^)）]*[)）]", "", raw).strip()
    primary = re.sub(r"\s+", "", primary)
    key = _alias_key(primary)
    if key:
        aliases.add(key)
    return primary, aliases


def normalize_city(value: Any) -> str:
    text = _text(value)
    if not text:
        return ""
    first = re.split(r"[/／]", text)[0]
    first = re.sub(r"(特别行政区|自治州|地区|市|区|县)$", "", first.strip())
    return first


@dataclass
class ClaimLine:
    type_name: str
    category: str
    amount: Optional[float]
    currency: str
    invoice_status: str
    period_start: Optional[date]
    period_end: Optional[date]
    consume_date: Optional[date]
    city: str
    reason: str
    raw: dict = field(default_factory=dict)

    @property
    def has_invoice(self) -> bool:
        return "已有发票" in self.invoice_status or "有发票" in self.invoice_status


@dataclass
class Claim:
    claim_no: str
    person: str
    person_aliases: set[str]
    project: str
    title: str
    submitted_date: Optional[date]
    start: Optional[date]
    end: Optional[date]
    date_source: str
    lines: list[ClaimLine]

    @property
    def category_totals(self) -> dict[str, float]:
        totals: dict[str, float] = {}
        for line in self.lines:
            if line.amount is None:
                continue
            totals[line.category] = round(totals.get(line.category, 0.0) + line.amount, 2)
        return totals

    def lines_for(self, category: str) -> list[ClaimLine]:
        return [line for line in self.lines if line.category == category]

    @property
    def cities(self) -> set[str]:
        return {line.city for line in self.lines if line.city}

    def as_dict(self) -> dict:
        return {
            "claim_no": self.claim_no,
            "person": self.person,
            "project": self.project,
            "title": self.title,
            "submitted_date": self.submitted_date.isoformat() if self.submitted_date else None,
            "start": self.start.isoformat() if self.start else None,
            "end": self.end.isoformat() if self.end else None,
            "date_source": self.date_source,
            "category_totals": self.category_totals,
            "line_count": len(self.lines),
            "cities": sorted(self.cities),
        }


@dataclass
class ClaimIndex:
    source_path: str
    sheet: str
    claims: list[Claim]
    warnings: list[str] = field(default_factory=list)

    def by_person(self, person: str) -> list[Claim]:
        _, aliases = normalize_person(person)
        if not aliases:
            return []
        matched = []
        for claim in self.claims:
            if aliases & claim.person_aliases:
                matched.append(claim)
        return matched

    def as_dict(self) -> dict:
        return {
            "source_path": self.source_path,
            "sheet": self.sheet,
            "claim_count": len(self.claims),
            "claims": [claim.as_dict() for claim in self.claims],
            "warnings": list(self.warnings),
        }


class ClaimWorkbookError(ValueError):
    """Raised when a claim workbook cannot provide a usable index."""


def _find_header_row(rows: list[tuple[Any, ...]]) -> int:
    for index, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        headers = {_text(value) for value in row}
        if "单号" in headers and any("费用金额" in header for header in headers):
            return index
    raise ClaimWorkbookError("报销明细缺少表头行（需包含“单号”和“费用金额”列）")


def _header_positions(header: Iterable[Any]) -> dict[str, list[int]]:
    positions: dict[str, list[int]] = {}
    for index, value in enumerate(header):
        name = _text(value)
        if not name:
            continue
        positions.setdefault(name, []).append(index)
    return positions


def _first_value(row: tuple[Any, ...], positions: dict[str, list[int]], *names: str) -> Any:
    for name in names:
        for index in positions.get(name, []):
            if index < len(row):
                value = row[index]
                if value not in (None, ""):
                    return value
    return None


def _amount_value(row: tuple[Any, ...], positions: dict[str, list[int]]) -> Optional[float]:
    for name in ("费用金额：本位币金额", "费用金额：外币金额"):
        for index in positions.get(name, []):
            if index < len(row):
                parsed = _number(row[index])
                if parsed is not None:
                    return parsed
    for name, indices in positions.items():
        if name.startswith("费用金额") and "本位币" in name:
            for index in indices:
                if index < len(row):
                    parsed = _number(row[index])
                    if parsed is not None:
                        return parsed
    return None


def load_claim_index(path: str, sheet: Optional[str] = None) -> ClaimIndex:
    """Read ``path`` and build a claim index without modifying the workbook."""
    source = Path(path).expanduser().resolve()
    if not source.is_file():
        raise ClaimWorkbookError(f"报销明细不存在: {source}")

    workbook = load_workbook(source, data_only=True, read_only=True)
    try:
        sheet_name = sheet or DEFAULT_SHEET
        if sheet_name not in workbook.sheetnames:
            candidates = [
                name for name in workbook.sheetnames
                if "reimbursement" in name.lower() or "报销" in name
            ]
            if not candidates:
                raise ClaimWorkbookError(
                    f"找不到报销明细工作表“{sheet_name}”，现有工作表: {workbook.sheetnames}"
                )
            sheet_name = candidates[0]
        worksheet = workbook[sheet_name]
        rows = [tuple(row) for row in worksheet.iter_rows(values_only=True)]
    finally:
        workbook.close()

    if not rows:
        raise ClaimWorkbookError("报销明细工作表为空")
    header_index = _find_header_row(rows)
    positions = _header_positions(rows[header_index])
    if "单号" not in positions:
        raise ClaimWorkbookError("报销明细缺少“单号”列")

    warnings: list[str] = []
    grouped: dict[str, list[dict]] = {}
    for row in rows[header_index + 1:]:
        if not row:
            continue
        claim_no = _text(_first_value(row, positions, "单号"))
        if not claim_no:
            continue
        grouped.setdefault(claim_no, []).append({
            "person": _text(_first_value(row, positions, "制单人Name")),
            "project": _text(_first_value(row, positions, "项目Name")),
            "title": _text(_first_value(row, positions, "标题")),
            "submitted": _first_value(row, positions, "提交日期"),
            "period": _first_value(row, positions, "消费起止日期"),
            "type_name": _text(_first_value(row, positions, "费用类型Name")),
            "amount": _amount_value(row, positions),
            "currency": _text(_first_value(row, positions, "费用金额：币种")),
            "invoice_status": _text(_first_value(row, positions, "发票")),
            "consume_date": _first_value(row, positions, "消费日期"),
            "city": normalize_city(_first_value(row, positions, "消费城市Name", "目的地Name")),
            "reason": _text(_first_value(row, positions, "消费事由")),
        })

    claims: list[Claim] = []
    for claim_no, entries in grouped.items():
        first = entries[0]
        lines: list[ClaimLine] = []
        row_starts: list[date] = []
        row_ends: list[date] = []
        for entry in entries:
            start, end = parse_period(entry["period"])
            if start and end:
                row_starts.append(start)
                row_ends.append(end)
            consume = parse_date(entry["consume_date"])
            lines.append(ClaimLine(
                type_name=entry["type_name"],
                category=normalize_category(entry["type_name"]),
                amount=entry["amount"],
                currency=entry["currency"],
                invoice_status=entry["invoice_status"],
                period_start=start,
                period_end=end,
                consume_date=consume,
                city=entry["city"],
                reason=entry["reason"],
            ))

        submitted = parse_date(first["submitted"])
        if row_starts:
            start, end, date_source = min(row_starts), max(row_ends), "CLAIM_PERIOD"
        else:
            consume_dates = [line.consume_date for line in lines if line.consume_date]
            if consume_dates:
                start, end, date_source = min(consume_dates), max(consume_dates), "CLAIM_CONSUME_DATES"
            else:
                title_start, title_end = parse_period(first["title"])
                if title_start:
                    start, end, date_source = title_start, title_end or title_start, "CLAIM_TITLE"
                elif submitted:
                    start = end = submitted
                    date_source = "DATE_INFERRED_FROM_CLAIM"
                else:
                    start = end = None
                    date_source = "MISSING"
                    warnings.append(f"{claim_no} 无法推断消费起止日期")

        person, aliases = normalize_person(first["person"])
        if not person:
            warnings.append(f"{claim_no} 缺少制单人")
        claims.append(Claim(
            claim_no=claim_no,
            person=person,
            person_aliases=aliases,
            project=first["project"],
            title=first["title"],
            submitted_date=submitted,
            start=start,
            end=end,
            date_source=date_source,
            lines=lines,
        ))

    claims.sort(key=lambda claim: (
        claim.start or date.max,
        claim.claim_no,
    ))
    return ClaimIndex(
        source_path=str(source),
        sheet=sheet_name,
        claims=claims,
        warnings=warnings,
    )


def reconcile_claim_categories(index: ClaimIndex) -> dict[str, dict]:
    """Return configured/observed category totals for a quick structural check."""
    summary: dict[str, dict] = {}
    for claim in index.claims:
        for category, amount in claim.category_totals.items():
            bucket = summary.setdefault(category, {"amount": 0.0, "claim_count": 0})
            bucket["amount"] = round(bucket["amount"] + amount, 2)
            bucket["claim_count"] += 1
    return summary
