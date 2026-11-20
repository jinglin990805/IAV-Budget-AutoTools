"""Reconcile parsed travel trips against an expense-claim export.

The claim export is treated as a baseline index.  It answers three questions:

* Which claim does this trip belong to?
* How much did the claim declare for each category?
* Where does the PDF evidence disagree with the declared amount?

Every invoice, including one that cannot be matched to a claim line, becomes a
row in the reconciliation workbook.  Missing amounts become rows too, so the
Excel file is the single place a reviewer needs to look.  The module never
changes a parsed expense and never approves a payment.
"""

from __future__ import annotations

import itertools
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Iterable, Optional

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .reimbursement_claim import (
    CATEGORY_ACCOMMODATION,
    CATEGORY_ALLOWANCE,
    CATEGORY_OTHER,
    CATEGORY_TICKET,
    CATEGORY_TRAFFIC,
    CLAIM_CATEGORY_LABELS,
    EXPENSE_CATEGORY_BUCKETS,
    Claim,
    ClaimIndex,
    normalize_city,
    normalize_person,
)


MATCH_BUFFER_DAYS = 7
AMOUNT_TOLERANCE = 0.01
INVOICE_CATEGORY_LABELS = {
    "accommodation": "住宿",
    "ticket": "机票/火车",
    "taxi": "打车/市内交通",
    "special": "打车/市内交通",
}
STATUS_LABELS = {
    "PASS": "一致",
    "MISSING": "缺失",
    "EXCESS": "超出",
    "AMOUNT_DIFF": "金额差异",
    "TYPE_UNCERTAIN": "类型待确认",
    "CANDIDATE": "待确认候选",
    "NO_CLAIM": "无报销单",
    "DATE_OUTSIDE": "日期越界",
}
REVIEW_STATUSES = {
    "MISSING", "EXCESS", "AMOUNT_DIFF", "TYPE_UNCERTAIN",
    "CANDIDATE", "NO_CLAIM", "DATE_OUTSIDE",
}


@dataclass
class ReconciliationRow:
    sequence: int
    source_page: Optional[int]
    service_date: Optional[date]
    invoice_type: str
    amount: Optional[float]
    claim_category: str
    claim_no: str
    expected_amount: Optional[float]
    status: str
    difference: Optional[float]
    basis: str
    risk: str
    manual_amount: Optional[float] = None
    manual_reason: str = ""
    adjustment: str = ""
    description: str = ""
    document_id: str = ""

    def as_dict(self) -> dict:
        return {
            "sequence": self.sequence,
            "source_page": self.source_page,
            "service_date": self.service_date.isoformat() if self.service_date else None,
            "invoice_type": self.invoice_type,
            "amount": self.amount,
            "claim_category": self.claim_category,
            "claim_no": self.claim_no,
            "expected_amount": self.expected_amount,
            "status": self.status,
            "status_label": STATUS_LABELS.get(self.status, self.status),
            "difference": self.difference,
            "basis": self.basis,
            "risk": self.risk,
            "manual_amount": self.manual_amount,
            "manual_reason": self.manual_reason,
            "adjustment": self.adjustment,
            "description": self.description,
            "document_id": self.document_id,
        }


@dataclass
class BucketResult:
    category: str
    expected: float
    matched: float
    difference: float
    status: str
    invoice_count: int
    notes: list[str] = field(default_factory=list)

    @property
    def label(self) -> str:
        return CLAIM_CATEGORY_LABELS.get(self.category, self.category)

    def as_dict(self) -> dict:
        return {
            "category": self.category,
            "label": self.label,
            "expected": round(self.expected, 2),
            "matched": round(self.matched, 2),
            "difference": round(self.difference, 2),
            "status": self.status,
            "invoice_count": self.invoice_count,
            "notes": list(self.notes),
        }


@dataclass
class TripReconciliation:
    trip_name: str
    trip_city: str
    trip_start: Optional[date]
    trip_end: Optional[date]
    source_file: str
    claim: Optional[Claim]
    match_status: str
    match_basis: str
    match_score: int
    rows: list[ReconciliationRow]
    buckets: list[BucketResult]

    @property
    def status(self) -> str:
        if self.claim is None:
            return "NO_CLAIM"
        if any(bucket.status in REVIEW_STATUSES for bucket in self.buckets):
            return "REVIEW"
        if any(row.status in REVIEW_STATUSES for row in self.rows):
            return "REVIEW"
        return "PASS"

    @property
    def exception_count(self) -> int:
        return sum(1 for row in self.rows if row.status in REVIEW_STATUSES)

    @property
    def invoice_total(self) -> float:
        return round(sum(row.amount or 0.0 for row in self.rows if row.status != "MISSING"), 2)

    @property
    def claim_total(self) -> float:
        if self.claim is None:
            return 0.0
        return round(sum(self.claim.category_totals.values()), 2)

    def as_dict(self) -> dict:
        return {
            "trip": {
                "name": self.trip_name,
                "city": self.trip_city,
                "start": self.trip_start.isoformat() if self.trip_start else None,
                "end": self.trip_end.isoformat() if self.trip_end else None,
                "source_file": self.source_file,
            },
            "claim": self.claim.as_dict() if self.claim else None,
            "match_status": self.match_status,
            "match_basis": self.match_basis,
            "match_score": self.match_score,
            "status": self.status,
            "exception_count": self.exception_count,
            "invoice_total": self.invoice_total,
            "claim_total": self.claim_total,
            "buckets": [bucket.as_dict() for bucket in self.buckets],
            "rows": [row.as_dict() for row in self.rows],
        }


# ---------------------------------------------------------------------------
# Claim matching


def _date_overlap_score(trip_start, trip_end, claim: Claim) -> int:
    if not trip_start or not trip_end or not claim.start or not claim.end:
        return 0
    buffer = timedelta(days=MATCH_BUFFER_DAYS)
    left = claim.start - buffer
    right = claim.end + buffer
    if trip_end < left or trip_start > right:
        return 0
    overlap_start = max(trip_start, claim.start)
    overlap_end = min(trip_end, claim.end)
    if overlap_start <= overlap_end:
        return 40
    return 20  # inside the buffer but not overlapping


def _amount_score(trip, claim: Claim) -> int:
    expected = claim.category_totals
    matched = 0.0
    for expense in getattr(trip, "expenses", []):
        bucket = EXPENSE_CATEGORY_BUCKETS.get(getattr(expense, "category", ""))
        if bucket in (CATEGORY_TICKET, CATEGORY_TRAFFIC, CATEGORY_ACCOMMODATION):
            matched += float(getattr(expense, "amount", 0.0) or 0.0)
    claim_amount = sum(expected.get(bucket, 0.0) for bucket in (
        CATEGORY_TICKET, CATEGORY_TRAFFIC, CATEGORY_ACCOMMODATION
    ))
    if claim_amount <= 0:
        return 0
    delta = abs(claim_amount - matched)
    if delta <= AMOUNT_TOLERANCE:
        return 30
    ratio = delta / claim_amount
    if ratio <= 0.02:
        return 25
    if ratio <= 0.10:
        return 15
    if ratio <= 0.25:
        return 5
    return 0


def _trip_people(trip) -> set[str]:
    names: set[str] = set()
    for value in (getattr(trip, "name", ""),):
        _, aliases = normalize_person(value)
        names |= aliases
    return names


def _person_score(trip, claim: Claim) -> tuple[int, str]:
    aliases = _trip_people(trip)
    if not aliases:
        return 0, ""
    if aliases & claim.person_aliases:
        return 60, f"人员一致（{claim.person}）"
    # A filename may include both the Chinese name and the pinyin alias.
    compact = "".join(aliases)
    for alias in claim.person_aliases:
        if alias and alias in compact:
            return 55, f"人员别名一致（{claim.person}）"
    return 0, ""


def _city_score(trip, claim: Claim) -> tuple[int, str]:
    trip_city = normalize_city(getattr(trip, "city", ""))
    if not trip_city or not claim.cities:
        return 0, ""
    if trip_city in claim.cities:
        return 20, f"城市一致（{trip_city}）"
    return 0, ""


def match_claim(trip, index: ClaimIndex) -> tuple[Optional[Claim], str, int, str]:
    """Choose the claim that best matches ``trip``.

    Returns ``(claim, match_status, score, basis)``.  A person match is the
    normal path; date + city + amount can still match an export whose name
    column uses a different spelling, but the result is marked as inferred.
    """
    best: tuple[int, Optional[Claim], list[str], bool] = (0, None, [], False)
    for claim in index.claims:
        person_score, person_note = _person_score(trip, claim)
        date_score = _date_overlap_score(
            getattr(trip, "start", None), getattr(trip, "end", None), claim
        )
        city_score, city_note = _city_score(trip, claim)
        amount_score = _amount_score(trip, claim)
        score = person_score + date_score + city_score + amount_score
        notes = [note for note in (person_note, city_note) if note]
        if date_score:
            notes.append("日期区间匹配")
        if amount_score:
            notes.append("类目金额接近")
        if score > best[0]:
            best = (score, claim, notes, bool(person_score))
    score, claim, notes, person_matched = best
    if claim is None:
        return None, "NO_CLAIM", 0, "报销明细中没有任何可用单号"
    if person_matched and score >= 60:
        return claim, "MATCHED", score, "；".join(notes)
    # Without a person match, require date + city + amount evidence.  Date and
    # city alone are too weak: another traveller can visit the same city on
    # overlapping dates with a completely different claim.
    if not person_matched and score >= 75 and amount_score >= 15:
        return claim, "MATCH_INFERRED", score, "；".join(notes) + "；人员未直接匹配，按日期/城市/金额推断"
    return None, "NO_CLAIM", score, "无法可靠匹配到报销单"


# ---------------------------------------------------------------------------
# Difference explanation helpers


def _find_subset(values: list[float], target: float, max_items: int = 4) -> Optional[list[int]]:
    """Return indices whose sum equals ``target`` within tolerance."""
    if target <= AMOUNT_TOLERANCE:
        return []
    order = sorted(range(len(values)), key=lambda index: abs(values[index] - target))
    if len(order) > 24:
        order = order[:24]
    for size in range(1, min(max_items, len(order)) + 1):
        for combination in itertools.combinations(order, size):
            if abs(sum(values[index] for index in combination) - target) <= AMOUNT_TOLERANCE:
                return list(combination)
    return None


def _row_from_expense(
    expense,
    sequence: int,
    claim: Claim,
    category: str,
    status: str,
    expected: Optional[float],
    difference: Optional[float],
    basis: str,
    risk: str,
) -> ReconciliationRow:
    service_date = getattr(expense, "start", None) or getattr(expense, "end", None)
    category_name = getattr(expense, "category", "")
    return ReconciliationRow(
        sequence=sequence,
        source_page=getattr(expense, "source_page", None),
        service_date=service_date,
        invoice_type=INVOICE_CATEGORY_LABELS.get(category_name, category_name or "未确认"),
        amount=round(float(getattr(expense, "amount", 0.0) or 0.0), 2),
        claim_category=CLAIM_CATEGORY_LABELS.get(category, category),
        claim_no=claim.claim_no,
        expected_amount=round(expected, 2) if expected is not None else None,
        status=status,
        difference=round(difference, 2) if difference is not None else None,
        basis=basis,
        risk=risk,
        description=getattr(expense, "description", "") or "",
        document_id=getattr(expense, "document_id", "") or "",
    )


def _date_outside(expense, claim: Claim) -> bool:
    service = getattr(expense, "start", None) or getattr(expense, "end", None)
    if not service or not claim.start or not claim.end:
        return False
    buffer = timedelta(days=MATCH_BUFFER_DAYS)
    return not (claim.start - buffer <= service <= claim.end + buffer)


def _expense_rows_for_aggregate(
    expenses: list,
    claim: Claim,
    category: str,
    expected: float,
    matched: float,
) -> tuple[list[ReconciliationRow], list[str]]:
    """Explain one aggregate claim line for a category."""
    rows: list[ReconciliationRow] = []
    notes: list[str] = []
    difference = round(expected - matched, 2)

    if abs(difference) <= AMOUNT_TOLERANCE:
        for expense in expenses:
            rows.append(_row_from_expense(
                expense, 0, claim, category, "PASS", expected, 0.0,
                "类目合计一致", "",
            ))
        return rows, notes

    if category == CATEGORY_ACCOMMODATION:
        for expense in expenses:
            risk = (
                f"住宿报销 ¥{expected:.2f} 与票面 ¥{matched:.2f} 不同；"
                "按差标口径不判金额异常，核对晚数/水单即可"
            )
            rows.append(_row_from_expense(
                expense, 0, claim, category, "PASS", expected, difference,
                "住宿仅核对票据与晚数", risk,
            ))
        notes.append("住宿金额按差标口径不参与金额异常判定")
        return rows, notes

    if len(expenses) == 1:
        expense = expenses[0]
        status = "DATE_OUTSIDE" if _date_outside(expense, claim) else "AMOUNT_DIFF"
        risk = (
            f"票面 ¥{getattr(expense, 'amount', 0.0):.2f}，报销单 ¥{expected:.2f}，"
            f"差 {difference:+.2f}；请核对票面实收/合计金额"
        )
        if _date_outside(expense, claim):
            risk += "；票据日期超出行程/报销区间"
        rows.append(_row_from_expense(
            expense, 0, claim, category, status, expected, difference,
            "单票对单类目", risk,
        ))
        return rows, notes

    missing = difference > AMOUNT_TOLERANCE
    if missing:
        for expense in expenses:
            risk = "" if not _date_outside(expense, claim) else "票据日期超出行程/报销区间"
            status = "DATE_OUTSIDE" if risk else "PASS"
            rows.append(_row_from_expense(
                expense, 0, claim, category, status, expected, None,
                "类目合计的一部分", risk,
            ))
        rows.append(ReconciliationRow(
            sequence=0,
            source_page=None,
            service_date=claim.start,
            invoice_type=f"缺失-{CLAIM_CATEGORY_LABELS.get(category, category)}",
            amount=None,
            claim_category=CLAIM_CATEGORY_LABELS.get(category, category),
            claim_no=claim.claim_no,
            expected_amount=round(expected, 2),
            status="MISSING",
            difference=round(difference, 2),
            basis="类目合计差额",
            risk=(
                f"报销单该类目 ¥{expected:.2f}，PDF 已识别 ¥{matched:.2f}，"
                f"缺 ¥{difference:.2f}；请在丢失票据中补录或确认是否含 PDF 外票据"
            ),
        ))
        notes.append(f"缺失 ¥{difference:.2f}")
        return rows, notes

    # PDF total exceeds the claim: point at the smallest subset that explains it.
    excess = abs(difference)
    amounts = [float(getattr(expense, "amount", 0.0) or 0.0) for expense in expenses]
    subset = _find_subset(amounts, excess)
    flagged: set[int] = set()
    if subset is not None and len(subset) < len(expenses):
        flagged = set(subset)
    for index, expense in enumerate(expenses):
        if index in flagged:
            rows.append(_row_from_expense(
                expense, 0, claim, category, "EXCESS",
                expected, -float(getattr(expense, "amount", 0.0) or 0.0),
                "金额与差额一致，疑似超出",
                (
                    f"报销单该类目 ¥{expected:.2f}，PDF 识别 ¥{matched:.2f}，"
                    f"超出 ¥{excess:.2f}；该票可能重复、不属于本单或类别错误"
                ),
            ))
        else:
            risk = "" if not _date_outside(expense, claim) else "票据日期超出行程/报销区间"
            status = "DATE_OUTSIDE" if risk else "PASS"
            rows.append(_row_from_expense(
                expense, 0, claim, category, status, expected, None,
                "类目合计的一部分", risk,
            ))
    if not flagged:
        rows.append(ReconciliationRow(
            sequence=0,
            source_page=None,
            service_date=claim.start,
            invoice_type=f"超出-{CLAIM_CATEGORY_LABELS.get(category, category)}",
            amount=None,
            claim_category=CLAIM_CATEGORY_LABELS.get(category, category),
            claim_no=claim.claim_no,
            expected_amount=round(expected, 2),
            status="EXCESS",
            difference=round(difference, 2),
            basis="类目合计差额",
            risk=(
                f"报销单该类目 ¥{expected:.2f}，PDF 识别 ¥{matched:.2f}，"
                f"超出 ¥{excess:.2f}；未能在单票层面定位，请核对是否重复或跨单"
            ),
        ))
    notes.append(f"超出 ¥{excess:.2f}")
    return rows, notes


def _expense_rows_for_lines(
    expenses: list,
    claim_lines: list,
    claim: Claim,
    category: str,
) -> tuple[list[ReconciliationRow], list[str]]:
    """Reconcile multiple per-ticket claim lines against parsed expenses."""
    rows: list[ReconciliationRow] = []
    notes: list[str] = []
    remaining_expenses = list(expenses)
    remaining_lines = [line for line in claim_lines if line.amount is not None]
    matched_line_ids: set[int] = set()

    # Pass 1: one-to-one exact amount matches.
    for line in sorted(remaining_lines, key=lambda item: item.amount or 0.0, reverse=True):
        candidates = [
            expense for expense in remaining_expenses
            if abs(float(getattr(expense, "amount", 0.0) or 0.0) - float(line.amount)) <= AMOUNT_TOLERANCE
        ]
        if candidates:
            expense = candidates[0]
            remaining_expenses.remove(expense)
            matched_line_ids.add(id(line))
            risk = "" if not _date_outside(expense, claim) else "票据日期超出行程/报销区间"
            status = "DATE_OUTSIDE" if risk else "PASS"
            rows.append(_row_from_expense(
                expense, 0, claim, category, status, float(line.amount), 0.0,
                f"按行金额匹配（{line.reason or '报销明细'}）", risk,
            ))
    remaining_lines = [
        line for line in remaining_lines if id(line) not in matched_line_ids
    ]

    # Pass 2: one claim line may cover a small group of invoices.
    for line in list(remaining_lines):
        amounts = [float(getattr(expense, "amount", 0.0) or 0.0) for expense in remaining_expenses]
        subset = _find_subset(amounts, float(line.amount), max_items=4)
        if not subset or len(subset) > len(remaining_expenses):
            continue
        covered = [remaining_expenses[index] for index in subset]
        for expense in covered:
            remaining_expenses.remove(expense)
            risk = "" if not _date_outside(expense, claim) else "票据日期超出行程/报销区间"
            status = "DATE_OUTSIDE" if risk else "PASS"
            rows.append(_row_from_expense(
                expense, 0, claim, category, status, float(line.amount), None,
                f"分组合计匹配（{line.reason or '报销明细'}）", risk,
            ))
        matched_line_ids.add(id(line))
    remaining_lines = [
        line for line in remaining_lines if id(line) not in matched_line_ids
    ]

    # Pass 3: remaining claim money may be an aggregate line rather than a
    # per-ticket line.  Compare the remaining totals before declaring each
    # invoice "excess" and each claim line "missing"; otherwise a single
    # aggregate taxi line would make every underlying invoice look wrong.
    expected_remaining = round(
        sum(float(line.amount or 0.0) for line in remaining_lines), 2
    )
    matched_remaining = round(
        sum(float(getattr(expense, "amount", 0.0) or 0.0) for expense in remaining_expenses), 2
    )
    difference = round(expected_remaining - matched_remaining, 2)

    if remaining_lines and remaining_expenses:
        if abs(difference) <= AMOUNT_TOLERANCE:
            for expense in remaining_expenses:
                risk = "" if not _date_outside(expense, claim) else "票据日期超出行程/报销区间"
                status = "DATE_OUTSIDE" if risk else "PASS"
                rows.append(_row_from_expense(
                    expense, 0, claim, category, status, expected_remaining, None,
                    "剩余报销行合计一致", risk,
                ))
            return rows, notes
        if difference > 0:
            for expense in remaining_expenses:
                risk = "" if not _date_outside(expense, claim) else "票据日期超出行程/报销区间"
                status = "DATE_OUTSIDE" if risk else "PASS"
                rows.append(_row_from_expense(
                    expense, 0, claim, category, status, expected_remaining, None,
                    "剩余报销行的一部分", risk,
                ))
            rows.append(ReconciliationRow(
                sequence=0,
                source_page=None,
                service_date=claim.start,
                invoice_type=f"缺失-{CLAIM_CATEGORY_LABELS.get(category, category)}",
                amount=None,
                claim_category=CLAIM_CATEGORY_LABELS.get(category, category),
                claim_no=claim.claim_no,
                expected_amount=expected_remaining,
                status="MISSING",
                difference=difference,
                basis="剩余报销行合计差额",
                risk=(
                    f"报销明细其余该类目合计 ¥{expected_remaining:.2f}，"
                    f"PDF 已识别 ¥{matched_remaining:.2f}，缺 ¥{difference:.2f}；"
                    "请补录缺失票据或确认是否含 PDF 外票据"
                ),
            ))
            notes.append(f"缺失 ¥{difference:.2f}")
            return rows, notes

        # PDF exceeds the remaining claim money.
        excess = abs(difference)
        amounts = [float(getattr(expense, "amount", 0.0) or 0.0) for expense in remaining_expenses]
        subset = _find_subset(amounts, excess)
        flagged: set[int] = set()
        if subset is not None and len(subset) < len(remaining_expenses):
            flagged = set(subset)
        for index, expense in enumerate(remaining_expenses):
            if index in flagged:
                rows.append(_row_from_expense(
                    expense, 0, claim, category, "EXCESS",
                    expected_remaining, -float(getattr(expense, "amount", 0.0) or 0.0),
                    "金额与差额一致，疑似超出",
                    (
                        f"报销明细其余该类目 ¥{expected_remaining:.2f}，"
                        f"PDF 识别 ¥{matched_remaining:.2f}，超出 ¥{excess:.2f}；"
                        "该票可能重复、不属于本单或类别错误"
                    ),
                ))
            else:
                risk = "" if not _date_outside(expense, claim) else "票据日期超出行程/报销区间"
                status = "DATE_OUTSIDE" if risk else "PASS"
                rows.append(_row_from_expense(
                    expense, 0, claim, category, status, expected_remaining, None,
                    "剩余报销行的一部分", risk,
                ))
        if not flagged:
            rows.append(ReconciliationRow(
                sequence=0,
                source_page=None,
                service_date=claim.start,
                invoice_type=f"超出-{CLAIM_CATEGORY_LABELS.get(category, category)}",
                amount=None,
                claim_category=CLAIM_CATEGORY_LABELS.get(category, category),
                claim_no=claim.claim_no,
                expected_amount=expected_remaining,
                status="EXCESS",
                difference=difference,
                basis="剩余报销行合计差额",
                risk=(
                    f"报销明细其余该类目合计 ¥{expected_remaining:.2f}，"
                    f"PDF 识别 ¥{matched_remaining:.2f}，超出 ¥{excess:.2f}；"
                    "未能在单票层面定位，请核对是否重复或跨单"
                ),
            ))
        notes.append(f"超出 ¥{excess:.2f}")
        return rows, notes

    # Only claim lines remain: each is a missing invoice.
    for line in remaining_lines:
        rows.append(ReconciliationRow(
            sequence=0,
            source_page=None,
            service_date=line.consume_date or claim.start,
            invoice_type=f"缺失-{CLAIM_CATEGORY_LABELS.get(category, category)}",
            amount=None,
            claim_category=CLAIM_CATEGORY_LABELS.get(category, category),
            claim_no=claim.claim_no,
            expected_amount=float(line.amount) if line.amount is not None else None,
            status="MISSING",
            difference=float(line.amount) if line.amount is not None else None,
            basis="报销明细逐行差额",
            risk=(
                f"报销明细有 ¥{float(line.amount or 0.0):.2f}（{line.reason or '未注明事由'}），"
                "PDF 未识别到对应票据"
            ),
        ))
        notes.append(f"缺失 ¥{float(line.amount or 0.0):.2f}")

    for expense in remaining_expenses:
        amount = float(getattr(expense, "amount", 0.0) or 0.0)
        rows.append(_row_from_expense(
            expense, 0, claim, category, "EXCESS", None, -amount,
            "未在报销明细逐行中找到对应金额",
            f"该票 ¥{amount:.2f} 在报销明细中没有对应行；可能跨单、重复或类别错误",
        ))
        notes.append(f"超出 ¥{amount:.2f}")
    return rows, notes


def _candidate_rows(
    trip,
    claim: Claim,
    category: str,
    missing_amount: float,
    linked_document_ids: set[str],
    start_sequence: int,
) -> list[ReconciliationRow]:
    """Find unlinked structured invoices whose amount explains a missing amount."""
    rows: list[ReconciliationRow] = []
    if missing_amount <= AMOUNT_TOLERANCE:
        return rows
    for invoice in getattr(trip, "structured_invoices", []) or []:
        document_id = getattr(invoice, "document_id", "") or ""
        if document_id and document_id in linked_document_ids:
            continue
        amount = getattr(invoice, "total_amount", None)
        if amount is None:
            continue
        if abs(float(amount) - missing_amount) > AMOUNT_TOLERANCE:
            continue
        source = getattr(invoice, "source", {}) or {}
        page = source.get("page") if isinstance(source, dict) else None
        expense_type = getattr(invoice, "expense_type", "") or ""
        rows.append(ReconciliationRow(
            sequence=start_sequence + len(rows),
            source_page=page,
            service_date=getattr(invoice, "invoice_date", None),
            invoice_type=f"待确认（票面类别 {expense_type or '未识别'}）",
            amount=round(float(amount), 2),
            claim_category=CLAIM_CATEGORY_LABELS.get(category, category),
            claim_no=claim.claim_no,
            expected_amount=missing_amount,
            status="CANDIDATE",
            difference=0.0,
            basis="金额与类目缺失额一致",
            risk=(
                f"该票金额 ¥{float(amount):.2f} 正好等于该类目缺失额；"
                "请确认票面类别后计入"
            ),
            document_id=document_id,
        ))
    return rows


def _bucket_expenses(trip, category: str) -> list:
    """Return this bucket's expenses in PDF page order."""
    indexed = []
    for index, expense in enumerate(getattr(trip, "expenses", []) or []):
        bucket = EXPENSE_CATEGORY_BUCKETS.get(getattr(expense, "category", ""), CATEGORY_OTHER)
        if bucket == category:
            indexed.append((index, expense))
    indexed.sort(key=lambda pair: _sort_key(pair[1], pair[0]))
    return [expense for _, expense in indexed]


def _sort_key(expense, index: int) -> tuple[int, int]:
    page = getattr(expense, "source_page", None)
    if isinstance(page, int) and page > 0:
        return page, index
    return 10 ** 6, index


def reconcile_trip(trip, index: ClaimIndex) -> TripReconciliation:
    claim, match_status, score, basis = match_claim(trip, index)
    trip_start = getattr(trip, "start", None)
    trip_end = getattr(trip, "end", None)
    trip_name = getattr(trip, "name", "") or ""
    trip_city = getattr(trip, "city", "") or ""
    source_file = getattr(trip, "source_file", "") or ""

    if claim is None:
        rows = []
        for sequence, expense in enumerate(
            sorted(getattr(trip, "expenses", []) or [], key=lambda item: (
                getattr(item, "source_page", None) or 10 ** 6
            )),
            start=1,
        ):
            rows.append(ReconciliationRow(
                sequence=sequence,
                source_page=getattr(expense, "source_page", None),
                service_date=getattr(expense, "start", None) or getattr(expense, "end", None),
                invoice_type=INVOICE_CATEGORY_LABELS.get(
                    getattr(expense, "category", ""), getattr(expense, "category", "") or "未确认"
                ),
                amount=round(float(getattr(expense, "amount", 0.0) or 0.0), 2),
                claim_category="未匹配",
                claim_no="",
                expected_amount=None,
                status="NO_CLAIM",
                difference=None,
                basis=basis,
                risk="未找到对应报销单；该票仍会录入，请在报销明细中确认归属",
                description=getattr(expense, "description", "") or "",
                document_id=getattr(expense, "document_id", "") or "",
            ))
        return TripReconciliation(
            trip_name=trip_name, trip_city=trip_city, trip_start=trip_start,
            trip_end=trip_end, source_file=source_file, claim=None,
            match_status=match_status, match_basis=basis, match_score=score,
            rows=rows, buckets=[],
        )

    rows: list[ReconciliationRow] = []
    buckets: list[BucketResult] = []
    linked_document_ids = {
        getattr(expense, "document_id", "") or "" for expense in getattr(trip, "expenses", []) or []
    }
    for category in (CATEGORY_ACCOMMODATION, CATEGORY_TICKET, CATEGORY_TRAFFIC):
        expenses = _bucket_expenses(trip, category)
        claim_lines = claim.lines_for(category)
        expected = round(sum(line.amount or 0.0 for line in claim_lines), 2)
        matched = round(sum(float(getattr(expense, "amount", 0.0) or 0.0) for expense in expenses), 2)
        difference = round(expected - matched, 2)
        bucket_rows: list[ReconciliationRow] = []
        notes: list[str] = []
        if not claim_lines and not expenses:
            continue
        if claim_lines and not expenses:
            for line in claim_lines:
                bucket_rows.append(ReconciliationRow(
                    sequence=0,
                    source_page=None,
                    service_date=line.consume_date or claim.start,
                    invoice_type=f"缺失-{CLAIM_CATEGORY_LABELS.get(category, category)}",
                    amount=None,
                    claim_category=CLAIM_CATEGORY_LABELS.get(category, category),
                    claim_no=claim.claim_no,
                    expected_amount=line.amount,
                    status="MISSING",
                    difference=line.amount,
                    basis="报销单有金额，PDF 无票据",
                    risk=(
                        f"报销单该类目 ¥{float(line.amount or 0.0):.2f}，"
                        "PDF 中未识别到任何该类票据"
                    ),
                ))
            notes.append(f"缺失 ¥{expected:.2f}")
        elif expenses and not claim_lines:
            for expense in expenses:
                bucket_rows.append(_row_from_expense(
                    expense, 0, claim, category, "EXCESS", 0.0,
                    -float(getattr(expense, "amount", 0.0) or 0.0),
                    "报销单无该类目",
                    f"报销单未申报该类目，但 PDF 识别到 ¥{float(getattr(expense, 'amount', 0.0) or 0.0):.2f}",
                ))
            notes.append(f"超出 ¥{matched:.2f}")
        elif len(claim_lines) > 1:
            bucket_rows, notes = _expense_rows_for_lines(expenses, claim_lines, claim, category)
        else:
            bucket_rows, notes = _expense_rows_for_aggregate(
                expenses, claim, category, expected, matched
            )
        bucket_rows.extend(_candidate_rows(
            trip, claim, category, max(difference, 0.0), linked_document_ids, len(bucket_rows)
        ))
        status = "PASS"
        if any(row.status in REVIEW_STATUSES for row in bucket_rows):
            status = "REVIEW"
        buckets.append(BucketResult(
            category=category,
            expected=expected,
            matched=matched,
            difference=difference,
            status=status,
            invoice_count=len(expenses),
            notes=notes,
        ))
        rows.extend(bucket_rows)

    allowance_lines = claim.lines_for(CATEGORY_ALLOWANCE)
    if allowance_lines:
        allowance_total = round(sum(line.amount or 0.0 for line in allowance_lines), 2)
        buckets.append(BucketResult(
            category=CATEGORY_ALLOWANCE,
            expected=allowance_total,
            matched=0.0,
            difference=allowance_total,
            status="PASS",
            invoice_count=0,
            notes=["补贴无需发票，不参与金额核对"],
        ))

    rows.sort(key=lambda row: (
        row.source_page if isinstance(row.source_page, int) else 10 ** 6,
        row.sequence if row.sequence else 10 ** 6,
    ))
    for sequence, row in enumerate(rows, start=1):
        row.sequence = sequence
    return TripReconciliation(
        trip_name=trip_name, trip_city=trip_city, trip_start=trip_start,
        trip_end=trip_end, source_file=source_file, claim=claim,
        match_status=match_status, match_basis=basis, match_score=score,
        rows=rows, buckets=buckets,
    )


def reconcile_trips(trips: Iterable, index: ClaimIndex) -> list[TripReconciliation]:
    return [reconcile_trip(trip, index) for trip in trips]


# ---------------------------------------------------------------------------
# Workbook output


DETAIL_HEADERS = [
    ("序号", 6),
    ("页码", 7),
    ("日期", 12),
    ("票据类型", 16),
    ("票面金额", 12),
    ("报销类别", 16),
    ("报销单号", 13),
    ("类目报销金额", 13),
    ("匹配状态", 12),
    ("差异金额", 11),
    ("匹配依据", 26),
    ("风险说明", 46),
    ("人工调整金额", 13),
    ("人工调整原因", 22),
    ("调整标记", 17),
]
SUMMARY_HEADERS = [
    ("报销单号", 13),
    ("姓名", 14),
    ("项目", 20),
    ("城市", 10),
    ("行程起", 12),
    ("行程止", 12),
    ("日期来源", 22),
    ("票据张数", 9),
    ("票面合计", 12),
    ("Hotel报销", 12),
    ("Flight&Train报销", 15),
    ("Traffic Fee报销", 14),
    ("补贴报销", 11),
    ("核对状态", 11),
    ("异常数", 8),
    ("匹配依据", 40),
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
BODY_FONT = Font(size=10)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
STATUS_FILLS = {
    "PASS": PatternFill("solid", fgColor="E2EFDA"),
    "MISSING": PatternFill("solid", fgColor="FCE4D6"),
    "EXCESS": PatternFill("solid", fgColor="FCE4D6"),
    "AMOUNT_DIFF": PatternFill("solid", fgColor="FFF2CC"),
    "TYPE_UNCERTAIN": PatternFill("solid", fgColor="FFF2CC"),
    "CANDIDATE": PatternFill("solid", fgColor="FFF2CC"),
    "NO_CLAIM": PatternFill("solid", fgColor="EDEDED"),
    "DATE_OUTSIDE": PatternFill("solid", fgColor="FFF2CC"),
}


def _write_header(worksheet, headers: list[tuple[str, int]], row: int = 1) -> None:
    for column, (title, width) in enumerate(headers, start=1):
        cell = worksheet.cell(row, column, title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        worksheet.column_dimensions[get_column_letter(column)].width = width
    worksheet.row_dimensions[row].height = 26


def add_reconciliation_sheets(
    workbook,
    reconciliations: list[TripReconciliation],
    *,
    claim_source: str = "",
) -> dict:
    """Add/refresh ``报销核对`` and ``核对汇总`` sheets on an open workbook."""
    for sheet_name in ("报销核对", "核对汇总"):
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]
    detail = workbook.create_sheet("报销核对")
    summary = workbook.create_sheet("核对汇总")
    _write_header(detail, DETAIL_HEADERS)
    _write_header(summary, SUMMARY_HEADERS)

    detail_row = 2
    summary_row = 2
    for result in reconciliations:
        for row in result.rows:
            values = [
                row.sequence,
                row.source_page,
                row.service_date,
                row.invoice_type,
                row.amount,
                row.claim_category,
                row.claim_no,
                row.expected_amount,
                STATUS_LABELS.get(row.status, row.status),
                row.difference,
                row.basis,
                row.risk,
                row.manual_amount,
                row.manual_reason,
                row.adjustment,
            ]
            for column, value in enumerate(values, start=1):
                cell = detail.cell(detail_row, column, value)
                cell.font = BODY_FONT
                cell.border = BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=column in {11, 12})
                if column == 3 and value is not None:
                    cell.number_format = "yyyy-mm-dd"
                if column in {5, 8, 10, 13}:
                    cell.number_format = "#,##0.00"
            fill = STATUS_FILLS.get(row.status)
            if fill:
                detail.cell(detail_row, 9).fill = fill
            detail_row += 1

        claim = result.claim
        category_totals = claim.category_totals if claim else {}
        values = [
            claim.claim_no if claim else "",
            claim.person if claim else result.trip_name,
            claim.project if claim else "",
            result.trip_city,
            result.trip_start,
            result.trip_end,
            claim.date_source if claim else "",
            len(result.rows),
            result.invoice_total,
            category_totals.get(CATEGORY_ACCOMMODATION, 0.0),
            category_totals.get(CATEGORY_TICKET, 0.0),
            category_totals.get(CATEGORY_TRAFFIC, 0.0),
            category_totals.get(CATEGORY_ALLOWANCE, 0.0),
            STATUS_LABELS.get(result.status, result.status),
            result.exception_count,
            result.match_basis,
        ]
        for column, value in enumerate(values, start=1):
            cell = summary.cell(summary_row, column, value)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=column == 16)
            if column in {5, 6} and value is not None:
                cell.number_format = "yyyy-mm-dd"
            if column in {9, 10, 11, 12, 13}:
                cell.number_format = "#,##0.00"
        fill = STATUS_FILLS.get(result.status)
        if fill:
            summary.cell(summary_row, 14).fill = fill
        summary_row += 1

    detail.freeze_panes = "A2"
    detail.auto_filter.ref = f"A1:O{detail_row - 1}"
    summary.freeze_panes = "A2"
    summary.auto_filter.ref = f"A1:P{summary_row - 1}"

    if claim_source:
        detail.cell(1, 17, "报销明细来源").font = Font(bold=True, size=9)
        detail.cell(2, 17, claim_source).font = Font(size=9)

    return {
        "workbook": getattr(workbook, "path", ""),
        "detail_sheet": "报销核对",
        "summary_sheet": "核对汇总",
        "detail_rows": detail_row - 2,
        "trip_count": len(reconciliations),
        "status": (
            "REVIEW" if any(item.status != "PASS" for item in reconciliations) else "PASS"
        ),
        "trip_statuses": [
            {
                "claim_no": item.claim.claim_no if item.claim else "",
                "trip": item.trip_name,
                "status": item.status,
                "exception_count": item.exception_count,
            }
            for item in reconciliations
        ],
    }


def write_reconciliation_workbook(
    workbook_path: str,
    reconciliations: list[TripReconciliation],
    *,
    claim_source: str = "",
) -> dict:
    """Write reconciliation sheets into a standalone copy of ``workbook_path``."""
    workbook = load_workbook(workbook_path)
    try:
        result = add_reconciliation_sheets(
            workbook, reconciliations, claim_source=claim_source
        )
        result["workbook"] = workbook_path
        workbook.save(workbook_path)
        return result
    finally:
        workbook.close()
