"""Field-level evidence and conservative extraction completeness checks.

The validators deliberately report problems instead of changing parsed values.  This
keeps the existing candidate/review workflow authoritative while making it possible
to explain why a scan was accepted, warned, or blocked.
"""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field
from typing import Any, Iterable

from .models import Evidence


@dataclass(frozen=True)
class ValidationIssue:
    code: str
    message: str
    severity: str = "WARNING"
    field: str = ""


@dataclass
class ValidationReport:
    status: str = "PASS"
    issues: list[ValidationIssue] = field(default_factory=list)
    checks: dict[str, Any] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "issues": [asdict(issue) for issue in self.issues],
            "checks": dict(self.checks),
        }


def _finish(issues: list[ValidationIssue], checks: dict[str, Any]) -> ValidationReport:
    if any(issue.severity == "BLOCKING" for issue in issues):
        status = "BLOCKED"
    elif issues:
        status = "WARNING"
    else:
        status = "PASS"
    return ValidationReport(status=status, issues=issues, checks=checks)


def validate_timesheet_fields(
    fields: dict[str, Any] | None,
    *,
    source_path: str = "",
    expected_period: str = "",
    known_po_ids: Iterable[str] = (),
    total_hours: float | None = None,
    daily_hours: Iterable[float] | None = None,
) -> ValidationReport:
    """Validate the four fields used by the Timesheet candidate schema.

    Missing critical fields are blocking; format and source mismatches are warnings
    so the existing review queue can continue to decide whether a candidate is usable.
    """
    values = fields or {}
    issues: list[ValidationIssue] = []
    checks: dict[str, Any] = {}
    required = {"period": "period", "po_id": "PO", "hours": "hours", "name": "name"}
    for key, label in required.items():
        present = bool(str(values.get(key, "")).strip())
        checks[f"{key}_present"] = present
        if not present:
            issues.append(ValidationIssue(
                f"MISSING_{key.upper()}", f"Missing {label} field", "BLOCKING", key,
            ))

    period = str(values.get("period", "")).strip()
    if period and not re.fullmatch(r"20\d{2}-(0[1-9]|1[0-2])", period):
        issues.append(ValidationIssue("INVALID_PERIOD", f"Invalid period: {period}", "WARNING", "period"))
    if expected_period and period and period != expected_period:
        issues.append(ValidationIssue(
            "PERIOD_MISMATCH", f"Period {period} differs from expected {expected_period}", "WARNING", "period",
        ))

    po_id = re.sub(r"\D", "", str(values.get("po_id", "")))
    known = {re.sub(r"\D", "", str(item)) for item in known_po_ids if str(item).strip()}
    checks["po_digits"] = po_id
    if po_id and len(po_id) < 8:
        issues.append(ValidationIssue("INVALID_PO", "PO number is shorter than eight digits", "WARNING", "po_id"))
    if known and po_id and po_id not in known:
        issues.append(ValidationIssue("UNKNOWN_PO", f"PO {po_id} is not in the project PO list", "WARNING", "po_id"))

    raw_hours = total_hours if total_hours is not None else values.get("hours")
    try:
        numeric_hours = float(raw_hours) if raw_hours not in (None, "") else None
    except (TypeError, ValueError):
        numeric_hours = None
    checks["hours"] = numeric_hours
    if numeric_hours is not None and not 0 <= numeric_hours <= 400:
        issues.append(ValidationIssue("INVALID_HOURS", f"Hours outside 0-400 range: {numeric_hours}", "BLOCKING", "hours"))
    if daily_hours is not None and numeric_hours is not None:
        try:
            daily_total = round(sum(float(value) for value in daily_hours), 2)
        except (TypeError, ValueError):
            daily_total = None
        checks["daily_hours_total"] = daily_total
        if daily_total is None:
            issues.append(ValidationIssue("INVALID_DAILY_HOURS", "Daily hours contain a non-numeric value", "BLOCKING", "hours"))
        elif abs(daily_total - numeric_hours) > 0.01:
            issues.append(ValidationIssue(
                "DAILY_TOTAL_MISMATCH",
                f"Daily hours total {daily_total} differs from monthly total {numeric_hours}",
                "BLOCKING",
                "hours",
            ))

    if source_path and period:
        basename = source_path.replace("\\", "/").rsplit("/", 1)[-1]
        checks["period_in_source_name"] = period.replace("-", "") in basename

    return _finish(issues, checks)


def validate_ocr_completeness(
    text: str,
    *,
    required_markers: Iterable[str] = (),
    minimum_characters: int = 20,
) -> ValidationReport:
    """Check that OCR returned enough usable content for downstream field parsing."""
    compact = re.sub(r"\s+", "", text or "")
    issues: list[ValidationIssue] = []
    checks: dict[str, Any] = {
        "character_count": len(compact),
        "digit_count": sum(char.isdigit() for char in compact),
    }
    if len(compact) < minimum_characters:
        issues.append(ValidationIssue("OCR_TOO_SHORT", "OCR output is too short to be complete", "WARNING"))
    for marker in required_markers:
        found = marker.casefold() in (text or "").casefold()
        checks[f"marker:{marker}"] = found
        if not found:
            issues.append(ValidationIssue("OCR_MARKER_MISSING", f"OCR marker not found: {marker}", "WARNING"))
    return _finish(issues, checks)


def _line_for_value(text: str, value: Any) -> str:
    target = str(value or "").strip()
    if not target:
        return ""
    for line in (text or "").splitlines():
        if target.casefold() in line.casefold():
            return re.sub(r"\s+", " ", line).strip()[:500]
    return ""


def _bbox_for_value(value: Any, blocks: Iterable[dict[str, Any]]) -> tuple[float, float, float, float] | None:
    target = str(value or "").casefold().strip()
    if not target:
        return None
    block_list = list(blocks)
    is_period = bool(re.fullmatch(r"20\d{2}-(0[1-9]|1[0-2])", target))
    for block in block_list:
        text = str(block.get("text", "")).casefold()
        if target in text or (text in target and not is_period):
            left, top = block.get("left"), block.get("top")
            width, height = block.get("width"), block.get("height")
            if all(value is not None for value in (left, top, width, height)):
                return (float(left), float(top), float(left + width), float(top + height))

    # OCR frequently emits a monthly period as two neighbouring cells (``08``
    # and ``2026``), while the business field is normalised to ``2026-08``.
    period = re.fullmatch(r"(20\d{2})-(0[1-9]|1[0-2])", target)
    if period:
        year, month = period.groups()
        month_blocks = [item for item in block_list if str(item.get("text", "")).strip() == month]
        year_blocks = [item for item in block_list if str(item.get("text", "")).strip() == year]
        for month_block in month_blocks:
            for year_block in year_blocks:
                month_y = float(month_block.get("top", 0))
                year_y = float(year_block.get("top", 0))
                month_h = float(month_block.get("height", 0))
                year_h = float(year_block.get("height", 0))
                if abs(month_y - year_y) <= max(month_h, year_h, 1) * 2:
                    values = (month_block, year_block)
                    if all(all(key in item and item[key] is not None for key in ("left", "top", "width", "height")) for item in values):
                        left = min(float(item["left"]) for item in values)
                        top = min(float(item["top"]) for item in values)
                        right = max(float(item["left"]) + float(item["width"]) for item in values)
                        bottom = max(float(item["top"]) + float(item["height"]) for item in values)
                        return (left, top, right, bottom)

    # Numeric OCR tokens may omit a decimal point.  Compare numeric values only
    # after the exact text search above, and keep the first positioned match.
    try:
        numeric_target = float(target.replace(",", "."))
    except ValueError:
        numeric_target = None
    if numeric_target is not None:
        for block in block_list:
            token = str(block.get("text", "")).strip().replace(",", ".")
            if not re.fullmatch(r"\d{1,5}(?:\.\d+)?", token):
                continue
            try:
                if abs(float(token) - numeric_target) > 0.01:
                    continue
            except ValueError:
                continue
            if all(block.get(key) is not None for key in ("left", "top", "width", "height")):
                return (
                    float(block["left"]), float(block["top"]),
                    float(block["left"]) + float(block["width"]),
                    float(block["top"]) + float(block["height"]),
                )
    return None


def build_timesheet_evidence(
    fields: dict[str, Any] | None,
    *,
    source_file: str,
    text: str = "",
    blocks: Iterable[dict[str, Any]] = (),
    field_blocks: dict[str, Iterable[dict[str, Any]]] | None = None,
    page: int = 1,
    ocr_engine: str = "",
    ocr_confidence: float | None = None,
) -> list[dict[str, Any]]:
    """Create serialisable evidence records without altering field values.

    ``field_blocks`` lets a layout-aware extractor provide a more specific
    source region for one field, such as the dedicated monthly-total footer.
    The general OCR blocks remain the fallback for every field.
    """
    values = fields or {}
    default_blocks = list(blocks)
    preferred_blocks = field_blocks or {}
    result: list[dict[str, Any]] = []
    for field_name in ("name", "period", "po_id", "hours"):
        value = values.get(field_name, "")
        if value in (None, ""):
            continue
        source_blocks = preferred_blocks.get(field_name, default_blocks)
        result.append(Evidence(
            field=field_name,
            value=value,
            source_file=source_file,
            page=page,
            bbox=_bbox_for_value(value, source_blocks),
            raw_text=_line_for_value(text, value),
            ocr_engine=ocr_engine,
            ocr_confidence=ocr_confidence,
            validation_status="PENDING",
        ).as_dict())
    return result
