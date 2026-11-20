# -*- coding: utf-8 -*-
"""Build controlled, read-only Timesheet review batches.

The current-month review flow is intentionally kept separate from historical
review.  This module only classifies persisted candidates and never approves,
dismisses, or mutates a fact.  It gives the Review Center a stable vocabulary
for the historical backlog while preserving the candidate -> review -> fact
boundary.
"""

from __future__ import annotations

import json
from collections import Counter
from typing import Any, Iterable


BATCH_ORDER = (
    "VALIDATION_BLOCKED",
    "PO_MISSING",
    "PO_POTENTIAL",
    "PO_UNKNOWN",
    "VALIDATION_WARNING",
    "READY",
)

BATCH_LABELS = {
    "VALIDATION_BLOCKED": "字段校验阻断",
    "PO_MISSING": "缺少 PO",
    "PO_POTENTIAL": "PO 潜在匹配",
    "PO_UNKNOWN": "未知 PO",
    "VALIDATION_WARNING": "字段校验警告/历史记录",
    "READY": "待人工确认",
}

BATCH_DESCRIPTIONS = {
    "VALIDATION_BLOCKED": "原始字段不完整或数值不可信，不能进入正式事实。",
    "PO_MISSING": "候选没有 PO 号，需要从原始 PDF 或 PO Master 确认。",
    "PO_POTENTIAL": "PO 号与 PO Master 存在相似候选，需要确认是否为 OCR 变形。",
    "PO_UNKNOWN": "PO 号未在当前 PO Master 中找到，需确认 PO 清单或原始文件。",
    "VALIDATION_WARNING": "旧候选缺少校验记录或存在警告，需按批次人工复核。",
    "READY": "字段校验通过且没有开放异常，但历史数据仍需显式人工确认。",
}


def _json_object(value: Any) -> dict[str, Any]:
    if isinstance(value, dict):
        return value
    if not value:
        return {}
    try:
        parsed = json.loads(value)
    except (TypeError, ValueError):
        return {}
    return parsed if isinstance(parsed, dict) else {}


def _period(candidate: dict[str, Any]) -> str:
    year = candidate.get("year")
    month = candidate.get("month")
    if year and month:
        return f"{int(year):04d}-{int(month):02d}"
    return "UNRESOLVED_PERIOD"


def _validation_state(candidate: dict[str, Any]) -> tuple[str, list[dict[str, Any]]]:
    payload = _json_object(candidate.get("parse_json"))
    validation = payload.get("validation") or {}
    if not isinstance(validation, dict):
        validation = {}
    status = str(validation.get("status") or "").upper()
    issues = [issue for issue in (validation.get("issues") or []) if isinstance(issue, dict)]
    if status == "BLOCKED" or any(str(issue.get("severity") or "").upper() == "BLOCKING" for issue in issues):
        return "BLOCKED", issues
    if status in {"WARNING", "PASS"}:
        return status, issues
    return "MISSING", issues


def _po_batch(candidate: dict[str, Any], open_reviews: Iterable[dict[str, Any]]) -> str | None:
    """Return the PO batch emitted by the latest audit, if any."""
    for item in open_reviews:
        queue_type = str(item.get("queue_type") or "").upper()
        if queue_type == "MISSING_PO":
            return "PO_MISSING"
        if queue_type == "UNKNOWN_PO":
            payload = _json_object(item.get("payload_json"))
            suggestions = payload.get("suggestions") or []
            return "PO_POTENTIAL" if suggestions else "PO_UNKNOWN"
    # A legacy candidate may have a PO value but no queue item after a partial
    # audit.  Do not infer an exact match here; the validation batch remains the
    # safe fallback.
    return None


def classify_candidate(candidate: dict[str, Any], review_items: Iterable[dict[str, Any]] = ()) -> dict[str, Any]:
    """Classify one candidate into one primary batch plus secondary reasons."""
    items = [dict(item) for item in review_items]
    open_items = [item for item in items if str(item.get("status") or "OPEN").upper() == "OPEN"]
    validation_status, validation_issues = _validation_state(candidate)
    po_batch = _po_batch(candidate, open_items)

    if validation_status == "BLOCKED" or any(
        str(item.get("queue_type") or "").upper() == "VALIDATION_BLOCKED" for item in open_items
    ):
        primary = "VALIDATION_BLOCKED"
    elif po_batch:
        primary = po_batch
    elif validation_status in {"WARNING", "MISSING"} or open_items:
        primary = "VALIDATION_WARNING"
    else:
        primary = "READY"

    secondary = []
    if po_batch and po_batch != primary:
        secondary.append(po_batch)
    if validation_status in {"WARNING", "MISSING"} and primary != "VALIDATION_WARNING":
        secondary.append("VALIDATION_WARNING")
    if validation_status == "BLOCKED" and primary != "VALIDATION_BLOCKED":
        secondary.append("VALIDATION_BLOCKED")

    return {
        "batch": primary,
        "batch_label": BATCH_LABELS[primary],
        "secondary_batches": list(dict.fromkeys(secondary)),
        "period": _period(candidate),
        "validation_status": validation_status,
        "validation_issues": validation_issues,
        "open_review_count": len(open_items),
        "open_queue_types": sorted({str(item.get("queue_type") or "") for item in open_items if item.get("queue_type")}),
    }


def summarise_batches(items: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    """Return deterministic counts for the Review Center batch selector."""
    grouped: dict[str, list[dict[str, Any]]] = {key: [] for key in BATCH_ORDER}
    for item in items:
        batch = item.get("batch")
        if batch in grouped:
            grouped[batch].append(item)

    result = []
    for batch in BATCH_ORDER:
        rows = grouped[batch]
        periods = Counter(row.get("period") or "UNRESOLVED_PERIOD" for row in rows)
        result.append({
            "key": batch,
            "label": BATCH_LABELS[batch],
            "description": BATCH_DESCRIPTIONS[batch],
            "count": len(rows),
            "periods": [{"period": period, "count": count} for period, count in sorted(periods.items())],
            "blocking": batch in {"VALIDATION_BLOCKED", "PO_MISSING", "PO_UNKNOWN"},
        })
    return result

