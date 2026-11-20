"""Conservative Daily Hours table association from positioned OCR blocks."""

from __future__ import annotations

import re
from collections import defaultdict
from typing import Any, Iterable


_NUMBER = re.compile(r"(?<![A-Za-z0-9])(\d{1,2}(?:[.,]\d{1,2})?)(?![A-Za-z0-9])")


def extract_daily_hours_from_blocks(
    blocks: Iterable[dict[str, Any]],
    *,
    page_height: float | None = None,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Extract only unambiguous daily-hour cells from OCR TSV blocks.

    The function deliberately returns no cells when the table anchors or row
    candidates are ambiguous.  It is suitable for diagnostics and review
    evidence; callers must still decide whether to use the values for totals.
    """

    items = [item for item in blocks if item.get("text") and item.get("left") is not None]
    daily = next((item for item in items if "daily" in str(item["text"]).casefold()), None)
    overtime = next((item for item in items if "overtime" in str(item["text"]).casefold()), None)
    total = next((item for item in items if str(item["text"]).casefold() == "total"), None)
    diagnostics: dict[str, Any] = {
        "status": "NO_LAYOUT",
        "block_count": len(items),
        "daily_anchor": bool(daily),
        "overtime_anchor": bool(overtime),
        "total_anchor": bool(total),
        "cells": 0,
    }
    if daily is None or total is None:
        return [], diagnostics

    daily_line = tuple(daily.get("line") or ())
    overtime_line = tuple(overtime.get("line") or ()) if overtime else ()
    total_line = tuple(total.get("line") or ())
    same_header_line = bool(daily_line and total_line and daily_line == total_line)
    daily_overtime_line = bool(daily_line and overtime_line and daily_line == overtime_line)
    def _near(left: dict[str, Any], right: dict[str, Any]) -> bool:
        left_top = float(left.get("top", 0))
        right_top = float(right.get("top", 0))
        left_height = float(left.get("height", 0))
        right_height = float(right.get("height", 0))
        return abs(left_top - right_top) <= max(left_height, right_height, 1) * 2

    nearby_header = bool(
        overtime is not None
        and _near(daily, overtime)
        and _near(daily, total)
        and float(total.get("left", 0)) > float(overtime.get("left", 0))
    )
    if daily_line and total_line and not same_header_line and not daily_overtime_line and not nearby_header:
        diagnostics["status"] = "NO_LAYOUT"
        diagnostics["reason"] = "table anchors are not on the same OCR line"
        return [], diagnostics

    daily_left = float(daily["left"])
    total_left = float(total["left"])
    if total_left <= daily_left:
        diagnostics["status"] = "INVALID_LAYOUT"
        return [], diagnostics
    # Prefer the Overtime boundary when OCR puts Daily and Total on separate
    # lines. This narrows the band around the actual Daily cells and avoids
    # task-description numbers to the left of the table.
    if (daily_overtime_line or nearby_header) and overtime is not None:
        right_bound = float(overtime["left"]) - max(6.0, float(overtime.get("width", 0)) * 0.2)
        estimated_width = max(120.0, float(daily.get("height", 0)) * 10.0)
        left_bound = max(daily_left + float(daily.get("width", 0)) * 0.25, right_bound - estimated_width)
        anchor_mode = "daily_overtime_boundary"
    else:
        left_bound = daily_left + float(daily.get("width", 0)) * 0.25
        right_bound = total_left - max(4.0, float(total.get("width", 0)) * 0.25)
        anchor_mode = "same_line_daily_total"
    header_line = daily_line
    header_bottom = float(daily.get("top", 0)) + float(daily.get("height", 0))
    rows: dict[tuple[int, ...] | int, list[dict[str, Any]]] = defaultdict(list)
    for index, item in enumerate(items):
        line = tuple(item.get("line") or ())
        if line and line == header_line:
            continue
        top = float(item.get("top", 0))
        if top < header_bottom:
            continue
        if page_height and top > page_height * 0.76:
            continue
        center = float(item["left"]) + float(item.get("width", 0)) / 2
        if not left_bound <= center <= right_bound:
            continue
        match = _NUMBER.fullmatch(str(item["text"]).strip().replace(",", "."))
        if not match:
            continue
        try:
            hours = float(match.group(1))
        except ValueError:
            continue
        if not 0 <= hours <= 24:
            continue
        key: tuple[int, ...] | int = line if line else round(top / 8)
        rows[key].append({"hours": hours, "block": item})

    cells: list[dict[str, Any]] = []
    ambiguous = 0
    for row_key, candidates in rows.items():
        if len(candidates) != 1:
            ambiguous += 1
            continue
        item = candidates[0]
        block = item["block"]
        cells.append({
            "row": list(row_key) if isinstance(row_key, tuple) else row_key,
            "hours": item["hours"],
            "raw_text": str(block["text"]),
            "bbox": [
                float(block["left"]), float(block.get("top", 0)),
                float(block["left"]) + float(block.get("width", 0)),
                float(block.get("top", 0)) + float(block.get("height", 0)),
            ],
        })
    diagnostics.update({
        "status": "OK" if cells else "NO_STABLE_CELLS",
        "cells": len(cells),
        "ambiguous_rows": ambiguous,
        "column_bounds": [left_bound, right_bound],
        "anchor_mode": anchor_mode,
    })
    return cells, diagnostics
