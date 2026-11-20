# -*- coding: utf-8 -*-
"""工时 PDF 增量扫描、文本/OCR解析和 Review Queue 登记。

本模块只生成解析候选，不把候选直接视为批准工时，也不写回 Excel。
精确重复文件按 SHA-256 去重；同一人员-月份-PO 出现不同文件版本时进入审核，
避免把版本文件自动相加。
"""

from __future__ import annotations

import argparse
import difflib
import json
import math
import os
import re
import unicodedata
from dataclasses import dataclass, asdict, field as dataclass_field
from pathlib import Path
from typing import Any, Iterable

import pdfplumber

import config
from document_processing import (
    TesseractProvider,
    assess_ocr_text,
    build_tesseract_router,
    build_timesheet_evidence,
    enhance_scan,
    needs_ocr_retry,
    ocr_candidate_rank,
    validate_ocr_completeness,
    validate_timesheet_fields,
)
from fact_store import FactStore
from project_registry import build_profile
from review_scope import path_in_scope, review_scope


@dataclass
class ParsedTimesheet:
    path: str
    raw_name: str = ""
    canonical_name: str = ""
    year: int | None = None
    month: int | None = None
    po_id: str = ""
    total_hours: float | None = None
    raw_unit: str = "hour"
    parse_method: str = "filename_only"
    text_excerpt: str = ""
    confidence: float = 0.0
    fields: dict[str, str] | None = None
    ocr_error: str = ""
    ocr_diagnostics: dict[str, Any] = dataclass_field(default_factory=dict)
    evidence: list[dict[str, Any]] = dataclass_field(default_factory=list)
    validation: dict[str, Any] = dataclass_field(default_factory=dict)


def _norm(value: str) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", value or "")).strip().casefold()


def _name_lookup() -> dict[str, str]:
    """跨项目第一版使用已有 CNTS 名单；未知人员必须进入 Review Queue。"""
    lookup: dict[str, str] = {}
    for canonical, info in getattr(config, "CNTS25165_PERSONNEL", {}).items():
        for value in [canonical, *info.get("aliases", [])]:
            lookup[_norm(value)] = canonical
    return lookup


def _extract_name_from_filename(path: Path) -> str:
    stem = path.stem
    lookup = _name_lookup()
    # 文件名常用下划线/连字符分隔（如 Worksheet_Pan_Chengchao_...），
    # 先当作空格再匹配，否则带空格的姓名变体会漏匹配。
    normalized_stem = _norm(re.sub(r"[_\-]+", " ", stem))
    for alias in sorted(lookup, key=len, reverse=True):
        if alias and alias in normalized_stem:
            return lookup[alias]

    # 删除日期、IAV、页码和常见文档前缀后，保留可用于人工确认的名称。
    value = re.sub(r"20\d{2}[-_ ]?\d{2}", " ", stem, flags=re.I)
    value = re.sub(r"\b(?:AI\d+|TMS|IAV|Worksheet|Timesheets?|External|management)\b", " ", value, flags=re.I)
    value = re.sub(r"[()_\-]+", " ", value)
    value = re.sub(r"\b\d+(?:\.\d+)?\b", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def _match_canonical(raw_name: str) -> str:
    if not raw_name:
        return ""
    return _name_lookup().get(_norm(raw_name), "")


def _extract_period(text: str, path: Path) -> tuple[int | None, int | None]:
    patterns = [
        # OCR often merges the form label into ``Month1Year`` or uses a
        # vertical bar for the slash.  Accept those variants but keep the
        # month/year numeric order strict.
        r"(?:Month\s*(?:[/|lI1]\s*)?Year|月\s*/\s*年)\s*[:：]?\s*(\d{1,2})\s*[/|]\s*(20\d{2})",
        r"\bYear\s*[:：]?\s*(20\d{2})\s+Month\s*[:：]?\s*(0?[1-9]|1[0-2])\b",
        r"\b(20\d{2})[-_/ ](0?[1-9]|1[0-2])\b",
        r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(?!\d)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            first, second = match.groups()
            if len(first) == 4:
                return int(first), int(second)
            return int(second), int(first)
    filename_month = re.search(r"(?<!\d)(0?[1-9]|1[0-2])\s*月", path.stem)
    if filename_month:
        for parent in [path.parent, *path.parents]:
            match = re.fullmatch(r"(20\d{2}|\d{2})(0[1-9]|1[0-2])", parent.name)
            if match:
                year = int(match.group(1))
                return (year if year > 100 else 2000 + year), int(filename_month.group(1))
    name_match = re.search(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(?!\d)", path.name)
    if name_match:
        return int(name_match.group(1)), int(name_match.group(2))
    for parent in [path.parent, *path.parents]:
        match = re.fullmatch(r"(20\d{2}|\d{2})(0[1-9]|1[0-2])", parent.name)
        if match:
            year = int(match.group(1))
            return (year if year > 100 else 2000 + year), int(match.group(2))
    return None, None


def _extract_po(text: str, path: Path) -> str:
    patterns = [
        r"(?:Purchase\s*Order\s*(?:No\.?|Number)?|要货单号|PO)[^0-9]{0,40}(\d{8,})",
        r"\bPO\s*(\d{8,})\b",
        r"\b(26\d{8})\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            return _normalise_po_id(match.group(1))
    for pattern in patterns:
        match = re.search(pattern, path.name, re.I)
        if match:
            return _normalise_po_id(match.group(1))
    return ""


def _known_po_ids() -> set[str]:
    known: set[str] = set()
    for item in getattr(config, "CNTS25165_PO_CONFIG", []):
        if item and item[0]:
            known.add(re.sub(r"\D", "", str(item[0])))
    return known


def _normalise_po_id(value: str) -> str:
    """Correct a one-character OCR digit error only when the project PO list makes it unambiguous."""
    digits = re.sub(r"\D", "", value or "")
    if not digits:
        return ""
    # On scanned forms Tesseract may read the ``O`` in the ``PO`` prefix as a
    # leading zero (for example ``P06600017130``).  The project PO format uses
    # ten digits for this historical family, so remove only this unambiguous
    # eleven-to-ten-digit OCR artefact before any master-data comparison.
    if len(digits) == 11 and digits.startswith("0"):
        digits = digits[1:]
    known = _known_po_ids()
    if digits in known or not known:
        return digits
    matches = sorted(
        ((difflib.SequenceMatcher(None, digits, candidate).ratio(), candidate) for candidate in known),
        reverse=True,
    )
    if matches and len(digits) >= 8 and len(digits) - len(matches[0][1]) in {-1, 0, 1} and matches[0][0] >= 0.9:
        return matches[0][1]
    return digits


_HOUR_MARKER_PATTERN = re.compile(
    r"ACTUAL_WORK_HOUR_VALUE\D{0,80}(\d{1,4}(?:\.\d+)?)", re.I,
)
_ESTIMATED_ROW = re.compile(r"estimated|预估|预\s*估", re.I)


def _to_hour(raw: str) -> float | None:
    try:
        value = float(raw)
    except (TypeError, ValueError):
        return None
    return value if 0 <= value <= 400 else None


def _extract_hours(text: str) -> float | None:
    # The OCR pass injects ACTUAL_WORK_HOUR_VALUE from a targeted crop of the
    # monthly-total cell, which makes it authoritative.  A loose full-page scan
    # can also match the label text with a misread value -- Chen Yakun 2026-09
    # read "Actual Work-hour of this month : 176.0" while the cell itself says
    # 160 -- and the old max() over every match then discarded the number the
    # dedicated crop had already read correctly.
    marker_values = [
        _to_hour(match.group(1)) for match in _HOUR_MARKER_PATTERN.finditer(text)
    ]
    marker_values = [value for value in marker_values if value is not None]
    if marker_values:
        return marker_values[0]

    patterns = [
        # Prefer the dedicated monthly footer total over daily/overtime values.
        r"(?:ACTUAL_WORK_HOUR_VALUE|Actual\s+Work[- ]?hour\s+of\s+this\s+month)\D{0,80}(\d{1,4}(?:\.\d+)?)",
        r"(?:Exact\s*[a-z]*\s*Work[- ]?hour|实际工作时间|Total\s*Hours?)\D{0,80}(\d{1,4}(?:\.\d+)?)\s*(?:Hours?|H|h)",
    ]
    for pattern in patterns:
        values = []
        for match in re.finditer(pattern, text, re.I):
            # The forms print an Estimated row under the Actual row.  When the
            # Actual cell is blank the non-digit gap runs past the next label
            # into the Estimated figure; refuse that match.
            if _ESTIMATED_ROW.search(text[match.start():match.start(1)]):
                continue
            value = _to_hour(match.group(1))
            if value is not None:
                values.append(value)
        if values:
            return max(values)
    return None


def _extract_daily_days_sum(text: str) -> float | None:
    """Sum explicit CGI-style ``Work location / Days`` values conservatively."""
    if not re.search(
        r"Date\s+Day\s+of\s+Week[\s\S]{0,120}Work\s+location\s+Days",
        text or "",
        re.I,
    ):
        return None
    explicit_sum = re.search(r"\bSUM\s+(\d{1,3}(?:[.,]\d+)?)\b", text or "", re.I)
    if explicit_sum:
        total = round(float(explicit_sum.group(1).replace(",", ".")), 2)
        if 0 <= total <= 400:
            return total
    values = [
        float(value)
        for value in re.findall(
            r"(?:Office|Remote|Home)\s+(0(?:\.\d+)?|1(?:\.\d+)?)\b",
            text or "",
            re.I,
        )
    ]
    if not values:
        return None
    total = round(sum(values), 2)
    return total if 0 <= total <= 400 else None


def _timesheet_field_score(text: str, path: Path) -> int:
    """Count extractable critical fields for native/OCR candidate routing."""
    year, month = _extract_period(text, path)
    name_signal = bool(re.search(
        r"(?:姓名\s*Name|(?<![A-Za-z])Name)\s*[:：]?[\s\S]{1,60}",
        text or "",
        re.I,
    )) or bool(_extract_name_from_filename(path))
    return sum((
        bool(year and month),
        bool(_extract_po(text, path)),
        _extract_hours(text) is not None,
        name_signal,
    ))


def _native_text_needs_ocr(text: str, path: Path) -> bool:
    """Trigger OCR only when a non-empty native text layer misses a key field."""
    return bool(text.strip()) and (
        _extract_hours(text) is None
        or not _extract_po(text, path)
        or _timesheet_field_score(text, path) < 4
    )


def _normalise_ocr_hour(value: str) -> float | None:
    """Convert one OCR token to a plausible monthly hour value.

    Tesseract often drops the decimal point in footer values such as
    ``156.0``.  Only four-digit integers outside the valid monthly range are
    repaired; ambiguous in-range integers are left unchanged.
    """
    match = re.search(r"(?<!\d)(\d{1,5}(?:[.,]\d+)?)(?!\d)", value or "")
    if not match:
        return None
    token = match.group(1).replace(",", ".")
    try:
        number = float(token)
    except ValueError:
        return None
    if number > 400 and "." not in token and len(token) in {3, 4} and 400 <= number <= 4000:
        number /= 10.0
    return number if 0 <= number <= 400 else None


def _hours_from_actual_line(text: str) -> float | None:
    """Read a value only from a line explicitly labelled as monthly actual."""
    for raw_line in (text or "").splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        actual = re.search(r"Actual\s+Work[- ]?hour", line, re.I)
        if not actual:
            continue
        # OCR commonly merges the words into ``thismonth|``.
        month = re.search(r"month", line[actual.end():], re.I)
        if not month:
            continue
        tail = line[actual.end() + month.end():]
        value = _normalise_ocr_hour(tail)
        if value is not None:
            return value
    return None


def _run_tesseract_image(executable: Path, image: Any, psm: str = "6", whitelist: str = "") -> tuple[str, str]:
    """Run the configured provider while preserving the legacy text/error API."""
    result = build_tesseract_router(
        executable, languages=config.OCR_LANGUAGE, timeout=90,
    ).recognize(image, psm=psm, whitelist=whitelist)
    return result.text, result.error


def _run_tesseract_tsv(executable: Path, image: Any, psm: str) -> tuple[list[dict[str, Any]], str]:
    """Return positioned words through the provider-neutral OCR block model."""
    result = TesseractProvider(
        executable, languages=config.OCR_LANGUAGE, timeout=90,
    ).recognize(image, psm=psm, page=1, output_mode="tsv")
    words: list[dict[str, Any]] = []
    for block in result.blocks:
        metadata = block.metadata
        if not block.text or block.bbox is None:
            continue
        left, top, right, bottom = block.bbox
        words.append({
            "line": block.line_id or (),
            "left": int(left), "top": int(top),
            "width": int(right - left), "height": int(bottom - top),
            "text": block.text,
        })
    return words, result.error


def _clean_footer_number_crop(image: Any, box: tuple[int, int, int, int]) -> Any:
    """Remove form grid lines that otherwise merge with footer digits."""
    from PIL import ImageDraw, ImageOps

    crop = ImageOps.grayscale(image.crop(box))
    width, height = crop.size
    binary = crop.point(lambda pixel: 0 if pixel < 170 else 255)
    pixels = binary.load()
    horizontal = [
        y for y in range(height)
        if sum(pixels[x, y] == 0 for x in range(width)) > width * 0.45
    ]
    vertical = [
        x for x in range(width)
        if sum(pixels[x, y] == 0 for y in range(height)) > height * 0.65
    ]
    cleaned = crop.copy()
    draw = ImageDraw.Draw(cleaned)
    line_width = max(3, round(min(width, height) / 12))
    for y in horizontal:
        draw.line((0, y, width, y), fill=255, width=line_width)
    for x in vertical:
        draw.line((x, 0, x, height), fill=255, width=line_width)
    cleaned = ImageOps.autocontrast(cleaned)
    return cleaned.resize((width * 8, height * 8))


def _footer_actual_row_center(image: Any) -> float | None:
    """Locate the monthly-actual row from the form's horizontal grid lines."""
    from PIL import ImageOps

    gray = ImageOps.grayscale(image)
    width, height = gray.size
    pixels = gray.load()
    narrow_left, narrow_right = int(width * 0.72), int(width * 0.84)
    broad_left, broad_right = int(width * 0.05), int(width * 0.95)
    candidates: list[tuple[int, float]] = []
    for y in range(int(height * 0.76), int(height * 0.91)):
        narrow_score = sum(
            pixels[x, y] < 160 for x in range(narrow_left, narrow_right)
        ) / max(1, narrow_right - narrow_left)
        broad_score = sum(
            pixels[x, y] < 160 for x in range(broad_left, broad_right)
        ) / max(1, broad_right - broad_left)
        if narrow_score >= 0.55 and broad_score >= 0.58:
            candidates.append((y, broad_score))

    groups: list[list[tuple[int, float]]] = []
    for item in candidates:
        if not groups or item[0] - groups[-1][-1][0] > 3:
            groups.append([])
        groups[-1].append(item)
    lines = sorted(max(group, key=lambda item: item[1])[0] for group in groups)
    if len(lines) < 2:
        return None

    max_footer_gap = height * 0.018
    last_gap = lines[-1] - lines[-2]
    if last_gap <= 3 or last_gap > max_footer_gap:
        return None
    # Two-row forms place Actual above Estimated; single-row forms end after
    # Actual.  Equal short gaps between the final three lines identify the
    # two-row variant without relying on a project/template name.
    if len(lines) >= 3:
        previous_gap = lines[-2] - lines[-3]
        if 3 < previous_gap <= max_footer_gap:
            return (lines[-3] + lines[-2]) / 2
    return (lines[-2] + lines[-1]) / 2


def _bbox_part(bbox: Any, index: int) -> float:
    try:
        return float(bbox[index])
    except (TypeError, IndexError, ValueError, KeyError):
        return float("nan")


def _pick_monthly_total(
    candidates: dict[float, dict[str, Any]],
    width: int | None = None,
) -> float:
    """Choose the candidate that is the monthly Actual total.

    Template facts decide this, in order:

    * The total sits in the right-hand footer block.  Daily cells in the left
      half of the footer band also have an ``Hours`` token beside them, so
      agreement alone can rank one of them above the real total.
    * These forms print ``Estimated`` directly below ``Actual`` and both rows
      carry an ``Hours`` cell, so two candidates in the same column tie on
      agreement.  The old ``max(..., key=(count, value))`` rule then always
      returned the larger -- which is the Estimated row by construction.
      The upper row is the actual one.
    * Every one of these cells is printed with one decimal place, so a read
      that kept the decimal came from a window holding the whole cell, while a
      bare integer came from a clipped one.  202605 刘旭 read ``155`` and
      ``158.0`` from two overlapping value-column crops that agreed with each
      other equally.

    A final larger-value tie-break keeps the remaining ambiguity deterministic
    and matches the historical behaviour where none of the above applies.
    """
    if len(candidates) <= 1:
        return next(iter(candidates))

    pool = candidates
    if width:
        right_side = {
            value: item for value, item in candidates.items()
            if _bbox_part(item.get("bbox"), 2) >= width * 0.55
        }
        if right_side:
            pool = right_side
    if len(pool) == 1:
        return next(iter(pool))

    best_count = max(item["count"] for item in pool.values())
    leaders = [value for value, item in pool.items() if item["count"] == best_count]
    if len(leaders) == 1:
        return leaders[0]

    lefts = [_bbox_part(pool[value].get("bbox"), 0) for value in leaders]
    if all(math.isfinite(left) for left in lefts):
        tolerance = (width or 0) * 0.05
        if max(lefts) - min(lefts) <= tolerance:
            # Same column: Actual is the row printed above Estimated.
            return min(leaders, key=lambda value: _bbox_part(pool[value].get("bbox"), 1))

    precise = [
        value for value in leaders
        if "." in str(pool[value].get("raw_text") or "")
    ]
    if len(precise) == 1:
        return precise[0]
    return max(leaders)


def _extract_bottom_total_ocr(
    executable: Path,
    image: Any,
    *,
    return_details: bool = False,
) -> float | tuple[float | None, dict[str, Any]] | None:
    """Read the dedicated monthly total cell from a scanned timesheet page."""
    from PIL import ImageOps

    def finish(
        value: float | None,
        method: str,
        *,
        bbox: tuple[int, int, int, int] | None = None,
        raw_text: str = "",
        extra: dict[str, Any] | None = None,
    ) -> float | tuple[float | None, dict[str, Any]]:
        details: dict[str, Any] = {
            "status": "FOUND" if value is not None else "NOT_FOUND",
            "method": method,
        }
        if value is not None:
            details["value"] = value
        if bbox is not None:
            details["bbox"] = list(bbox)
        if raw_text:
            details["raw_text"] = raw_text[:300]
        if extra:
            details.update(extra)
        return (value, details) if return_details else value

    width, height = image.size
    # Tesseract reads the footer digits more reliably at the original form
    # scale (roughly 150 DPI) than on the 300 DPI full-page render.
    if width > 1600:
        image = image.resize((width // 2, height // 2))
        width, height = image.size
    # August introduced a second CEA layout whose total cell is farther right
    # and whose footer moves vertically with the table.  OCR the whole footer
    # first so the English label anchors the row instead of guessing one fixed
    # crop.  Positioned TSV data then lets us make a tight, grid-line-free crop
    # when Tesseract sees the label but misses its number.
    footer_box = (
        int(width * 0.03), int(height * 0.78),
        int(width * 0.98), int(height * 0.91),
    )
    footer = ImageOps.autocontrast(image.crop(footer_box).convert("L"))
    scale = 3
    footer = footer.resize((footer.width * scale, footer.height * scale))
    label_anchors: list[tuple[float, float]] = []
    labelled_hour_values: set[float] = set()
    unlabelled_hour_values: set[float] = set()
    positioned_hour_values: dict[float, dict[str, Any]] = {}

    def is_hours_word(value: str) -> bool:
        """Recognise common OCR variants of the footer ``Hours`` cell."""
        token = re.sub(r"[^a-z]", "", value or "").casefold()
        if not token:
            return False
        if token == "h" or token in {"hour", "hours", "hous", "hows", "frows", "fours", "tours"}:
            return True
        return len(token) >= 4 and difflib.SequenceMatcher(None, token, "hours").ratio() >= 0.58

    # Keep the original low-resolution page coordinates for this pass. On
    # Q1 GR forms, cropping and enlarging the footer changes ``1140`` into
    # ``$140``; the unscaled TSV preserves the four digits and the existing
    # OCR-hour normaliser safely restores the dropped decimal.
    direct_positioned_hour_values: dict[float, dict[str, Any]] = {}
    for psm in ("11", "12"):
        words, _error = _run_tesseract_tsv(executable, image.copy(), psm)
        numeric_words = []
        hours_words = []
        for word in words:
            # The monthly block holds two rows -- ``Actual`` above
            # ``Estimated`` -- and slides up the page when the timesheet has
            # few daily rows.  Xu Chang 2026-09 put its Actual row at 79.7% of
            # the page height and the real 160.0 was dropped by the old 80%
            # floor, leaving the Estimated 176.0 as the only candidate.  The
            # page fraction is only a coarse guard now; the right-hand Hours
            # column below is what actually identifies the monthly cells.
            if word["top"] < height * 0.75 or word["top"] > height * 0.95:
                continue
            number_match = re.search(r"(?<!\d)(\d{1,5}(?:[.,]\d+)?)(?!\d)", word["text"] or "")
            if number_match:
                value = _normalise_ocr_hour(number_match.group(1))
                raw_number = number_match.group(1)
                precise = "." in raw_number or "," in raw_number or len(raw_number) >= 2
                # A footer total such as 184.0 may appear as a five-digit
                # noisy token (for example ``18600``). Keep that token only
                # when it is later tied to the Hours cell and re-read from a
                # tight crop; never accept the five-digit value itself.
                #
                # Four-digit tokens need the same treatment when the hour
                # normaliser rejects them as out of range: Xu Chang 2026-09
                # read the Actual cell ``160.0`` as ``7600``, which is over
                # 4000 and so survives no repair, and the row disappeared --
                # leaving the Estimated 176.0 as the only candidate.
                noisy_token = (
                    value is None and raw_number.isdigit() and len(raw_number) == 4
                )
                # ``_normalise_ocr_hour`` restores a dropped decimal point by
                # dividing a bare three or four digit token by ten, which
                # guesses the magnitude rather than reading it.  A tight crop of
                # the cell is better evidence: Yang Yumeng 2026-09 holds 64.0
                # and the full-page token was ``680``.
                repaired = (
                    value is not None
                    and raw_number.isdigit()
                    and len(raw_number) in (3, 4)
                    and abs(value * 10.0 - float(raw_number)) < 1e-9
                )
                if (value is not None or len(raw_number) == 5 or noisy_token) and precise:
                    numeric_words.append((word, value, repaired))
            if is_hours_word(word["text"]) and word["left"] >= width * 0.75:
                # Only the right-hand Hours cells belong to the monthly rows.
                # The absence row carries its own Hours cell in the left half,
                # and its number would otherwise qualify as a monthly total.
                hours_words.append(word)
        for number_word, value, repaired in numeric_words:
            number_right = number_word["left"] + number_word["width"]
            number_center = number_word["top"] + number_word["height"] / 2
            for hours_word in hours_words:
                hours_left = hours_word["left"]
                hours_center = hours_word["top"] + hours_word["height"] / 2
                same_line = number_word["line"] == hours_word["line"]
                close_line = abs(number_center - hours_center) <= max(12, hours_word["height"] * 1.5)
                if number_right <= hours_left + 12 and (same_line or close_line):
                    if value is None or repaired:
                        crop_box = (
                            max(0, number_word["left"] - 16),
                            max(0, number_word["top"] - 8),
                            min(width, number_right + 10),
                            min(height, number_word["top"] + number_word["height"] + 8),
                        )
                        number_crop = image.crop(crop_box).resize((
                            max(160, (crop_box[2] - crop_box[0]) * 4),
                            max(64, (crop_box[3] - crop_box[1]) * 4),
                        ))
                        # The re-read must carry the same number of digits the
                        # token was seen with -- ``7600`` and ``160.0`` both
                        # carry four, ``680`` and ``64.0`` both carry three --
                        # so an unrelated fragment inside the crop cannot
                        # substitute itself.  Five-digit tokens are excluded:
                        # the extra digit is the noise being removed.
                        expected_digits = (
                            len(raw_number)
                            if raw_number.isdigit() and len(raw_number) in (3, 4)
                            else 0
                        )
                        for crop_psm in ("6", "7", "10", "11", "12"):
                            crop_text, _crop_error = _run_tesseract_image(
                                executable, number_crop, psm=crop_psm,
                                whitelist="0123456789.",
                            )
                            candidate_value = _normalise_ocr_hour(crop_text)
                            if candidate_value is None:
                                continue
                            if expected_digits and len(
                                re.sub(r"\D", "", crop_text or "")
                            ) != expected_digits:
                                continue
                            value = candidate_value
                            break
                        number_crop.close()
                    if value is None:
                        break
                    item = direct_positioned_hour_values.setdefault(value, {
                        "count": 0,
                        "bbox": [
                            number_word["left"], number_word["top"],
                            number_word["left"] + number_word["width"],
                            number_word["top"] + number_word["height"],
                        ],
                        "raw_text": number_word["text"],
                    })
                    item["count"] += 1
                    break
    if direct_positioned_hour_values:
        selected = _pick_monthly_total(direct_positioned_hour_values, width)
        item = direct_positioned_hour_values[selected]
        if item["count"] >= 2:
            return finish(
                selected,
                "direct_positioned_hours_cell_tsv",
                bbox=tuple(round(float(value)) for value in item["bbox"]),
                raw_text=str(item.get("raw_text", "")),
                extra={
                    "page_width": width,
                    "rival_values": sorted(
                        v for v in direct_positioned_hour_values if v != selected
                    ),
                },
            )

    for psm in ("6", "11", "12"):
        words, _error = _run_tesseract_tsv(executable, footer, psm)
        # Q1 GR forms sometimes preserve the footer number but break the
        # label into a separate TSV line.  When the number is immediately
        # before an OCR variant of ``Hours``, keep it as a positioned
        # candidate.  Requiring agreement across sparse-page modes prevents
        # ordinary table numbers from becoming a monthly total.
        numeric_words = []
        hours_words = []
        for word in words:
            number_match = re.search(r"(?<!\d)(\d{1,5}(?:[.,]\d+)?)(?!\d)", word["text"] or "")
            if number_match:
                value = _normalise_ocr_hour(number_match.group(1))
                raw_number = number_match.group(1)
                precise = "." in raw_number or "," in raw_number or len(raw_number) >= 2
                if value is not None and precise:
                    numeric_words.append((word, value))
            if is_hours_word(word["text"]):
                hours_words.append(word)
        for number_word, value in numeric_words:
            number_right = number_word["left"] + number_word["width"]
            number_center = number_word["top"] + number_word["height"] / 2
            for hours_word in hours_words:
                hours_left = hours_word["left"]
                hours_center = hours_word["top"] + hours_word["height"] / 2
                same_line = number_word["line"] == hours_word["line"]
                close_line = abs(number_center - hours_center) <= max(12, hours_word["height"] * 1.5)
                if number_right <= hours_left + 12 and (same_line or close_line):
                    item = positioned_hour_values.setdefault(value, {
                        "count": 0,
                        "bbox": [
                            footer_box[0] + number_word["left"] / scale,
                            footer_box[1] + number_word["top"] / scale,
                            footer_box[0] + (number_word["left"] + number_word["width"]) / scale,
                            footer_box[1] + (number_word["top"] + number_word["height"]) / scale,
                        ],
                        "raw_text": number_word["text"],
                    })
                    item["count"] += 1
                    break
        grouped: dict[tuple[int, ...], list[dict[str, Any]]] = {}
        for word in words:
            grouped.setdefault(word["line"], []).append(word)
        for line_words in grouped.values():
            line_words.sort(key=lambda item: item["left"])
            line = " ".join(item["text"] for item in line_words)
            labelled_value = _hours_from_actual_line(line)
            if labelled_value is not None:
                # Wide footer OCR establishes a useful candidate, but grid
                # lines can turn 175 into 178 or 159 into 188.  Use it only if
                # the positioned, line-cleaned number cell cannot be read.
                labelled_hour_values.add(labelled_value)

            actual_index = next(
                (index for index, item in enumerate(line_words) if "actual" in item["text"].casefold()),
                None,
            )
            if actual_index is not None and any(
                "work-hour" in item["text"].casefold() or item["text"].casefold() == "work"
                for item in line_words[actual_index:]
            ):
                month_word = next(
                    (item for item in line_words[actual_index:] if "month" in item["text"].casefold()),
                    None,
                )
                if month_word is None:
                    continue
                label_right = footer_box[0] + (month_word["left"] + month_word["width"]) / scale
                actual_word = line_words[actual_index]
                center_y = footer_box[1] + (
                    actual_word["top"] + actual_word["height"] / 2
                ) / scale
                label_anchors.append((label_right, center_y))

            # Some scans lose the label but still read the sole footer value
            # and its Hours cell.  Keep this only as an unambiguous fallback.
            for match in re.finditer(r"(\d{1,4}(?:[.,]\d+)?)\s*[|\[\]]?\s*Hou(?:rs?|s)\b", line, re.I):
                value = _normalise_ocr_hour(match.group(1))
                if value is not None:
                    unlabelled_hour_values.add(value)
        if label_anchors:
            break

    if label_anchors:
        # Prefer the rightmost/bottommost label observation; sparse-page modes
        # generally locate the footer more accurately than full-block mode.
        label_right, center_y = max(label_anchors, key=lambda item: (item[1], item[0]))
        # The value cell starts just after the English ``month`` label.  Its
        # absolute x-position differs between AI2, AI4 and reduced-width scans,
        # so derive the crop from the detected label instead of the template.
        left = max(0, round(label_right + width * 0.015))
        # Stop before the adjacent ``Hours`` cell; its vertical border/text can
        # otherwise turn 144.5 into a spurious 144.93.
        right = min(width, round(label_right + width * 0.105))
        pad_y = max(18, round(height * 0.011))
        number_crop = _clean_footer_number_crop(image, (
            left, max(0, round(center_y) - pad_y),
            right, min(height, round(center_y) + pad_y),
        ))
        for psm in ("11", "6"):
            output, _error = _run_tesseract_image(
                executable, number_crop, psm=psm, whitelist="0123456789.",
            )
            value = _normalise_ocr_hour(output)
            raw_number = re.search(r"\d{1,5}(?:[.,]\d+)?", output or "")
            # A lone digit is usually a clipped grid/text fragment (for
            # example the old Chen Yakun crop returned ``1`` for 156.0).
            # Genuine short assignments in these forms retain a decimal.
            precise = bool(raw_number and (
                "." in raw_number.group(0) or "," in raw_number.group(0)
                or len(raw_number.group(0)) >= 2
            ))
            if value is not None and precise:
                return finish(
                    value,
                    "label_anchored_number_crop",
                    bbox=(left, max(0, round(center_y) - pad_y), right, min(height, round(center_y) + pad_y)),
                    raw_text=output,
                )

    if positioned_hour_values:
        # Bounding boxes here were translated back into page coordinates, so
        # the same right-hand-block rule applies as in the direct pass.
        selected = _pick_monthly_total(positioned_hour_values, width)
        item = positioned_hour_values[selected]
        if item["count"] >= 2:
            return finish(
                selected,
                "positioned_hours_cell_tsv",
                bbox=tuple(round(float(value)) for value in item["bbox"]),
                raw_text=str(item.get("raw_text", "")),
                extra={
                    "page_width": width,
                    "rival_values": sorted(
                        v for v in positioned_hour_values if v != selected
                    ),
                },
            )

    # Some signed multi-page scans keep the monthly footer label and value on
    # page 2, but Tesseract does not place the value and ``Hours`` in one TSV
    # line. Require the explicit footer label first, then read only the value
    # column immediately to its right.
    footer_label_box = (
        int(width * 0.35), int(height * 0.82),
        int(width * 0.70), int(height * 0.91),
    )
    footer_label_crop = image.crop(footer_label_box).resize((
        (footer_label_box[2] - footer_label_box[0]) * 3,
        (footer_label_box[3] - footer_label_box[1]) * 3,
    ))
    footer_label_found = False
    try:
        for psm in ("6", "11", "12"):
            label_text, _error = _run_tesseract_image(
                executable, footer_label_crop, psm=psm,
            )
            if re.search(
                r"Actual\s+Work[- ]?hour[\s\S]{0,80}month",
                label_text or "",
                re.I,
            ):
                footer_label_found = True
                break
    finally:
        footer_label_crop.close()
    if footer_label_found:
        value_box = (
            int(width * 0.667), int(height * 0.848),
            int(width * 0.785), int(height * 0.885),
        )
        value_crop = image.crop(value_box).resize((
            (value_box[2] - value_box[0]) * 4,
            (value_box[3] - value_box[1]) * 4,
        ))
        try:
            for psm in ("6", "7", "10", "11", "12"):
                output, _error = _run_tesseract_image(
                    executable, value_crop, psm=psm,
                    whitelist="0123456789.",
                )
                value = _normalise_ocr_hour(output)
                raw_number = re.search(r"\d{1,5}(?:[.,]\d+)?", output or "")
                precise = bool(raw_number and (
                    "." in raw_number.group(0)
                    or "," in raw_number.group(0)
                    or len(raw_number.group(0)) >= 2
                ))
                if value is not None and precise:
                    return finish(
                        value,
                        "footer_label_value_crop",
                        bbox=value_box,
                        raw_text=output,
                    )
        finally:
            value_crop.close()

    # If OCR loses the English label entirely, locate the footer row from the
    # table grid and try the known value-column bands.  Repeated agreement is
    # preferred, while impossible results (over 400 hours) are discarded.
    row_center = _footer_actual_row_center(image)
    if row_center is not None:
        value_counts: dict[float, dict[str, Any]] = {}
        pad_y = max(18, round(height * 0.011))
        for left, right in ((0.66, 0.80), (0.70, 0.84), (0.76, 0.88), (0.82, 0.93)):
            crop_box = (
                int(width * left), max(0, round(row_center) - pad_y),
                int(width * right), min(height, round(row_center) + pad_y),
            )
            number_crop = _clean_footer_number_crop(image, (
                *crop_box,
            ))
            output, _error = _run_tesseract_image(
                executable, number_crop, psm="11", whitelist="0123456789.",
            )
            value = _normalise_ocr_hour(output)
            raw_number = re.search(r"\d{1,5}(?:[.,]\d+)?", output or "")
            precise = bool(raw_number and (
                "." in raw_number.group(0) or "," in raw_number.group(0)
                or len(raw_number.group(0)) >= 2
            ))
            if value is not None and precise:
                item = value_counts.setdefault(value, {"count": 0, "bbox": list(crop_box), "raw_text": output})
                item["count"] += 1
        if value_counts:
            # Every crop here already sits on the grid-derived actual row and
            # in the right-hand value band, so agreement breaks the tie.
            selected = _pick_monthly_total(value_counts, width)
            item = value_counts[selected]
            return finish(
                selected,
                "grid_row_consensus",
                bbox=tuple(item["bbox"]),
                raw_text=str(item.get("raw_text", "")),
                extra={
                    "page_width": width,
                    "rival_values": sorted(v for v in value_counts if v != selected),
                },
            )

    if len(labelled_hour_values) == 1:
        return finish(next(iter(labelled_hour_values)), "labelled_footer_line")
    if len(unlabelled_hour_values) == 1:
        return finish(next(iter(unlabelled_hour_values)), "hours_cell_line")
    return finish(None, "not_found")


def _read_pdf_text(path: Path) -> str:
    with pdfplumber.open(str(path)) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages).strip()


def _ocr_pdf_with_diagnostics(
    path: Path,
) -> tuple[str, str, list[dict[str, Any]], dict[str, Any]]:
    """本地 OCR；用临时 PNG 避免 Tesseract 处理含中文路径时失败。"""
    executable = Path(getattr(config, "TESSERACT_EXE", ""))
    if not executable.exists():
        error = f"tesseract_not_found:{executable}"
        return "", error, [], {"route": "tesseract", "status": "ERROR", "error": error}
    temp_dir: Path | None = None
    document = None
    try:
        import pypdfium2 as pdfium
        document = pdfium.PdfDocument(str(path))
        rendered_pages = []
        for index in range(len(document)):
            page = document[index]
            bitmap = page.render(scale=float(getattr(config, "OCR_DPI", 300)) / 72.0)
            image = bitmap.to_pil()
            # Keep a lower-resolution copy for the small footer total cell.
            low_bitmap = page.render(scale=150.0 / 72.0)
            low_image = low_bitmap.to_pil()
            rendered_pages.append((image.copy(), low_image.copy()))
            image.close()
            bitmap.close()
            low_image.close()
            low_bitmap.close()
            page.close()
        document.close()
        document = None

        pages: list[str] = []
        evidence_blocks: list[dict[str, Any]] = []
        diagnostics: dict[str, Any] = {
            "route": "tesseract",
            "page_count": len(rendered_pages),
            "quality_gate": {
                "min_chars": 180,
                "required_markers": ["Actual", "Purchase Order"],
            },
            "pages": [],
        }
        for page_index, (image, low_image) in enumerate(rendered_pages):
            # Windows 环境下直接使用 mkstemp 比在临时目录中再创建文件可靠。
            full_text, full_error = _run_tesseract_image(executable, image, psm="6")
            selected_text = full_text
            selected_variant = "tesseract_psm6"
            page_attempts: list[dict[str, Any]] = []
            if page_index == 0:
                first_quality = assess_ocr_text(
                    full_text,
                    min_chars=180,
                    required_markers=("Actual", "Purchase Order"),
                )
                first_po = _extract_po(full_text, Path(""))
                page_attempts.append({
                    "variant": "tesseract_psm6",
                    **first_quality.as_dict(),
                    "po_id_present": bool(first_po),
                })
            # Keep the fast, proven path for normal pages.  Weak or incomplete
            # first passes get a contrast-normalised PSM 11 retry.  The better
            # complete candidate is selected instead of concatenating attempts,
            # which prevents duplicate OCR fragments from confusing parsers.
            if page_index == 0 and (
                not page_attempts or page_attempts[0]["should_retry"]
            ):
                enhanced = enhance_scan(image)
                try:
                    retry_text, _retry_error = _run_tesseract_image(
                        executable, enhanced, psm="11",
                    )
                    retry_quality = assess_ocr_text(
                        retry_text,
                        min_chars=180,
                        required_markers=("Actual", "Purchase Order"),
                    )
                    retry_po = _extract_po(retry_text, Path(""))
                    page_attempts.append({
                        "variant": "tesseract_enhanced_psm11",
                        **retry_quality.as_dict(),
                        "po_id_present": bool(retry_po),
                    })
                    if ocr_candidate_rank(
                        retry_text,
                        min_chars=180,
                        required_markers=("Actual", "Purchase Order"),
                    ) > ocr_candidate_rank(
                        full_text,
                        min_chars=180,
                        required_markers=("Actual", "Purchase Order"),
                    ) and (retry_po or not first_po):
                        selected_text = retry_text
                        selected_variant = "tesseract_enhanced_psm11"
                finally:
                    enhanced.close()
            # Sparse CGI/Q1 GR scans can hide the PO from both the normal and
            # contrast-normalised passes. Run PSM 12 only in that case; this
            # is a targeted field fallback rather than a third pass for every
            # Timesheet page.
            selected_po = _extract_po(selected_text, Path(""))
            if page_index == 0 and not selected_po:
                sparse_text, _sparse_error = _run_tesseract_image(
                    executable, image, psm="12",
                )
                sparse_quality = assess_ocr_text(
                    sparse_text,
                    min_chars=180,
                    required_markers=("Actual", "Purchase Order"),
                )
                sparse_po = _extract_po(sparse_text, Path(""))
                page_attempts.append({
                    "variant": "tesseract_psm12_po_fallback",
                    **sparse_quality.as_dict(),
                    "po_id_present": bool(sparse_po),
                })
                if sparse_po:
                    selected_text = sparse_text
                    selected_variant = "tesseract_psm12_po_fallback"
            # Some Q1 GR scans have a readable PO in the header but the full
            # page modes omit that narrow line. OCR the header field itself as
            # a final, field-scoped fallback; it is still sourced only from
            # the page image and never from PO master data.
            selected_po = _extract_po(selected_text, Path(""))
            if page_index == 0 and not selected_po:
                low_width, low_height = low_image.size
                # The CEA form places the PO number in the same narrow header
                # band across its scanned variants. Restricting the crop to
                # that field and whitelisting PO digits avoids table numbers
                # and makes handwritten/low-contrast PO text recoverable.
                header_po = ""
                header_text = ""
                header_attempt: dict[str, Any] = {
                    "variant": "tesseract_header_po_fallback",
                    "status": "FIELD_FALLBACK",
                    "po_id_present": False,
                }
                try:
                    # Most CEA forms work with the low-resolution template
                    # crop. Some embedded-text/scanned variants print the
                    # PO number smaller and need the original 300-DPI image
                    # plus a slightly wider field crop. Both routes remain
                    # field-scoped and accept a value only when a complete
                    # PO-like number is recognized.
                    header_variants = [
                        (
                            low_image,
                            (
                                int(low_width * 0.45), int(low_height * 0.055),
                                int(low_width * 0.70), int(low_height * 0.11),
                            ),
                            4,
                            "low_template",
                        ),
                        (
                            image,
                            (
                                int(image.width * 0.42), int(image.height * 0.04),
                                int(image.width * 0.76), int(image.height * 0.14),
                            ),
                            4,
                            "high_resolution_wide",
                        ),
                    ]
                    for header_source, header_box, resize_factor, source_name in header_variants:
                        header_image = header_source.crop(header_box).resize((
                            (header_box[2] - header_box[0]) * resize_factor,
                            (header_box[3] - header_box[1]) * resize_factor,
                        ))
                        try:
                            for header_psm in ("6", "11", "12"):
                                candidate_text, _header_error = _run_tesseract_image(
                                    executable, header_image, psm=header_psm,
                                    whitelist="PO0123456789",
                                )
                                candidate_po = _extract_po(candidate_text, Path(""))
                                if candidate_po:
                                    header_po = candidate_po
                                    header_text = candidate_text
                                    header_attempt["psm"] = header_psm
                                    header_attempt["source"] = source_name
                                    break
                            if header_po:
                                break
                        finally:
                            header_image.close()
                    header_attempt["po_id_present"] = bool(header_po)
                    header_attempt["text_excerpt"] = header_text[:300]
                    page_attempts.append(header_attempt)
                    if header_po:
                        selected_text = f"{selected_text}\nPurchase Order No.: {header_po}".strip()
                        selected_variant = "tesseract_header_po_fallback"
                finally:
                    pass
            if selected_text:
                pages.append(selected_text)
            if page_index == 0:
                diagnostics["pages"].append({
                    "page": page_index + 1,
                    "attempts": page_attempts,
                    "selected": selected_variant,
                })
            if page_index == 0:
                # Persist positioned blocks once per document for field-level
                # evidence; the text route remains the source of parsed values.
                evidence_blocks, _blocks_error = _run_tesseract_tsv(executable, image, psm="6")
            # The monthly total is usually on page 1, but signed multi-page
            # scans can place the same footer on page 2. Search later pages
            # only while no footer total has been found, so deliverable pages
            # cannot add unrelated numbers after a valid total is known.
            monthly_total_found = (
                isinstance(diagnostics.get("monthly_total"), dict)
                and diagnostics["monthly_total"].get("status") == "FOUND"
            )
            if page_index == 0 or not monthly_total_found:
                bottom_total, total_details = _extract_bottom_total_ocr(
                    executable, low_image, return_details=True,
                )
                if bottom_total is not None and not monthly_total_found:
                    total_details["page"] = page_index + 1
                    diagnostics["monthly_total"] = total_details
                    pages.append(f"ACTUAL_WORK_HOUR_VALUE: {bottom_total:g}")
            if full_error:
                # Large multi-page scans may time out on full-page OCR even
                # though their header and monthly footer fields are readable
                # from low-resolution crops. Preserve those field candidates
                # as a read-only OCR result instead of discarding the file.
                if page_index == 0 and len(rendered_pages) >= 3:
                    fallback_total = None
                    fallback_details: dict[str, Any] = {}
                    for candidate_page, (_candidate_image, candidate_low_image) in enumerate(rendered_pages):
                        candidate_total, candidate_details = _extract_bottom_total_ocr(
                            executable, candidate_low_image, return_details=True,
                        )
                        if candidate_total is not None:
                            fallback_total = candidate_total
                            fallback_details = dict(candidate_details)
                            fallback_details["page"] = candidate_page + 1
                            break
                    if fallback_total is not None:
                        if not any("ACTUAL_WORK_HOUR_VALUE:" in item for item in pages):
                            pages.append(f"ACTUAL_WORK_HOUR_VALUE: {fallback_total:g}")
                        diagnostics["monthly_total"] = fallback_details
                        diagnostics["status"] = "OK_WITH_FIELD_FALLBACK"
                        diagnostics["full_page_error"] = full_error
                        diagnostics["field_fallback"] = "header_and_monthly_footer_crops"
                        for rendered_image, rendered_low_image in rendered_pages:
                            rendered_image.close()
                            rendered_low_image.close()
                        return "\n".join(pages).strip(), "", evidence_blocks, diagnostics
                diagnostics["status"] = "ERROR"
                diagnostics["error"] = full_error
                return "\n".join(pages), full_error, evidence_blocks, diagnostics
            image.close()
            low_image.close()
        diagnostics["status"] = "OK"
        return "\n".join(pages).strip(), "", evidence_blocks, diagnostics
    except Exception as exc:  # OCR failure becomes a review item, not a hard failure.
        error = f"ocr_error:{type(exc).__name__}:{exc}"
        diagnostics = locals().get("diagnostics", {"route": "tesseract"})
        diagnostics["status"] = "ERROR"
        diagnostics["error"] = error
        return "", error, [], diagnostics
    finally:
        if document is not None and hasattr(document, "close"):
            document.close()
        if temp_dir is not None:
            import shutil
            shutil.rmtree(temp_dir, ignore_errors=True)


def _ocr_pdf(path: Path) -> tuple[str, str, list[dict[str, Any]]]:
    """Backward-compatible OCR helper without diagnostic metadata."""
    text, error, blocks, _diagnostics = _ocr_pdf_with_diagnostics(path)
    return text, error, blocks


def parse_pdf(path: str | os.PathLike[str], use_ocr: bool = True) -> ParsedTimesheet:
    pdf_path = Path(path)
    filename_name = _extract_name_from_filename(pdf_path)
    filename_canonical = _match_canonical(filename_name)
    year, month = _extract_period("", pdf_path)
    try:
        text = _read_pdf_text(pdf_path)
    except Exception as exc:
        text = ""
        read_error = f"pdf_text_error:{type(exc).__name__}:{exc}"
    else:
        read_error = ""

    native_text = text
    method = "pdf_text" if text else "filename_only"
    ocr_error = read_error
    ocr_blocks: list[dict[str, Any]] = []
    ocr_diagnostics: dict[str, Any] = {}
    if use_ocr and (not text or _native_text_needs_ocr(text, pdf_path)):
        ocr_text, ocr_candidate_error, ocr_blocks, ocr_diagnostics = _ocr_pdf_with_diagnostics(pdf_path)
        if not text:
            text = ocr_text
            ocr_error = ocr_candidate_error
            if text:
                method = "ocr"
        elif ocr_text and _timesheet_field_score(ocr_text, pdf_path) > _timesheet_field_score(native_text, pdf_path):
            text = ocr_text
            ocr_error = ocr_candidate_error
            method = "ocr"
        elif ocr_candidate_error and not native_text:
            ocr_error = ocr_candidate_error

    text_year, text_month = _extract_period(text, pdf_path)
    native_year, native_month = _extract_period(native_text, pdf_path) if native_text else (None, None)
    if native_year and native_month:
        # A native text layer may be missing only the PO while still carrying
        # a more reliable period than a low-quality OCR retry. Preserve that
        # field instead of replacing an otherwise correct 06/2026 with an OCR
        # digit artifact such as 08/2026.
        year, month = native_year, native_month
    elif text_year and text_month:
        year, month = text_year, text_month
    raw_name = filename_name
    canonical = filename_canonical
    name_match = re.search(
        r"(?:姓名\s*Name|(?<![A-Za-z])Name)\s*[:：]\s*"
        r"([A-Za-z][A-Za-z .'-]{1,60}?)(?=\s+(?:要货单号|Purchase\s*Order|专业科室|Department|公司|Company|中/外方|Lokal|月/年|Month)|$)",
        text, re.I,
    )
    if name_match:
        ocr_name = name_match.group(1).strip()
        ocr_canonical = _match_canonical(ocr_name)
        # The English OCR model often turns a Chinese name into an unrelated
        # Latin string (e.g. 赵士辉 -> "AHF SAMS").  A known filename name or
        # alias is more reliable in that case, so do not let bad OCR overwrite
        # a trustworthy file-name mapping.
        if ocr_canonical or not filename_canonical:
            raw_name = ocr_name
            canonical = ocr_canonical
    hours = _extract_hours(text)
    hours_source = "monthly_label" if hours is not None else ""
    if hours is None:
        hours = _extract_daily_days_sum(text)
        if hours is not None:
            hours_source = "daily_days_sum"
    if native_text:
        native_hours = _extract_hours(native_text)
        native_hours_source = "monthly_label" if native_hours is not None else ""
        if native_hours is None:
            native_hours = _extract_daily_days_sum(native_text)
            if native_hours is not None:
                native_hours_source = "daily_days_sum"
        if native_hours is not None:
            # As with the period, keep fields already present in the native
            # text layer when OCR was invoked only to fill a missing field.
            hours = native_hours
            hours_source = native_hours_source
    po_id = _extract_po(text, pdf_path)
    native_po_id = _extract_po(native_text, pdf_path) if native_text else ""
    if not po_id and native_po_id:
        po_id = native_po_id
    if method == "ocr" and native_text:
        preferred_fields = []
        if native_year and native_month:
            preferred_fields.append("period")
        if native_hours is not None:
            preferred_fields.append("hours")
        if native_po_id and po_id == native_po_id:
            preferred_fields.append("po_id")
        if preferred_fields:
            ocr_diagnostics["native_field_preference"] = preferred_fields
    fields = {
        "period": f"{year:04d}-{month:02d}" if year and month else "",
        "po_id": po_id,
        "hours": "" if hours is None else str(hours),
        "name": raw_name,
    }
    filled = sum(bool(value) for value in fields.values())
    confidence = min(1.0, filled / 4.0)
    if method == "ocr":
        confidence *= 0.85
    validation_report = validate_timesheet_fields(
        fields,
        source_path=str(pdf_path),
        known_po_ids=_known_po_ids(),
        total_hours=hours,
    )
    if hours_source:
        validation_report.checks["hours_source"] = hours_source
    if hours_source == "daily_days_sum":
        validation_report.checks["daily_days_sum"] = hours
    if method == "ocr":
        completeness = validate_ocr_completeness(
            text,
            required_markers=("Actual", "Purchase Order"),
        )
        validation_report.issues.extend(completeness.issues)
        validation_report.checks["ocr_completeness"] = completeness.as_dict()
        if completeness.status == "WARNING" and validation_report.status == "PASS":
            validation_report.status = "WARNING"
    field_blocks: dict[str, list[dict[str, Any]]] = {}
    monthly_total = ocr_diagnostics.get("monthly_total", {})
    monthly_bbox = monthly_total.get("bbox") if isinstance(monthly_total, dict) else None
    if hours is not None and isinstance(monthly_bbox, list) and len(monthly_bbox) == 4:
        left, top, right, bottom = (float(value) for value in monthly_bbox)
        field_blocks["hours"] = [{
            "text": str(hours),
            "left": left,
            "top": top,
            "width": max(0.0, right - left),
            "height": max(0.0, bottom - top),
        }]
    evidence = build_timesheet_evidence(
        fields,
        source_file=str(pdf_path),
        text=text,
        blocks=ocr_blocks,
        field_blocks=field_blocks,
        ocr_engine="tesseract" if method == "ocr" else method,
    )
    field_issues: dict[str, list[Any]] = {}
    for issue in validation_report.issues:
        if issue.field:
            field_issues.setdefault(issue.field, []).append(issue)
    for item in evidence:
        field_name = str(item.get("field", ""))
        issues = field_issues.get(field_name, [])
        item["validation_status"] = (
            "BLOCKED" if any(issue.severity == "BLOCKING" for issue in issues)
            else "WARNING" if issues else "PASS"
        )
    return ParsedTimesheet(
        path=str(pdf_path), raw_name=raw_name, canonical_name=canonical,
        year=year, month=month, po_id=po_id, total_hours=hours,
        parse_method=method, text_excerpt=text[:1200], confidence=confidence,
        fields=fields, ocr_error=ocr_error,
        ocr_diagnostics=ocr_diagnostics,
        evidence=evidence, validation=validation_report.as_dict(),
    )


def _version_key(parsed: ParsedTimesheet) -> str:
    name = _norm(parsed.canonical_name or parsed.raw_name) or "unknown"
    period = f"{parsed.year:04d}-{parsed.month:02d}" if parsed.year and parsed.month else "unknown-period"
    return "|".join((name, period, parsed.po_id or "unknown-po"))


def _review_types(parsed: ParsedTimesheet, existing: list[Any]) -> list[tuple[str, str]]:
    items: list[tuple[str, str]] = []
    if not parsed.canonical_name:
        items.append(("NEW_MEMBER", "无法将 PDF 名称匹配到项目人员名单"))
    if not parsed.year or not parsed.month:
        items.append(("MISSING_PERIOD", "无法可靠识别工时月份"))
    if parsed.total_hours is None:
        items.append(("MISSING_HOURS", "无法可靠识别总工时"))
    if not parsed.po_id:
        items.append(("MISSING_PO", "无法可靠识别 PO 号，需要确认 PO 归属"))
    if parsed.parse_method == "ocr":
        items.append(("OCR_REVIEW", "记录来自本地 OCR，需要人工确认识别结果"))
    if parsed.ocr_error:
        items.append(("OCR_ERROR", parsed.ocr_error))
    validation_status = str(parsed.validation.get("status", "")).upper()
    validation_issues = parsed.validation.get("issues") or []
    blocking_issues = [
        issue for issue in validation_issues
        if str(issue.get("severity", "")).upper() == "BLOCKING"
    ]
    if validation_status == "BLOCKED" or blocking_issues:
        items.append(("VALIDATION_BLOCKED", "字段校验存在阻断问题，需要人工确认"))
    elif validation_status == "WARNING" or validation_issues:
        items.append(("VALIDATION_REVIEW", "字段校验存在警告，需要人工复核"))
    if existing:
        items.append(("VERSION_CONFLICT", "同一人员-月份-PO 已存在不同 PDF 版本，禁止自动相加"))
    return items


def _refresh_candidate_reviews(
    store: FactStore,
    project_id: str,
    candidate_id: str,
    run_id: str,
    version_key: str,
    parsed: ParsedTimesheet,
    candidate: dict[str, Any],
) -> list[tuple[str, str]]:
    store.resolve_candidate_reviews(candidate_id, "解析结果更新，旧队列重新计算")
    existing = [row for row in store.find_candidates(project_id, version_key) if row["candidate_id"] != candidate_id]
    review_types = _review_types(parsed, existing)
    for queue_type, message in review_types:
        store.add_review_item(
            project_id, run_id, queue_type,
            payload={"message": message, "candidate": candidate, "source_path": parsed.path},
            candidate_id=candidate_id, target_key=version_key,
            severity="BLOCKING" if queue_type in {"NEW_MEMBER", "MISSING_PO", "MISSING_HOURS", "VERSION_CONFLICT", "VALIDATION_BLOCKED"} else "WARNING",
        )
    return review_types


def scan_project(
    project_id: str,
    use_ocr: bool = True,
    include_archive: bool = True,
    limit: int | None = None,
) -> dict[str, Any]:
    profile = build_profile(project_id)
    timesheet_dir = profile.source_dirs.get("timesheet_pdf", "")
    root = Path(timesheet_dir)
    if not root.exists():
        raise FileNotFoundError(timesheet_dir)
    pdfs = sorted(root.rglob("*.pdf"), key=lambda p: str(p).casefold())
    if not include_archive:
        pdfs = [p for p in pdfs if not re.search(r"archive|resign", str(p), re.I)]
    if limit is not None:
        pdfs = pdfs[:limit]

    with FactStore() as store:
        store.initialize()
        store.register_project(profile)
        run_id = store.create_run(project_id, source_cutoff=datetime_now(), notes="timesheet PDF incremental scan")
        summary = {
            "project_id": project_id, "run_id": run_id, "scanned": len(pdfs),
            "new": 0, "exact_duplicates": 0, "candidates": 0, "review_items": 0,
            "methods": {}, "review_types": {}, "errors": [],
        }
        for pdf_path in pdfs:
            try:
                fingerprint = store.fingerprint_file(pdf_path)
                existing_source = store.find_source_by_hash(project_id, fingerprint["sha256"])
                if existing_source:
                    summary["exact_duplicates"] += 1
                    continue
                source_id = store.record_source_file(
                    project_id, "timesheet_pdf", pdf_path, run_id=run_id,
                    metadata={"scan_scope": str(root), "archive_path": bool(re.search(r"archive|resign", str(pdf_path), re.I))},
                )
                parsed = parse_pdf(pdf_path, use_ocr=use_ocr)
                version_key = _version_key(parsed)
                existing_candidates = store.find_candidates(project_id, version_key)
                review_types = _review_types(parsed, existing_candidates)
                status = "REVIEW" if review_types else "PENDING"
                candidate = asdict(parsed)
                candidate.update({
                    "project_id": project_id, "run_id": run_id, "source_id": source_id,
                    "version_key": version_key, "review_status": status,
                    "raw_unit": "hour", "parse_json": {
                        "fields": parsed.fields,
                        "text_excerpt": parsed.text_excerpt,
                        "ocr_diagnostics": parsed.ocr_diagnostics,
                        "evidence": parsed.evidence,
                        "validation": parsed.validation,
                    },
                })
                candidate_id = store.insert_timesheet_candidate(candidate)
                summary["new"] += 1
                summary["candidates"] += 1
                summary["methods"][parsed.parse_method] = summary["methods"].get(parsed.parse_method, 0) + 1
                for queue_type, message in review_types:
                    store.add_review_item(
                        project_id, run_id, queue_type,
                        payload={"message": message, "candidate": candidate, "source_path": str(pdf_path)},
                        candidate_id=candidate_id, target_key=version_key,
                        severity="BLOCKING" if queue_type in {"NEW_MEMBER", "MISSING_PO", "MISSING_HOURS", "VERSION_CONFLICT", "VALIDATION_BLOCKED"} else "WARNING",
                    )
                    summary["review_items"] += 1
                    summary["review_types"][queue_type] = summary["review_types"].get(queue_type, 0) + 1
            except Exception as exc:
                summary["errors"].append({"path": str(pdf_path), "error": f"{type(exc).__name__}: {exc}"})
        return summary


def ocr_pending_candidates(
    project_id: str,
    limit: int | None = None,
    current_month_only: bool = False,
    candidate_ids: Iterable[str] | None = None,
) -> dict[str, Any]:
    """对当前审核范围的候选补做/重做 OCR，允许修正之前不完整的 OCR 结果。"""
    with FactStore() as store:
        store.initialize()
        rows = list(store.connection.execute(
            """SELECT c.*, s.path AS source_path FROM timesheet_candidates c
               JOIN source_files s ON s.source_id=c.source_id
               WHERE c.project_id=? AND c.parse_method IN ('filename_only', 'ocr')
                 AND c.review_status NOT IN ('APPROVED', 'SUPERSEDED', 'DISMISSED')
               ORDER BY c.created_at""",
            (project_id,),
        ))
        requested_ids = {str(value) for value in (candidate_ids or []) if str(value).strip()}
        if requested_ids:
            rows = [row for row in rows if row["candidate_id"] in requested_ids]
        if current_month_only:
            scope = review_scope(project_id)
            rows = [
                row for row in rows
                if row["year"] == scope["year"]
                and row["month"] == scope["month_number"]
                and path_in_scope(row["source_path"], scope)
            ]
        if limit is not None:
            rows = rows[:limit]
        result = {"project_id": project_id, "pending": len(rows), "processed": 0,
                  "ocr_success": 0, "still_pending": 0, "errors": []}
        for row in rows:
            try:
                parsed = parse_pdf(row["source_path"], use_ocr=True)
                candidate = asdict(parsed)
                version_key = _version_key(parsed)
                candidate["parse_json"] = {
                    "fields": parsed.fields, "text_excerpt": parsed.text_excerpt,
                    "ocr_error": parsed.ocr_error,
                    "ocr_diagnostics": parsed.ocr_diagnostics,
                    "evidence": parsed.evidence,
                    "validation": parsed.validation,
                }
                candidate.update({"version_key": version_key, "review_status": "REVIEW"})
                store.update_timesheet_candidate(row["candidate_id"], candidate)
                review_types = _refresh_candidate_reviews(
                    store, project_id, row["candidate_id"], row["run_id"],
                    version_key, parsed, candidate,
                )
                if not review_types:
                    store.add_review_item(
                        project_id, row["run_id"], "OCR_REVIEW",
                        payload={"message": "补充 OCR 后需人工确认", "candidate": candidate,
                                 "source_path": row["source_path"]},
                        candidate_id=row["candidate_id"], target_key=version_key,
                        severity="WARNING",
                    )
                result["processed"] += 1
                if parsed.parse_method == "ocr":
                    result["ocr_success"] += 1
                else:
                    result["still_pending"] += 1
            except Exception as exc:
                result["errors"].append({"path": row["source_path"], "error": f"{type(exc).__name__}: {exc}"})
        return result


def datetime_now() -> str:
    from datetime import datetime, timezone
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Incremental timesheet PDF ingest")
    sub = parser.add_subparsers(dest="command", required=True)
    scan = sub.add_parser("scan")
    scan.add_argument("--project", default="CNTS-25165")
    scan.add_argument("--no-ocr", action="store_true")
    scan.add_argument("--exclude-archive", action="store_true")
    scan.add_argument("--limit", type=int, default=None)
    scan.add_argument("--json", action="store_true")
    ocr = sub.add_parser("ocr-pending")
    ocr.add_argument("--project", default="CNTS-25165")
    ocr.add_argument("--limit", type=int, default=None)
    ocr.add_argument("--current-month", action="store_true")
    ocr.add_argument("--candidate-id", action="append", default=[])
    ocr.add_argument("--json", action="store_true")
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    if args.command == "scan":
        result = scan_project(
            args.project, use_ocr=not args.no_ocr,
            include_archive=not args.exclude_archive, limit=args.limit,
        )
        print(json.dumps(result, ensure_ascii=False, indent=2) if args.json else result)
        return 0 if not result["errors"] else 1
    if args.command == "ocr-pending":
        result = ocr_pending_candidates(
            args.project, limit=args.limit, current_month_only=args.current_month,
            candidate_ids=args.candidate_id,
        )
        print(json.dumps(result, ensure_ascii=False, indent=2) if args.json else result)
        return 0 if not result["errors"] else 1
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
