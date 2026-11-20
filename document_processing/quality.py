"""Provider-neutral OCR quality gates for retry and review routing.

The gate evaluates OCR completeness only.  It does not decide business values,
approve facts, or replace the domain validators used by Timesheet and Travel.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterable


@dataclass(frozen=True)
class OCRQualityReport:
    status: str
    character_count: int
    matched_markers: tuple[str, ...]
    missing_markers: tuple[str, ...]
    reasons: tuple[str, ...]
    score: float
    confidence: float | None = None

    @property
    def should_retry(self) -> bool:
        return self.status == "RETRY"

    def as_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "character_count": self.character_count,
            "matched_markers": list(self.matched_markers),
            "missing_markers": list(self.missing_markers),
            "reasons": list(self.reasons),
            "score": self.score,
            "confidence": self.confidence,
            "should_retry": self.should_retry,
        }


def assess_ocr_text(
    text: str,
    *,
    min_chars: int = 80,
    required_markers: Iterable[str] = (),
    confidence: float | None = None,
    min_confidence: float | None = None,
) -> OCRQualityReport:
    """Return a deterministic retry decision for one OCR text candidate.

    ``required_markers`` are route-level anchors such as ``Actual`` and
    ``Purchase Order``.  Business-field correctness remains the responsibility
    of the existing validators after parsing.
    """

    compact = re.sub(r"\s+", "", text or "")
    character_count = len(compact)
    markers = tuple(str(item) for item in required_markers if str(item).strip())
    normalised = (text or "").casefold()
    matched = tuple(marker for marker in markers if marker.casefold() in normalised)
    missing = tuple(marker for marker in markers if marker not in matched)
    reasons: list[str] = []
    minimum = max(1, int(min_chars))
    if not compact:
        reasons.append("EMPTY_TEXT")
    elif character_count < minimum:
        reasons.append("SHORT_TEXT")
    if missing:
        reasons.append("MISSING_MARKER")
    if min_confidence is not None and (
        confidence is None or float(confidence) < float(min_confidence)
    ):
        reasons.append("LOW_CONFIDENCE")

    length_score = min(1.0, character_count / minimum) if minimum else 1.0
    marker_score = len(matched) / len(markers) if markers else 1.0
    confidence_score = (
        min(1.0, max(0.0, float(confidence))) if confidence is not None else 0.0
    )
    score = round(length_score * 0.45 + marker_score * 0.45 + confidence_score * 0.10, 4)
    return OCRQualityReport(
        status="PASS" if not reasons else "RETRY",
        character_count=character_count,
        matched_markers=matched,
        missing_markers=missing,
        reasons=tuple(reasons),
        score=score,
        confidence=confidence,
    )


def ocr_candidate_rank(
    text: str,
    *,
    min_chars: int = 80,
    required_markers: Iterable[str] = (),
    confidence: float | None = None,
) -> tuple[int, int, int, float, int]:
    """Return a stable ranking tuple for selecting among OCR attempts."""

    report = assess_ocr_text(
        text,
        min_chars=min_chars,
        required_markers=required_markers,
        confidence=confidence,
    )
    return (
        int(report.status == "PASS"),
        len(report.matched_markers),
        min(report.character_count, max(1, int(min_chars))),
        report.score,
        report.character_count,
    )
