# -*- coding: utf-8 -*-
"""Scope rules for the monthly timesheet review gate."""

from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from project_registry import build_profile


def latest_closed_month(today: date | None = None) -> str:
    """Return the previous calendar month as ``YYYY-MM``."""
    current = today or datetime.now().date()
    first_of_current = current.replace(day=1)
    previous = first_of_current - timedelta(days=1)
    return previous.strftime("%Y-%m")


def review_scope(project_id: str, today: date | None = None) -> dict[str, Any]:
    month = latest_closed_month(today)
    year, month_number = (int(part) for part in month.split("-", 1))
    profile = build_profile(project_id)
    base = Path(profile.source_dirs.get("timesheet_pdf", "")).resolve()
    # CNTS-25165 stores monthly PDF folders as YYMM, e.g. 2607.  Keep a
    # YYYYMM fallback for projects that use the long folder convention.
    short_dir = base / f"{year % 100:02d}{month_number:02d}"
    long_dir = base / f"{year:04d}{month_number:02d}"
    source_dir = short_dir if short_dir.exists() or not long_dir.exists() else long_dir
    return {
        "project": project_id,
        "month": month,
        "year": year,
        "month_number": month_number,
        "source_dir": str(source_dir),
        "source_dir_exists": source_dir.is_dir(),
        "historical_rule": "Before the review month, use Budget Report and Monthly Closing; PDF review is limited to the latest closed month.",
    }


def path_in_scope(path: str | None, scope: dict[str, Any]) -> bool:
    if not path:
        return False
    try:
        candidate = Path(path).resolve()
        root = Path(scope["source_dir"]).resolve()
        return candidate.is_relative_to(root)
    except (OSError, ValueError):
        return False
