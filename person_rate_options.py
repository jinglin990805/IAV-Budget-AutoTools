# -*- coding: utf-8 -*-
"""Build person/Level/customer-rate choices for timesheet review.

The review page should use existing customer-facing tables as the first source
of truth.  Local configuration and previously confirmed facts are fallbacks;
the source workbooks are read only and are never rewritten here.
"""

from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

import openpyxl

import config


def _text(value: Any) -> str:
    return str(value).strip() if value is not None else ""


def _key(value: Any) -> str:
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]", "", _text(value).lower())


def _canonical_table_name(value: Any) -> str:
    """Convert common Excel ``Last, First`` display names to our canonical form."""
    name = re.sub(r"\s+", " ", _text(value))
    if "," in name:
        last, first = (part.strip() for part in name.split(",", 1))
        if last and first:
            return f"{last} {first}"
    return name


def _usage_files(project_id: str) -> list[Path]:
    project = config.PROJECTS.get(project_id, {})
    directory = project.get("po_usage_dir") or project.get("budget_report_dir")
    if not directory:
        return []
    root = Path(directory)
    if not root.exists():
        return []
    found = []
    for path in root.rglob("*.xlsx"):
        name = path.name.lower()
        if "po usage status" not in name and "po_usage_status" not in name:
            continue
        if any(token in name for token in ("working", "backup", "~$")):
            continue
        found.append(path)
    return sorted(found, key=lambda item: (item.stat().st_mtime_ns, str(item).lower()), reverse=True)


def _read_usage_people(project_id: str) -> tuple[dict[str, dict[str, Any]], list[str]]:
    people: dict[str, dict[str, Any]] = {}
    source_files: list[str] = []
    # Use the newest customer-facing file.  Older v2/v3 snapshots can contain
    # people who already left the project and should not silently appear in the
    # current review dropdown.
    for path in _usage_files(project_id)[:1]:
        try:
            workbook = openpyxl.load_workbook(
                path, data_only=True, read_only=True, keep_links=True
            )
            source_files.append(str(path))
            for sheet in workbook.worksheets:
                header_row = None
                name_column = None
                level_column = None
                for row_number, row in enumerate(
                    sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 20), values_only=True), 1
                ):
                    for index, value in enumerate(row):
                        normalized = _key(value)
                        if normalized in {"fte", "employee", "person", "人员", "员工"}:
                            name_column = index
                        if normalized in {"role", "level", "岗位", "级别"}:
                            level_column = index
                    if name_column is not None and level_column is not None:
                        header_row = row_number
                        break
                if header_row is None or name_column is None:
                    continue
                for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
                    if name_column >= len(row):
                        continue
                    name = _canonical_table_name(row[name_column])
                    if not name or _key(name) in {"total", "subtotal", "fte"}:
                        continue
                    level = _text(row[level_column]) if level_column < len(row) else ""
                    if not level:
                        continue
                    entry = people.setdefault(
                        name,
                        {"canonical_name": name, "aliases": set(), "levels": Counter(), "source": "PO Usage Status"},
                    )
                    entry["levels"][level] += 1
                    entry["aliases"].add(_text(row[name_column]))
            workbook.close()
        except Exception:
            # A damaged/locked auxiliary workbook must not prevent the review
            # page from using configuration and database fallbacks.
            continue
    return people, source_files


def _read_engineer_rates() -> tuple[dict[str, float], dict[str, float], list[str]]:
    """Return VCTC level rates, all level rates, and available numeric rates."""
    path = Path(config.INPUT_DIR) / "engineer level.xlsx"
    vctc: dict[str, float] = {}
    all_levels: dict[str, float] = {}
    all_rates: set[float] = set()
    if not path.exists():
        return vctc, all_levels, []
    try:
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
        sheet = workbook.active
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if len(row) < 3:
                continue
            customer, level, raw_rate = _text(row[0]), _text(row[1]), row[2]
            if not level or raw_rate in (None, ""):
                continue
            try:
                rate = float(raw_rate)
            except (TypeError, ValueError):
                continue
            all_levels.setdefault(level, rate)
            all_rates.add(rate)
            if _key(customer) == "vctc":
                vctc[level] = rate
        workbook.close()
    except Exception:
        return vctc, all_levels, []
    return vctc, all_levels, sorted(all_rates)


def _add_person(
    people: dict[str, dict[str, Any]],
    canonical_name: str,
    level: str,
    aliases: list[str] | None = None,
    source: str = "configuration",
) -> None:
    name = _canonical_table_name(canonical_name)
    if not name:
        return
    item = people.setdefault(
        name,
        {"canonical_name": name, "aliases": set(), "levels": Counter(), "source": source},
    )
    if level:
        item["levels"][level] += 1
    item["aliases"].update(_text(alias) for alias in (aliases or []) if _text(alias))
    if item.get("source") == "configuration" and source == "PO Usage Status":
        item["source"] = source


def build_review_options(project_id: str, connection: Any | None = None) -> dict[str, Any]:
    """Return dropdown data and automatic person -> Level -> rate mappings."""
    people, source_files = _read_usage_people(project_id)

    personnel = getattr(config, f"{project_id.replace('-', '')}_PERSONNEL", {})
    if project_id == "CNTS-25165":
        personnel = getattr(config, "CNTS25165_PERSONNEL", {})
    for name, info in personnel.items():
        _add_person(
            people, name, _text(info.get("level")), info.get("aliases", []), "configuration"
        )

    if connection is not None:
        for row in connection.execute(
            """SELECT p.canonical_name, p.aliases_json, r.level
               FROM person_master p
               JOIN project_person_rates r ON r.person_key=p.person_key
               WHERE r.project_id=? ORDER BY p.canonical_name""",
            (project_id,),
        ):
            try:
                aliases = json.loads(row["aliases_json"] or "[]")
            except json.JSONDecodeError:
                aliases = []
            _add_person(people, row["canonical_name"], _text(row["level"]), aliases, "confirmed")

    vctc_rates, all_level_rates, all_rates = _read_engineer_rates()
    project_rates = dict(vctc_rates)
    project_rates.update({
        _text(level): float(rate)
        for level, rate in getattr(config, "VCTC_RATES", {}).items()
    })
    # If the project has a customer-specific rate card, do not expose unrelated
    # CARIAD/other-customer roles in the review dropdown.
    levels = set(project_rates) if project_rates else set(all_level_rates)
    for item in people.values():
        if item["levels"]:
            levels.update(item["levels"])

    result_people = []
    for name in sorted(people, key=str.casefold):
        item = people[name]
        level = item["levels"].most_common(1)[0][0] if item["levels"] else ""
        rate = project_rates.get(level) or all_level_rates.get(level)
        result_people.append({
            "canonical_name": name,
            "aliases": sorted(item["aliases"]),
            "level": level,
            "customer_day_rate": rate,
            "source": item["source"],
        })

    level_options = [
        {"value": level, "customer_day_rate": project_rates.get(level) or all_level_rates.get(level)}
        for level in sorted(levels, key=str.casefold)
    ]
    rate_values = sorted(set(project_rates.values()) if project_rates else set(all_rates))
    return {
        "project": project_id,
        "people": result_people,
        "levels": level_options,
        "rates": [{"value": rate, "label": f"¥{rate:g}/天"} for rate in rate_values],
        "source_files": source_files,
    }
