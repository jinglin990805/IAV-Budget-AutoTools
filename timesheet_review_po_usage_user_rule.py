# -*- coding: utf-8 -*-
"""Apply the user's authoritative historical Timesheet handling rule.

The PO Usage Status workbook is authoritative through the historical cutoff:

* duplicate PDF versions keep a signed version when one exists;
* a person/month with multiple workbook PO rows is represented by one
  candidate per workbook PO row;
* PO and hours on retained/derived candidates are overwritten from Excel.

The command is dry-run by default.  ``--apply`` only changes Timesheet
candidate/review data.  It never creates an approved Timesheet fact.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any

from fact_store import FactStore
from timesheet_review_po_usage_reconcile import (
    DATA_SOURCE,
    PROJECT,
    _digits,
    _float_equal,
    _update_payload,
    _usage_for,
    _usage_people,
    _workbook_name,
    load_usage_rows,
)


def _candidate_ids(csv_path: str | Path) -> set[str]:
    with Path(csv_path).open("r", encoding="utf-8-sig", newline="") as handle:
        return {
            str(row.get("candidate_id") or "").strip()
            for row in csv.DictReader(handle)
            if str(row.get("candidate_id") or "").strip()
        }


def _hours(value: Any) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _hour_text(value: Any) -> str:
    number = float(value)
    return str(int(number)) if number.is_integer() else str(number)


def _source_priority(row: dict[str, Any]) -> tuple[int, int, int, float, str, str]:
    """Return a deterministic version preference key.

    ``signed`` is the explicit user rule.  The remaining fields only resolve
    groups where no signed copy exists: updated copy, OCR confidence, and a
    non-large render.  The path/id tie-breakers make the operation repeatable.
    """

    file_name = str(row.get("source_file") or "").casefold()
    path = str(row.get("source_path") or "").casefold()
    signed = int("signed" in file_name or "signed" in path)
    updated = int("update" in file_name or "update" in path)
    non_large = int("large" not in file_name and "large" not in path)
    confidence = float(row.get("confidence") or 0.0)
    # Ascending path/id is used after the semantic score.  It keeps identical
    # imported copies stable without pretending Excel can distinguish them.
    return signed, updated, non_large, confidence, path, str(row.get("candidate_id") or "")


def _load_candidates(candidate_ids: set[str]) -> list[dict[str, Any]]:
    if not candidate_ids:
        return []
    placeholders = ",".join("?" for _ in candidate_ids)
    with FactStore() as store:
        store.initialize()
        rows = store.connection.execute(
            f"""SELECT c.*, s.file_name AS source_file, s.path AS source_path
                FROM timesheet_candidates c
                LEFT JOIN source_files s ON s.source_id=c.source_id
                WHERE c.project_id=?
                  AND c.review_status NOT IN ('SUPERSEDED','APPROVED','DISMISSED')
                  AND c.candidate_id IN ({placeholders})""",
            (PROJECT, *sorted(candidate_ids)),
        ).fetchall()
    return [dict(row) for row in rows]


def _pair_score(row: dict[str, Any], entry: dict[str, Any]) -> tuple[int, int, float]:
    candidate_hours = _hours(row.get("total_hours"))
    target_hours = _hours(entry.get("hours")) or 0.0
    exact_hours = int(candidate_hours is not None and _float_equal(candidate_hours, target_hours))
    exact_po = int(_digits(row.get("po_id")) == _digits(entry.get("po_id")))
    distance = abs((candidate_hours if candidate_hours is not None else 0.0) - target_hours)
    return exact_hours, exact_po, distance


def _best_assignment(
    candidates: list[dict[str, Any]],
    entries: list[dict[str, Any]],
) -> dict[str, int]:
    """Match each candidate to at most one workbook entry.

    Hours matching is the strongest evidence because the user explicitly
    requested Excel-hour overwrite.  PO matching is next, followed by source
    version preference.  A small exhaustive search is appropriate here: the
    historical groups contain at most a few versions/PO rows.
    """

    ordered = sorted(candidates, key=lambda row: _source_priority(row), reverse=True)
    best_key: tuple[Any, ...] | None = None
    best_mapping: dict[str, int] = {}
    row_by_id = {row["candidate_id"]: row for row in ordered}

    def visit(index: int, used: set[int], mapping: dict[str, int]) -> None:
        nonlocal best_key, best_mapping
        if index >= len(ordered):
            pairs = [
                (row_by_id[candidate_id], entries[target_index])
                for candidate_id, target_index in mapping.items()
            ]
            metrics = [
                _pair_score(row, entry)
                for row, entry in pairs
            ]
            signed = sum(_source_priority(row)[0] for row, _ in pairs)
            updated = sum(_source_priority(row)[1] for row, _ in pairs)
            non_large = sum(_source_priority(row)[2] for row, _ in pairs)
            confidence = sum(_source_priority(row)[3] for row, _ in pairs)
            key = (
                len(pairs),
                sum(metric[0] for metric in metrics),
                sum(metric[1] for metric in metrics),
                -sum(metric[2] for metric in metrics),
                signed,
                updated,
                non_large,
                confidence,
            )
            if best_key is None or key > best_key:
                best_key = key
                best_mapping = dict(mapping)
            return

        row = ordered[index]
        # Leave this version unused when there are more candidate versions than
        # workbook rows.  The assignment score prefers a full matching set.
        visit(index + 1, used, mapping)
        for target_index in range(len(entries)):
            if target_index in used:
                continue
            mapping[row["candidate_id"]] = target_index
            used.add(target_index)
            visit(index + 1, used, mapping)
            used.remove(target_index)
            mapping.pop(row["candidate_id"], None)

    visit(0, set(), {})
    return best_mapping


def _complete_exact(candidates: list[dict[str, Any]], entries: list[dict[str, Any]]) -> bool:
    remaining = list(candidates)
    for entry in entries:
        matches = [
            row for row in remaining
            if _digits(row.get("po_id")) == _digits(entry.get("po_id"))
            and _float_equal(row.get("total_hours"), entry.get("hours"))
        ]
        if len(matches) != 1:
            return False
        remaining.remove(matches[0])
    return True


def _is_full_month_version(row: dict[str, Any], entries: list[dict[str, Any]]) -> bool:
    total = sum(float(entry["hours"]) for entry in entries)
    return _float_equal(row.get("total_hours"), total)


def _preferred(rows: list[dict[str, Any]]) -> dict[str, Any]:
    return sorted(rows, key=_source_priority, reverse=True)[0]


def _build_plan(workbook_path: str | Path, candidate_csv: str | Path) -> dict[str, Any]:
    usage = load_usage_rows(workbook_path)
    people = _usage_people(usage)
    rows = _load_candidates(_candidate_ids(candidate_csv))
    groups: dict[tuple[str, int, int], list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        raw_person = row.get("canonical_name") or row.get("raw_name") or ""
        person = _workbook_name(raw_person, people)
        if row.get("year") is None or row.get("month") is None:
            continue
        groups[(person, int(row["year"]), int(row["month"]))].append(row)

    updates: list[dict[str, Any]] = []
    supersedes: list[dict[str, Any]] = []
    derived: list[dict[str, Any]] = []
    unresolved: list[dict[str, Any]] = []
    processed_groups: list[dict[str, Any]] = []

    for (person, year, month), group in sorted(groups.items()):
        entries = _usage_for(usage, person, year, month)
        if not entries:
            continue
        entries = list(entries)
        period = f"{year:04d}-{month:02d}"
        if not group:
            unresolved.append({"person": person, "period": period, "reason": "no active candidate"})
            continue

        assignment: dict[str, int] = {}
        merge_source: dict[str, Any] | None = None
        mode = "version selection"

        # A full-month candidate is a merged multi-PO version.  Use it when
        # there is no already-complete one-to-one PO/hour set.  This catches
        # both one-file merges and duplicate copies of a merged file.
        if len(entries) > 1 and not _complete_exact(group, entries):
            full_versions = [row for row in group if _is_full_month_version(row, entries)]
            if full_versions:
                merge_source = _preferred(full_versions)
                matching_index = next(
                    (
                        index for index, entry in enumerate(entries)
                        if _digits(entry.get("po_id")) == _digits(merge_source.get("po_id"))
                    ),
                    0,
                )
                assignment[merge_source["candidate_id"]] = matching_index
                mode = "merged multi-PO split from full-month candidate"
            elif len(group) == 1:
                merge_source = group[0]
                matching_index = next(
                    (
                        index for index, entry in enumerate(entries)
                        if _digits(entry.get("po_id")) == _digits(merge_source.get("po_id"))
                    ),
                    None,
                )
                if matching_index is None:
                    hour_matches = [
                        index for index, entry in enumerate(entries)
                        if _float_equal(merge_source.get("total_hours"), entry.get("hours"))
                    ]
                    matching_index = hour_matches[0] if len(hour_matches) == 1 else 0
                assignment[merge_source["candidate_id"]] = matching_index
                mode = "merged multi-PO split from single candidate"

        if not assignment:
            assignment = _best_assignment(group, entries)

        if not assignment:
            unresolved.append({
                "person": person,
                "period": period,
                "reason": "could not assign an active candidate to workbook rows",
                "candidate_ids": [row["candidate_id"] for row in group],
            })
            continue

        target_by_id = {
            candidate_id: entries[target_index]
            for candidate_id, target_index in assignment.items()
        }
        assigned_target_indexes = set(assignment.values())
        retained_ids = set(assignment)

        for row in group:
            candidate_id = row["candidate_id"]
            if candidate_id in retained_ids:
                target = target_by_id[candidate_id]
                updates.append({
                    "candidate_id": candidate_id,
                    "person": person,
                    "period": period,
                    "new_po": str(target["po_id"]),
                    "new_hours": float(target["hours"]),
                    "source_file": row.get("source_file") or "",
                    "old_po": row.get("po_id") or "",
                    "old_hours": row.get("total_hours") or "",
                    "mode": mode,
                    "workbook_row": target.get("row"),
                })
            else:
                same_po = [
                    retained_id
                    for retained_id in retained_ids
                    if _digits(target_by_id[retained_id].get("po_id")) == _digits(row.get("po_id"))
                ]
                replacement_id = sorted(same_po or retained_ids)[0] if (same_po or retained_ids) else ""
                supersedes.append({
                    "candidate_id": candidate_id,
                    "replacement_id": replacement_id,
                    "person": person,
                    "period": period,
                    "source_file": row.get("source_file") or "",
                    "replacement_file": next(
                        (item.get("source_file") or "" for item in group if item["candidate_id"] == replacement_id),
                        "",
                    ),
                    "reason": "duplicate/version removed by user rule",
                })

        # Every workbook row must have a candidate.  Missing rows are derived
        # from the retained source while preserving source_id and audit trail.
        template = merge_source or _preferred([row for row in group if row["candidate_id"] in retained_ids])
        for target_index, target in enumerate(entries):
            if target_index in assigned_target_indexes:
                continue
            derived_id = f"tsc-wbs-{template['candidate_id'][4:16]}-{_digits(target['po_id'])}"
            derived.append({
                "candidate_id": derived_id,
                "from_candidate_id": template["candidate_id"],
                "person": person,
                "period": period,
                "new_po": str(target["po_id"]),
                "new_hours": float(target["hours"]),
                "source_file": template.get("source_file") or "",
                "workbook_row": target.get("row"),
                "mode": mode,
            })

        processed_groups.append({
            "person": person,
            "period": period,
            "mode": mode,
            "workbook": [
                {"po_id": str(entry["po_id"]), "hours": float(entry["hours"])}
                for entry in entries
            ],
            "retained": [
                {
                    "candidate_id": row["candidate_id"],
                    "source_file": row.get("source_file") or "",
                    "assigned_po": target_by_id[row["candidate_id"]]["po_id"],
                    "assigned_hours": target_by_id[row["candidate_id"]]["hours"],
                }
                for row in group if row["candidate_id"] in retained_ids
            ],
        })

    return {
        "candidate_rows": len(rows),
        "groups": len(processed_groups),
        "updates": updates,
        "supersedes": supersedes,
        "derived": derived,
        "unresolved": unresolved,
        "processed_groups": processed_groups,
    }


def _version_key(row: dict[str, Any], po_id: str, hours: float, suffix: str = "") -> str:
    base = str(row.get("version_key") or row.get("candidate_id") or "candidate")
    tail = f"::workbook::{_digits(po_id)}::{_hour_text(hours)}"
    return f"{base}{tail}{suffix}"


def _update_existing(store: FactStore, item: dict[str, Any], row: dict[str, Any]) -> None:
    po_id = str(item["new_po"])
    hours = float(item["new_hours"])
    payload = _update_payload(row.get("parse_json"), po_id, hours, item["person"])
    note = (
        f"{DATA_SOURCE} user rule: {item['period']} {item['person']} -> "
        f"PO {po_id}, {hours:g}h; Excel overrides PO/hours; mode={item['mode']}; "
        f"source={item['source_file']}."
    )
    version_key = _version_key(row, po_id, hours)
    store.connection.execute(
        """UPDATE timesheet_candidates SET
               canonical_name=?, po_id=?, total_hours=?, parse_json=?,
               version_key=?, review_status='REVIEW', is_preferred=1, notes=?
           WHERE candidate_id=? AND project_id=?""",
        (item["person"], po_id, hours, payload, version_key, note, item["candidate_id"], PROJECT),
    )
    store.connection.execute(
        """UPDATE review_queue
           SET status='RESOLVED', resolved_at=datetime('now'), resolution_note=?
           WHERE candidate_id=? AND project_id=? AND status='OPEN'
             AND queue_type IN ('NEW_MEMBER','MISSING_PO','UNKNOWN_PO','MISSING_HOURS','VALIDATION_BLOCKED')""",
        (note, item["candidate_id"], PROJECT),
    )


def _apply_plan(plan: dict[str, Any]) -> dict[str, int]:
    applied_updates = 0
    applied_supersedes = 0
    applied_derived = 0
    with FactStore() as store:
        store.initialize()
        row_ids = {
            item["candidate_id"]
            for item in plan["updates"]
        }
        placeholders = ",".join("?" for _ in row_ids)
        rows = {}
        if row_ids:
            rows = {
                row["candidate_id"]: dict(row)
                for row in store.connection.execute(
                    f"SELECT * FROM timesheet_candidates WHERE project_id=? AND candidate_id IN ({placeholders})",
                    (PROJECT, *sorted(row_ids)),
                )
            }
        for item in plan["updates"]:
            row = rows.get(item["candidate_id"])
            if row is None:
                continue
            _update_existing(store, item, row)
            applied_updates += 1

        for item in plan["supersedes"]:
            store.supersede_candidate(
                item["candidate_id"],
                item["replacement_id"],
                f"{DATA_SOURCE} user rule: superseded duplicate/version for "
                f"{item['person']} {item['period']}; keep retained workbook-aligned candidate.",
            )
            applied_supersedes += 1

        template_ids = {item["from_candidate_id"] for item in plan["derived"]}
        template_rows = {}
        if template_ids:
            placeholders = ",".join("?" for _ in template_ids)
            template_rows = {
                row["candidate_id"]: dict(row)
                for row in store.connection.execute(
                    f"SELECT * FROM timesheet_candidates WHERE project_id=? AND candidate_id IN ({placeholders})",
                    (PROJECT, *sorted(template_ids)),
                )
            }
        for item in plan["derived"]:
            template = template_rows.get(item["from_candidate_id"])
            if template is None:
                continue
            exists = store.connection.execute(
                "SELECT 1 FROM timesheet_candidates WHERE candidate_id=?",
                (item["candidate_id"],),
            ).fetchone()
            if exists:
                continue
            po_id = str(item["new_po"])
            hours = float(item["new_hours"])
            derived = dict(template)
            derived.update({
                "candidate_id": item["candidate_id"],
                "canonical_name": item["person"],
                "po_id": po_id,
                "total_hours": hours,
                "raw_unit": "hour",
                "parse_method": "workbook_override",
                "version_key": _version_key(template, po_id, hours, "::split"),
                "parse_json": json.loads(_update_payload(template.get("parse_json"), po_id, hours, item["person"])),
                "review_status": "REVIEW",
                "is_preferred": 1,
            })
            note = (
                f"{DATA_SOURCE} user rule: derived split candidate from "
                f"{item['from_candidate_id']} for {item['person']} {item['period']} -> "
                f"PO {po_id}, {hours:g}h; source={item['source_file']}."
            )
            store.insert_timesheet_candidate(derived)
            store.connection.execute(
                "UPDATE timesheet_candidates SET notes=? WHERE candidate_id=?",
                (note, item["candidate_id"]),
            )
            applied_derived += 1
        store.connection.commit()
    return {
        "updates": applied_updates,
        "supersedes": applied_supersedes,
        "derived": applied_derived,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Apply the user's Excel-authoritative historical Timesheet rule")
    parser.add_argument("--workbook", required=True)
    parser.add_argument("--candidate-csv", required=True)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args(argv)

    plan = _build_plan(args.workbook, args.candidate_csv)
    applied = _apply_plan(plan) if args.apply else {"updates": 0, "supersedes": 0, "derived": 0}
    result = {
        "ok": not plan["unresolved"],
        "project": PROJECT,
        "workbook": str(args.workbook),
        "apply": args.apply,
        "fact_approval_performed": False,
        "candidate_rows": plan["candidate_rows"],
        "groups": plan["groups"],
        "update_plan": len(plan["updates"]),
        "supersede_plan": len(plan["supersedes"]),
        "derived_plan": len(plan["derived"]),
        "unresolved": plan["unresolved"],
        "applied": applied,
        "processed_groups": plan["processed_groups"],
        "updates": plan["updates"],
        "supersedes": plan["supersedes"],
        "derived": plan["derived"],
    }
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
