# -*- coding: utf-8 -*-
"""Keep only Timesheet candidate versions matching authoritative PO usage rows."""

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
    _float_equal,
    _usage_for,
    _workbook_name,
    load_usage_rows,
)


PROJECT = "CNTS-25165"
DATA_SOURCE = "PO Usage Status_26917.xlsx / PO Overview Actual_To Customer"


def _digits(value: Any) -> str:
    return re.sub(r"\D", "", str(value or "")).lstrip("0")


def _candidate_ids(csv_path: str | Path) -> set[str]:
    with Path(csv_path).open("r", encoding="utf-8-sig", newline="") as handle:
        return {
            str(row.get("candidate_id") or "").strip()
            for row in csv.DictReader(handle)
            if str(row.get("candidate_id") or "").strip()
        }


def build_plan(workbook_path: str | Path, csv_path: str | Path) -> dict[str, Any]:
    ids = _candidate_ids(csv_path)
    usage = load_usage_rows(workbook_path)
    people = {item["person"] for item in usage.get(("timesheet_review_po_usage_reconcile", "__people__"), [])}
    with FactStore() as store:
        store.initialize()
        rows = [dict(row) for row in store.connection.execute(
            """SELECT c.candidate_id, c.raw_name, c.canonical_name, c.year, c.month,
                      c.total_hours, c.po_id, c.review_status, s.file_name AS source_file
               FROM timesheet_candidates c
               LEFT JOIN source_files s ON s.source_id=c.source_id
               WHERE c.project_id=? AND c.review_status NOT IN ('SUPERSEDED','APPROVED','DISMISSED')
                 AND c.candidate_id IN (%s)""" % ",".join("?" for _ in ids),
            (PROJECT, *ids),
        )]

    groups: dict[tuple[str, int, int], list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        person = row.get("canonical_name") or row.get("raw_name") or ""
        groups[(person, int(row["year"]), int(row["month"]))].append(row)

    plan: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    ambiguous_groups: list[dict[str, Any]] = []
    for (person, year, month), group in sorted(groups.items()):
        usage_rows = _usage_for(usage, person, year, month)
        if not usage_rows:
            continue
        matches: dict[str, list[dict[str, Any]]] = {}
        ambiguous = False
        for item in usage_rows:
            matching = [
                row for row in group
                if _digits(row.get("po_id")) == _digits(item["po_id"])
                and _float_equal(row.get("total_hours"), item["hours"])
            ]
            matches[item["po_id"]] = matching
            if len(matching) != 1:
                ambiguous = True
        if ambiguous:
            ambiguous_groups.append({
                "person": person,
                "period": f"{year:04d}-{month:02d}",
                "candidate_ids": [row["candidate_id"] for row in group],
                "candidates": [
                    {
                        "candidate_id": row["candidate_id"],
                        "po_id": row.get("po_id") or "",
                        "hours": row.get("total_hours") or "",
                        "source_file": row.get("source_file") or "",
                    }
                    for row in group
                ],
                "candidate_files": [row.get("source_file") or "" for row in group],
                "workbook": [{"po_id": item["po_id"], "hours": item["hours"]} for item in usage_rows],
            })
            continue
        keep = {matching[0]["candidate_id"] for matching in matches.values()}
        extras = [row for row in group if row["candidate_id"] not in keep]
        if not extras:
            continue
        replacements = {item["candidate_id"]: item for item in group if item["candidate_id"] in keep}
        for extra in extras:
            replacement_id = next(iter(replacements), "")
            plan.append({
                "candidate_id": extra["candidate_id"],
                "replacement_id": replacement_id,
                "period": f"{year:04d}-{month:02d}",
                "person": person,
                "source_file": extra.get("source_file") or "",
                "replacement_file": replacements[replacement_id].get("source_file") or "",
                "note": (
                    f"{DATA_SOURCE} version reconciliation: retained candidate {replacement_id} "
                    f"because its PO and hours match the workbook; superseded duplicate/version candidate."
                ),
            })
        skipped.append({
            "person": person,
            "period": f"{year:04d}-{month:02d}",
            "kept": sorted(keep),
            "superseded": [row["candidate_id"] for row in extras],
        })
    return {
        "candidate_rows": len(rows),
        "groups": len(skipped),
        "ambiguous_groups": ambiguous_groups,
        "plan": plan,
    }


def apply_plan(plan: dict[str, Any]) -> int:
    applied = 0
    with FactStore() as store:
        store.initialize()
        for item in plan["plan"]:
            store.supersede_candidate(item["candidate_id"], item["replacement_id"], item["note"])
            applied += 1
        store.connection.commit()
    return applied


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Keep workbook-consistent historical Timesheet versions")
    parser.add_argument("--workbook", required=True)
    parser.add_argument("--candidate-csv", required=True)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args(argv)
    plan = build_plan(args.workbook, args.candidate_csv)
    applied = apply_plan(plan) if args.apply else 0
    result = {
        "ok": True,
        "project": PROJECT,
        "apply": args.apply,
        "candidate_rows": plan["candidate_rows"],
        "duplicate_groups": plan["groups"],
        "ambiguous_groups": len(plan["ambiguous_groups"]),
        "supersede_plan": len(plan["plan"]),
        "applied": applied,
        "fact_approval_performed": False,
        "items": plan["plan"],
        "ambiguous": plan["ambiguous_groups"],
    }
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
