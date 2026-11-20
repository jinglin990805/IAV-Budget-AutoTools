# -*- coding: utf-8 -*-
"""Close stale historical review reminders after Excel reconciliation.

Only candidates whose person/month/PO/hours exactly match the authoritative
PO Usage Status workbook are eligible.  This closes stale OCR, validation and
version-conflict queue items, but deliberately leaves NEW_MEMBER items open.
It never changes candidate facts and never creates an approved Timesheet fact.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from pathlib import Path
from typing import Any

from fact_store import FactStore
from timesheet_review_po_usage_reconcile import (
    DATA_SOURCE,
    PROJECT,
    _digits,
    _float_equal,
    _usage_for,
    load_usage_rows,
)


RESOLVABLE_QUEUE_TYPES = {"OCR_REVIEW", "VALIDATION_REVIEW", "VERSION_CONFLICT"}


def _candidate_ids(path: str | Path) -> set[str]:
    with Path(path).open("r", encoding="utf-8-sig", newline="") as handle:
        return {
            str(row.get("candidate_id") or "").strip()
            for row in csv.DictReader(handle)
            if str(row.get("candidate_id") or "").strip()
        }


def _load_rows(ids: set[str]) -> list[dict[str, Any]]:
    if not ids:
        return []
    placeholders = ",".join("?" for _ in ids)
    with FactStore() as store:
        store.initialize()
        rows = store.connection.execute(
            f"""SELECT c.candidate_id, c.raw_name, c.canonical_name, c.year, c.month,
                       c.po_id, c.total_hours, c.review_status
                FROM timesheet_candidates c
                WHERE c.project_id=? AND c.candidate_id IN ({placeholders})
                  AND c.review_status NOT IN ('SUPERSEDED','APPROVED','DISMISSED')""",
            (PROJECT, *sorted(ids)),
        ).fetchall()
    return [dict(row) for row in rows]


def build_plan(workbook: str | Path, candidate_csv: str | Path) -> dict[str, Any]:
    usage = load_usage_rows(workbook)
    rows = _load_rows(_candidate_ids(candidate_csv))
    plan: list[dict[str, Any]] = []
    deferred: list[dict[str, Any]] = []
    with FactStore() as store:
        store.initialize()
        for row in rows:
            person = row.get("canonical_name") or row.get("raw_name") or ""
            entries = _usage_for(usage, person, int(row["year"]), int(row["month"]))
            exact = any(
                _digits(row.get("po_id")) == _digits(entry.get("po_id"))
                and _float_equal(row.get("total_hours"), entry.get("hours"))
                for entry in entries
            )
            open_items = [
                dict(item)
                for item in store.connection.execute(
                    """SELECT review_id, queue_type FROM review_queue
                       WHERE candidate_id=? AND status='OPEN'""",
                    (row["candidate_id"],),
                )
            ]
            if not entries or not exact:
                if open_items:
                    deferred.append({
                        "candidate_id": row["candidate_id"],
                        "person": person,
                        "period": f"{int(row['year']):04d}-{int(row['month']):02d}",
                        "reason": "no exact workbook person/month/PO/hours match",
                        "open_queue_types": sorted({item["queue_type"] for item in open_items}),
                    })
                continue
            for item in open_items:
                if item["queue_type"] not in RESOLVABLE_QUEUE_TYPES:
                    deferred.append({
                        "candidate_id": row["candidate_id"],
                        "person": person,
                        "period": f"{int(row['year']):04d}-{int(row['month']):02d}",
                        "reason": f"queue type retained: {item['queue_type']}",
                        "open_queue_types": [item["queue_type"]],
                    })
                    continue
                plan.append({
                    "review_id": item["review_id"],
                    "candidate_id": row["candidate_id"],
                    "queue_type": item["queue_type"],
                    "person": person,
                    "period": f"{int(row['year']):04d}-{int(row['month']):02d}",
                    "po_id": row.get("po_id") or "",
                    "hours": row.get("total_hours") or "",
                })
    return {
        "candidate_rows": len(rows),
        "resolve_plan": plan,
        "deferred": deferred,
    }


def apply_plan(plan: dict[str, Any]) -> int:
    if not plan["resolve_plan"]:
        return 0
    with FactStore() as store:
        store.initialize()
        for item in plan["resolve_plan"]:
            note = (
                f"{DATA_SOURCE} final historical reconciliation: "
                f"{item['person']} {item['period']} PO {item['po_id']} "
                f"{item['hours']}h matches workbook; stale {item['queue_type']} "
                "reminder resolved. No Timesheet fact approval performed."
            )
            store.connection.execute(
                """UPDATE review_queue
                   SET status='RESOLVED', resolved_at=datetime('now'), resolution_note=?
                   WHERE review_id=? AND status='OPEN'""",
                (note, item["review_id"]),
            )
        store.connection.commit()
    return len(plan["resolve_plan"])


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Resolve stale historical review reminders after workbook reconciliation")
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
        "resolve_plan": len(plan["resolve_plan"]),
        "deferred_count": len(plan["deferred"]),
        "applied": applied,
        "fact_approval_performed": False,
        "deferred": plan["deferred"],
        "items": plan["resolve_plan"],
    }
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
