# -*- coding: utf-8 -*-
"""Approve only historical Timesheet candidates matching the PO workbook.

This is a separate, explicit historical approval command.  It is intentionally
limited to candidate IDs in the exported historical review CSV and requires an
exact workbook match for person/month/PO/hours.  Non-matching candidates are
left untouched.  The command is dry-run by default.
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


def _candidate_ids(path: str | Path) -> set[str]:
    with Path(path).open("r", encoding="utf-8-sig", newline="") as handle:
        return {
            str(row.get("candidate_id") or "").strip()
            for row in csv.DictReader(handle)
            if str(row.get("candidate_id") or "").strip()
        }


def _load_candidates(ids: set[str]) -> list[dict[str, Any]]:
    if not ids:
        return []
    placeholders = ",".join("?" for _ in ids)
    with FactStore() as store:
        store.initialize()
        rows = store.connection.execute(
            f"""SELECT c.*
                FROM timesheet_candidates c
                WHERE c.project_id=? AND c.candidate_id IN ({placeholders})
                  AND c.review_status NOT IN ('APPROVED','SUPERSEDED','DISMISSED')
                ORDER BY c.year,c.month,c.candidate_id""",
            (PROJECT, *sorted(ids)),
        ).fetchall()
    return [dict(row) for row in rows]


def build_plan(workbook: str | Path, candidate_csv: str | Path) -> dict[str, Any]:
    usage = load_usage_rows(workbook)
    rows = _load_candidates(_candidate_ids(candidate_csv))
    approve: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    for row in rows:
        person = row.get("canonical_name") or row.get("raw_name") or ""
        year = int(row["year"])
        month_number = int(row["month"])
        entries = _usage_for(usage, person, year, month_number)
        matches = [
            entry for entry in entries
            if _digits(row.get("po_id")) == _digits(entry.get("po_id"))
            and _float_equal(row.get("total_hours"), entry.get("hours"))
        ]
        if len(matches) != 1:
            skipped.append({
                "candidate_id": row["candidate_id"],
                "person": person,
                "period": f"{year:04d}-{month_number:02d}",
                "candidate_po": row.get("po_id") or "",
                "candidate_hours": row.get("total_hours"),
                "reason": "no unique exact workbook person/month/PO/hours match",
                "workbook_rows": [
                    {"po_id": entry.get("po_id"), "hours": entry.get("hours")}
                    for entry in entries
                ],
            })
            continue
        entry = matches[0]
        approve.append({
            "candidate_id": row["candidate_id"],
            "fact_id": f"tsf-{row['candidate_id']}",
            "run_id": row["run_id"],
            "source_id": row.get("source_id"),
            "person": person,
            "period": f"{year:04d}-{month_number:02d}",
            "po_id": str(entry["po_id"]),
            "hours": float(entry["hours"]),
        })
    return {
        "candidate_rows": len(rows),
        "approve": approve,
        "skipped": skipped,
    }


def _resolve_open_validation_reviews_for_approved(store: FactStore) -> int:
    rows = store.connection.execute(
        """SELECT q.review_id
           FROM review_queue q
           JOIN timesheet_candidates c ON c.candidate_id=q.candidate_id
           WHERE c.project_id=? AND c.review_status='APPROVED'
             AND q.queue_type='VALIDATION_REVIEW' AND q.status='OPEN'""",
        (PROJECT,),
    ).fetchall()
    if not rows:
        return 0
    note = (
        f"{DATA_SOURCE}: candidate is already approved after exact workbook match; "
        "stale validation reminder resolved."
    )
    for row in rows:
        store.connection.execute(
            """UPDATE review_queue
               SET status='RESOLVED', resolved_at=datetime('now'), resolution_note=?
               WHERE review_id=? AND status='OPEN'""",
            (note, row["review_id"]),
        )
    return len(rows)


def apply_plan(plan: dict[str, Any]) -> tuple[int, int]:
    applied = 0
    with FactStore() as store:
        store.initialize()
        for item in plan["approve"]:
            note = (
                f"User-authorized historical approval: {DATA_SOURCE}; "
                f"{item['person']} {item['period']} PO {item['po_id']} "
                f"{item['hours']:g}h exactly matches workbook."
            )
            store.save_approved_timesheet(
                item["candidate_id"], PROJECT, item["run_id"], item["person"],
                item["period"], item["po_id"], item["hours"], item["source_id"],
                note=note,
            )
            store.connection.execute(
                """UPDATE timesheet_candidates
                   SET canonical_name=?, po_id=?, total_hours=?,
                       review_status='APPROVED', is_preferred=1, notes=?
                   WHERE candidate_id=? AND project_id=?""",
                (
                    item["person"], item["po_id"], item["hours"], note,
                    item["candidate_id"], PROJECT,
                ),
            )
            store.connection.execute(
                """UPDATE review_queue
                   SET status='RESOLVED', resolved_at=datetime('now'), resolution_note=?
                   WHERE candidate_id=? AND status='OPEN'""",
                (note, item["candidate_id"]),
            )
            applied += 1
        stale_validation_resolved = _resolve_open_validation_reviews_for_approved(store)
        store.connection.commit()
    return applied, stale_validation_resolved


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Approve exact workbook-matching historical Timesheet facts")
    parser.add_argument("--workbook", required=True)
    parser.add_argument("--candidate-csv", required=True)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args(argv)
    plan = build_plan(args.workbook, args.candidate_csv)
    applied, stale_validation_resolved = apply_plan(plan) if args.apply else (0, 0)
    result = {
        "ok": True,
        "project": PROJECT,
        "apply": args.apply,
        "candidate_rows": plan["candidate_rows"],
        "approve_plan": len(plan["approve"]),
        "skipped_count": len(plan["skipped"]),
        "applied": applied,
        "stale_approved_validation_resolved": stale_validation_resolved,
        "fact_approval_performed": bool(args.apply and applied),
        "skipped": plan["skipped"],
        "approved": plan["approve"],
    }
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
