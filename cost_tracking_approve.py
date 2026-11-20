# -*- coding: utf-8 -*-
"""Approve audited Cost Tracking facts without marking them as booked."""

from __future__ import annotations

import argparse
import json
import sys
import uuid
from datetime import datetime, timezone

from fact_store import FactStore


PROJECT = "CNTS-25165"


def approve_cost_facts(source_id: str, run_id: str) -> dict:
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    with FactStore() as store:
        store.initialize()
        run = store.connection.execute(
            "SELECT run_id, status FROM runs WHERE project_id=? AND run_id=?",
            (PROJECT, run_id),
        ).fetchone()
        if not run or run["status"] != "FROZEN":
            raise ValueError(f"Frozen Run not found or not FROZEN: {run_id}")

        source = store.connection.execute(
            """SELECT source_id, sha256, path FROM source_files
               WHERE project_id=? AND source_id=? AND source_type='cost_tracking_workbook'""",
            (PROJECT, source_id),
        ).fetchone()
        if not source:
            raise ValueError(f"Cost Tracking source not found: {source_id}")

        rows = store.connection.execute(
            """SELECT fact_id, approval_status, booking_status
               FROM cost_facts
               WHERE project_id=? AND run_id=? AND source_id=?
               ORDER BY fact_id""",
            (PROJECT, run_id, source_id),
        ).fetchall()
        if not rows:
            raise ValueError("No cost facts found for the selected source and run")
        pending = [row for row in rows if row["approval_status"] == "PENDING"]
        unexpected = [
            row for row in rows
            if row["approval_status"] not in {"PENDING", "APPROVED"}
            or row["booking_status"] not in {"PENDING", "POSTED"}
        ]
        if unexpected:
            raise ValueError(f"Unexpected cost fact statuses: {len(unexpected)}")

        store.connection.execute(
            """UPDATE cost_facts
               SET approval_status='APPROVED'
               WHERE project_id=? AND run_id=? AND source_id=?
                 AND approval_status='PENDING'""",
            (PROJECT, run_id, source_id),
        )
        approval_id = f"apr-cost-facts-{uuid.uuid4().hex[:20]}"
        note = (
            f"User-authorized approval of {len(rows)} Cost Tracking facts from source {source_id}. "
            "GR, invoice and payment are tracked separately; an invoice may set booking_status to POSTED. "
            "Xu Chang 2026-08 -244h/-91826.96 retained as confirmed WBS cost transfer."
        )
        store.connection.execute(
            """INSERT INTO approvals(
                   approval_id, project_id, run_id, approval_type, status,
                   approved_by, approved_at, note)
               VALUES(?, ?, ?, 'COST_FACT_APPROVAL', 'APPROVED', ?, ?, ?)""",
            (approval_id, PROJECT, run_id, "user-authorized", now, note),
        )
        store.connection.commit()

        final = store.connection.execute(
            """SELECT approval_status, booking_status, COUNT(*) AS count
               FROM cost_facts
               WHERE project_id=? AND run_id=? AND source_id=?
               GROUP BY approval_status, booking_status""",
            (PROJECT, run_id, source_id),
        ).fetchall()

    return {
        "ok": True,
        "project": PROJECT,
        "run_id": run_id,
        "source_id": source_id,
        "source_sha256": source["sha256"],
        "source_path": source["path"],
        "approval_id": approval_id,
        "approval_type": "COST_FACT_APPROVAL",
        "total_facts": len(rows),
        "approved_now": len(pending),
        "status_summary": [dict(row) for row in final],
        "approval_status": "APPROVED",
        "booking_rule": "Invoice recorded => booking_status POSTED; payment remains separately tracked",
        "note": note,
    }


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Approve audited Cost Tracking facts")
    parser.add_argument("--source-id", required=True)
    parser.add_argument("--run-id", required=True)
    args = parser.parse_args()
    print(json.dumps(approve_cost_facts(args.source_id, args.run_id), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
