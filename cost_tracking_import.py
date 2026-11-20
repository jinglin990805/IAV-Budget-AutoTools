# -*- coding: utf-8 -*-
"""Import an audited Cost Tracking workbook as PENDING cost facts."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from cost_tracking_audit import audit
from fact_store import FactStore


PROJECT = "CNTS-25165"


def _fact_id(record: dict, source_sha256: str) -> str:
    key = "|".join(str(record.get(field, "")) for field in (
        "wbs", "month", "cost_layer", "employee_id", "employee_name", "amount", "hours",
    ))
    return "cost-" + hashlib.sha256(f"{source_sha256}|{key}".encode("utf-8")).hexdigest()[:24]


def apply_import(source: str | Path, run_id: str) -> dict:
    source_path = Path(source).resolve()
    dry_run = audit(source_path)
    if dry_run["arithmetic_check_count"] or dry_run["duplicate_count"]:
        raise ValueError("Cost Tracking dry-run has arithmetic or duplicate failures")

    records = [record for sheet in dry_run["sheets"] for record in sheet["records"]]
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    inserted = 0
    skipped = 0
    source_id = None
    with FactStore() as store:
        store.initialize()
        run = store.connection.execute(
            "SELECT run_id, status FROM runs WHERE project_id=? AND run_id=?",
            (PROJECT, run_id),
        ).fetchone()
        if not run or run["status"] != "FROZEN":
            raise ValueError(f"Frozen Run not found or not FROZEN: {run_id}")

        source_id = store.record_source_file(
            PROJECT,
            "cost_tracking_workbook",
            source_path,
            run_id=run_id,
            metadata={
                "source_sha256": dry_run["source"]["sha256"],
                "data_as_of": "2026-08",
                "dry_run_record_count": dry_run["record_count"],
                "transfer_rule": "Negative cost transfer values are retained as source values; Xu Chang 2026-08 -244h/-91826.96 is a WBS transfer paired with a positive receiving-side movement.",
                "approval_status": "PENDING",
            },
        )
        for record in records:
            fact_id = _fact_id(record, dry_run["source"]["sha256"])
            exists = store.connection.execute(
                "SELECT 1 FROM cost_facts WHERE fact_id=?", (fact_id,)
            ).fetchone()
            if exists:
                skipped += 1
                continue
            note = {
                "wbs": record["wbs"],
                "sheet": record["sheet"],
                "employee_id": record["employee_id"],
                "employee_name": record["employee_name"],
                "hours": record["hours"],
                "source_file": str(source_path),
                "transfer_values_retained": True,
            }
            store.connection.execute(
                """INSERT INTO cost_facts(
                       fact_id, project_id, run_id, month, source_type, po_id,
                       cost_layer, currency, original_amount, amount,
                       booking_status, source_id, approval_status, notes)
                   VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    fact_id, PROJECT, run_id, record["month"], "cost_tracking",
                    None, record["cost_layer"], record["currency"], record["amount"],
                    record["amount"], "PENDING", source_id, "PENDING",
                    json.dumps(note, ensure_ascii=False, sort_keys=True),
                ),
            )
            inserted += 1
        store.connection.commit()

    return {
        "ok": True,
        "project": PROJECT,
        "run_id": run_id,
        "source_id": source_id,
        "source": dry_run["source"],
        "dry_run_record_count": dry_run["record_count"],
        "inserted": inserted,
        "skipped_existing": skipped,
        "status": "APPLIED_PENDING",
        "transfer_rule": "Negative cost transfer values retained as source values.",
        "approval_status": "PENDING",
        "booking_status": "PENDING",
        "write_to_original": False,
    }


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Import audited Cost Tracking as PENDING facts")
    parser.add_argument("--source", required=True)
    parser.add_argument("--run-id", required=True)
    args = parser.parse_args()
    print(json.dumps(apply_import(args.source, args.run_id), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
