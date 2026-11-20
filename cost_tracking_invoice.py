# -*- coding: utf-8 -*-
"""Record invoice evidence and mark matching approved cost facts as POSTED."""

from __future__ import annotations

import argparse
import json
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path

from fact_store import FactStore


PROJECT = "CNTS-25165"


def record_invoice(
    source_id: str,
    run_id: str,
    invoice_ref: str,
    invoice_date: str | None = None,
    invoice_file: str | Path | None = None,
    gr_ref: str | None = None,
    fact_ids: list[str] | None = None,
    po_id: str | None = None,
) -> dict:
    invoice_path = Path(invoice_file).resolve() if invoice_file else None
    if invoice_path and not invoice_path.is_file():
        raise FileNotFoundError(invoice_path)
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    with FactStore() as store:
        store.initialize()
        run = store.connection.execute(
            "SELECT run_id, status FROM runs WHERE project_id=? AND run_id=?",
            (PROJECT, run_id),
        ).fetchone()
        if not run or run["status"] != "FROZEN":
            raise ValueError(f"Frozen Run not found or not FROZEN: {run_id}")
        fact_query = (
            "SELECT fact_id, approval_status, booking_status "
            "FROM cost_facts WHERE project_id=? AND run_id=? AND source_id=?"
        )
        fact_params: list[object] = [PROJECT, run_id, source_id]
        if fact_ids:
            placeholders = ",".join("?" for _ in fact_ids)
            fact_query += f" AND fact_id IN ({placeholders})"
            fact_params.extend(fact_ids)
        facts = store.connection.execute(fact_query, fact_params).fetchall()
        if not facts:
            raise ValueError("No cost facts found for the selected source and run")
        not_approved = [row for row in facts if row["approval_status"] != "APPROVED"]
        if not_approved:
            raise ValueError("Marking invoice-booked costs requires all matching cost facts to be APPROVED")

        invoice_source_id = None
        if invoice_path:
            invoice_source_id = store.record_source_file(
                PROJECT,
                "cost_invoice",
                invoice_path,
                run_id=run_id,
                metadata={
                    "invoice_ref": invoice_ref,
                    "invoice_date": invoice_date,
                    "gr_ref": gr_ref,
                },
            )
        update_query = (
            "UPDATE cost_facts SET po_id=COALESCE(?, po_id), gr_ref=?, invoice_ref=?, invoice_date=?, "
            "invoice_source_id=?, booking_status='POSTED' "
            "WHERE project_id=? AND run_id=? AND source_id=? "
            "AND approval_status='APPROVED'"
        )
        update_params: list[object] = [
            po_id, gr_ref, invoice_ref, invoice_date, invoice_source_id,
            PROJECT, run_id, source_id,
        ]
        if fact_ids:
            placeholders = ",".join("?" for _ in fact_ids)
            update_query += f" AND fact_id IN ({placeholders})"
            update_params.extend(fact_ids)
        store.connection.execute(update_query, update_params)
        approval_id = f"apr-cost-invoice-{uuid.uuid4().hex[:20]}"
        note = (
            f"Invoice {invoice_ref} recorded for {len(facts)} approved Cost Tracking facts"
            f"{(' selected by fact_id' if fact_ids else '')}. "
            f"{('PO ' + po_id + ' linked. ' if po_id else '')}"
            "booking_status set to POSTED under the project rule that invoicing is sufficient; "
            "payment remains separately tracked."
        )
        store.connection.execute(
            """INSERT INTO approvals(
                   approval_id, project_id, run_id, approval_type, status,
                   approved_by, approved_at, note)
               VALUES(?, ?, ?, 'COST_INVOICE_RECORDING', 'APPROVED', ?, ?, ?)""",
            (approval_id, PROJECT, run_id, "user-authorized", now, note),
        )
        store.connection.commit()
        status_query = (
            "SELECT approval_status, booking_status, payment_status, COUNT(*) AS count "
            "FROM cost_facts WHERE project_id=? AND run_id=? AND source_id=?"
        )
        status_params: list[object] = [PROJECT, run_id, source_id]
        if fact_ids:
            placeholders = ",".join("?" for _ in fact_ids)
            status_query += f" AND fact_id IN ({placeholders})"
            status_params.extend(fact_ids)
        status_query += " GROUP BY approval_status, booking_status, payment_status"
        status_summary = store.connection.execute(status_query, status_params).fetchall()

    return {
        "ok": True,
        "project": PROJECT,
        "run_id": run_id,
        "source_id": source_id,
        "invoice_ref": invoice_ref,
        "invoice_date": invoice_date,
        "po_id": po_id,
        "gr_ref": gr_ref,
        "invoice_source_id": invoice_source_id,
        "approval_id": approval_id,
        "updated_facts": len(facts),
        "status_summary": [dict(row) for row in status_summary],
        "approval_type": "COST_INVOICE_RECORDING",
        "booking_status": "POSTED",
        "payment_status": "PENDING",
        "fact_ids": fact_ids or [],
        "note": note,
    }


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Record invoice and mark approved costs POSTED")
    parser.add_argument("--source-id", required=True)
    parser.add_argument("--run-id", required=True)
    parser.add_argument("--invoice-ref", required=True)
    parser.add_argument("--invoice-date")
    parser.add_argument("--invoice-file")
    parser.add_argument("--gr-ref")
    parser.add_argument("--po-id", help="Business PO linked to the selected cost facts")
    parser.add_argument(
        "--fact-id", dest="fact_ids", action="append",
        help="Restrict the invoice to selected cost fact IDs; repeat the option for multiple IDs",
    )
    args = parser.parse_args()
    print(json.dumps(record_invoice(
        args.source_id, args.run_id, args.invoice_ref,
        args.invoice_date, args.invoice_file, args.gr_ref, args.fact_ids,
        args.po_id,
    ), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
