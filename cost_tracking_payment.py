# -*- coding: utf-8 -*-
"""Preview or record payment status for invoiced Cost Tracking records.

Payment is deliberately separate from invoice booking.  The command defaults
to a read-only preview; ``--apply`` is required before any payment fields are
written.  ``payment_ref`` is an optional business payment reference and is not
used as a financial voucher number.
"""

from __future__ import annotations

import argparse
import json
import sys
import uuid
from datetime import datetime, timezone

from fact_store import FactStore


PROJECT = "CNTS-25165"


def _validate_payment_date(value: str) -> str:
    try:
        datetime.fromisoformat(value)
    except ValueError as exc:
        raise ValueError("payment_date must be ISO format, for example 2026-09-22") from exc
    return value


def record_payment(
    run_id: str,
    payment_date: str,
    payment_ref: str | None = None,
    invoice_ref: str | None = None,
    fact_ids: list[str] | None = None,
    allocation_ids: list[str] | None = None,
    apply: bool = False,
) -> dict:
    """Preview or apply payment status to already invoiced targets."""

    payment_date = _validate_payment_date(payment_date)
    fact_ids = fact_ids or []
    allocation_ids = allocation_ids or []
    if not invoice_ref and not fact_ids and not allocation_ids:
        raise ValueError("Provide --invoice-ref, --fact-id, or --allocation-id")

    with FactStore() as store:
        store.initialize()
        run = store.connection.execute(
            "SELECT status FROM runs WHERE project_id=? AND run_id=?",
            (PROJECT, run_id),
        ).fetchone()
        if not run or run["status"] != "FROZEN":
            raise ValueError(f"Frozen Run not found or not FROZEN: {run_id}")

        facts = []
        if fact_ids:
            placeholders = ",".join("?" for _ in fact_ids)
            facts = store.connection.execute(
                f"""SELECT fact_id, invoice_ref, booking_status, payment_status
                    FROM cost_facts
                    WHERE project_id=? AND run_id=? AND fact_id IN ({placeholders})""",
                [PROJECT, run_id, *fact_ids],
            ).fetchall()
        elif invoice_ref:
            facts = store.connection.execute(
                """SELECT fact_id, invoice_ref, booking_status, payment_status
                   FROM cost_facts
                   WHERE project_id=? AND run_id=? AND invoice_ref=?""",
                (PROJECT, run_id, invoice_ref),
            ).fetchall()

        allocations = []
        if allocation_ids:
            placeholders = ",".join("?" for _ in allocation_ids)
            allocations = store.connection.execute(
                f"""SELECT allocation_id, invoice_ref, booking_status, payment_status
                    FROM cost_fact_po_allocations
                    WHERE project_id=? AND run_id=? AND allocation_id IN ({placeholders})""",
                [PROJECT, run_id, *allocation_ids],
            ).fetchall()
        elif invoice_ref:
            allocations = store.connection.execute(
                """SELECT allocation_id, invoice_ref, booking_status, payment_status
                   FROM cost_fact_po_allocations
                   WHERE project_id=? AND run_id=? AND invoice_ref=?""",
                (PROJECT, run_id, invoice_ref),
            ).fetchall()

        if fact_ids and len(facts) != len(fact_ids):
            raise ValueError("One or more requested cost fact IDs were not found")
        if allocation_ids and len(allocations) != len(allocation_ids):
            raise ValueError("One or more requested allocation IDs were not found")
        if not facts and not allocations:
            raise ValueError("No invoiced payment targets found")

        invalid_facts = [
            row["fact_id"] for row in facts
            if not row["invoice_ref"]
            or row["booking_status"] != "POSTED"
            or row["payment_status"] != "PENDING"
        ]
        invalid_allocations = [
            row["allocation_id"] for row in allocations
            if not row["invoice_ref"]
            or row["booking_status"] != "POSTED"
            or row["payment_status"] != "PENDING"
        ]
        if invalid_facts or invalid_allocations:
            raise ValueError(
                "Payment requires POSTED, invoiced, PENDING targets; "
                f"invalid facts={invalid_facts}, allocations={invalid_allocations}"
            )

        result = {
            "ok": True,
            "project": PROJECT,
            "run_id": run_id,
            "invoice_ref": invoice_ref,
            "payment_date": payment_date,
            "payment_ref": payment_ref,
            "fact_ids": [row["fact_id"] for row in facts],
            "allocation_ids": [row["allocation_id"] for row in allocations],
            "target_count": len(facts) + len(allocations),
            "status": "APPLIED" if apply else "DRY_RUN",
            "payment_status": "PAID" if apply else "PENDING",
        }
        if not apply:
            return result

        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        if facts:
            fact_placeholders = ",".join("?" for _ in facts)
            store.connection.execute(
                f"""UPDATE cost_facts
                    SET payment_status='PAID', payment_ref=?, payment_date=?
                    WHERE project_id=? AND run_id=?
                      AND fact_id IN ({fact_placeholders})""",
                [payment_ref, payment_date, PROJECT, run_id,
                 *[row["fact_id"] for row in facts]],
            )
        if allocations:
            allocation_placeholders = ",".join("?" for _ in allocations)
            store.connection.execute(
                f"""UPDATE cost_fact_po_allocations
                    SET payment_status='PAID', payment_ref=?, payment_date=?
                    WHERE project_id=? AND run_id=?
                      AND allocation_id IN ({allocation_placeholders})""",
                [payment_ref, payment_date, PROJECT, run_id,
                 *[row["allocation_id"] for row in allocations]],
            )
        approval_id = f"apr-cost-payment-{uuid.uuid4().hex[:20]}"
        store.connection.execute(
            """INSERT INTO approvals(
                 approval_id, project_id, run_id, approval_type, status,
                 approved_by, approved_at, note)
               VALUES(?, ?, ?, 'COST_PAYMENT_RECORDING', 'APPROVED', ?, ?, ?)""",
            (
                approval_id, PROJECT, run_id, "user-authorized", now,
                "Payment status recorded for explicitly selected invoiced targets; "
                "financial voucher numbers are outside this system scope.",
            ),
        )
        store.connection.commit()
        result["approval_id"] = approval_id
        return result


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Preview or record Cost Tracking payment status")
    parser.add_argument("--run-id", required=True)
    parser.add_argument("--payment-date", required=True)
    parser.add_argument("--payment-ref")
    parser.add_argument("--invoice-ref")
    parser.add_argument("--fact-id", dest="fact_ids", action="append")
    parser.add_argument("--allocation-id", dest="allocation_ids", action="append")
    parser.add_argument("--apply", action="store_true", help="Write PAID status; default is read-only preview")
    args = parser.parse_args()
    print(json.dumps(record_payment(
        run_id=args.run_id,
        payment_date=args.payment_date,
        payment_ref=args.payment_ref,
        invoice_ref=args.invoice_ref,
        fact_ids=args.fact_ids,
        allocation_ids=args.allocation_ids,
        apply=args.apply,
    ), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
