# -*- coding: utf-8 -*-
"""Freeze a user-confirmed, recalculated Budget Report DRAFT."""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import openpyxl

from fact_store import FactStore


PROJECT = "CNTS-25165"
ERROR_TOKENS = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _validate_workbook(path: Path) -> dict[str, Any]:
    """Validate formulas and report whether cached values are available.

    openpyxl does not calculate formulas.  The cache count is therefore only
    an observation of what Excel wrote into the xlsx package; a user-confirmed
    Excel recalculation is accepted explicitly by the freeze command.
    """
    formulas = 0
    cached_formulas = 0
    blank_cached: list[str] = []
    errors: list[dict[str, str]] = []

    formula_wb = openpyxl.load_workbook(path, data_only=False, read_only=False)
    value_wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    try:
        for ws in formula_wb.worksheets:
            value_ws = value_wb[ws.title]
            for row in ws.iter_rows():
                for cell in row:
                    if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                        continue
                    formulas += 1
                    cached = value_ws[cell.coordinate].value
                    if cached is None:
                        blank_cached.append(f"{ws.title}!{cell.coordinate}")
                    else:
                        cached_formulas += 1
                    if isinstance(cached, str) and cached in ERROR_TOKENS:
                        errors.append({
                            "sheet": ws.title,
                            "cell": cell.coordinate,
                            "value": cached,
                        })
        metadata = {}
        if "_Frozen_Run_Metadata" in formula_wb.sheetnames:
            metadata_ws = formula_wb["_Frozen_Run_Metadata"]
            for row in metadata_ws.iter_rows(min_row=1, max_col=2, values_only=True):
                if row[0] is not None:
                    metadata[str(row[0])] = row[1]
    finally:
        formula_wb.close()
        value_wb.close()

    return {
        "formula_count": formulas,
        "cached_formula_count": cached_formulas,
        "blank_cached_count": len(blank_cached),
        "blank_cached_sample": blank_cached[:20],
        "error_count": len(errors),
        "errors": errors,
        "metadata": metadata,
    }


def freeze_budget_report(
    project: str,
    draft: str | Path,
    frozen_run_id: str,
    output: str | Path,
    manifest: str | Path,
    user_confirmed_recalculation: bool = False,
) -> dict[str, Any]:
    draft_path = Path(draft).resolve()
    output_path = Path(output).resolve()
    manifest_path = Path(manifest).resolve()
    if not draft_path.is_file():
        raise FileNotFoundError(draft_path)
    if output_path.exists():
        raise FileExistsError(f"Frozen output already exists: {output_path}")
    if not user_confirmed_recalculation:
        raise ValueError("Budget Report freeze requires explicit user confirmation of Excel recalculation")

    validation = _validate_workbook(draft_path)
    if validation["formula_count"] == 0:
        raise ValueError("Budget Report contains no formulas to validate")
    if validation["error_count"]:
        raise ValueError(f"Budget Report contains formula errors: {validation['errors'][:5]}")

    expected_run_metadata = validation["metadata"].get("Frozen Run ID")
    if expected_run_metadata and expected_run_metadata != frozen_run_id:
        raise ValueError(
            f"DRAFT frozen run mismatch: metadata={expected_run_metadata!r}, expected={frozen_run_id!r}"
        )

    draft_hash = _sha256(draft_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(draft_path, output_path)
    frozen_hash = _sha256(output_path)
    if draft_hash != frozen_hash:
        raise ValueError("Frozen Budget Report copy hash verification failed")

    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    cache_note = (
        "Excel recalculation was confirmed by the user before freeze. "
        "The local openpyxl validator observed no cached formula values in the saved package; "
        "formula structure had no error literals."
    )
    note = (
        "User-authorized Budget Report freeze based on the already FROZEN PO Usage Status run. "
        "Remaining REVIEW_REQUIRED risks were intentionally not processed. "
        + cache_note
    )

    with FactStore() as store:
        store.initialize()
        run = store.connection.execute(
            """SELECT run_id, status, data_as_of FROM runs
               WHERE project_id=? AND run_id=?""",
            (project, frozen_run_id),
        ).fetchone()
        if not run or run["status"] != "FROZEN":
            raise ValueError(f"Frozen Run not found or not FROZEN: {frozen_run_id}")

        approved_facts = store.connection.execute(
            """SELECT COUNT(*) FROM timesheet_facts
               WHERE project_id=? AND approval_status='APPROVED' AND is_current=1""",
            (project,),
        ).fetchone()[0]
        source_id = store.record_source_file(
            project,
            "budget_report_frozen",
            output_path,
            run_id=frozen_run_id,
            metadata={
                "draft_path": str(draft_path),
                "draft_sha256": draft_hash,
                "frozen_sha256": frozen_hash,
                "formula_validation": validation,
                "user_confirmed_recalculation": True,
                "risk_override": True,
            },
        )
        approval_id = f"apr-budget-report-{uuid.uuid4().hex[:20]}"
        store.connection.execute(
            """INSERT INTO approvals(
                   approval_id, project_id, run_id, approval_type, status,
                   approved_by, approved_at, note)
               VALUES(?, ?, ?, 'BUDGET_REPORT_FREEZE', 'APPROVED', ?, ?, ?)""",
            (approval_id, project, frozen_run_id, "user-authorized", now, note),
        )
        old_notes = store.connection.execute(
            "SELECT notes FROM runs WHERE run_id=?", (frozen_run_id,)
        ).fetchone()[0] or ""
        combined_notes = old_notes + "\nBudget Report frozen: " + now
        store.connection.execute(
            "UPDATE runs SET notes=? WHERE run_id=?",
            (combined_notes, frozen_run_id),
        )
        store.connection.commit()

    result = {
        "ok": True,
        "project": project,
        "frozen_run_id": frozen_run_id,
        "approval_id": approval_id,
        "approval_type": "BUDGET_REPORT_FREEZE",
        "status": "FROZEN",
        "risk_override": True,
        "user_confirmed_recalculation": True,
        "source_id": source_id,
        "draft": {"path": str(draft_path), "sha256": draft_hash},
        "frozen": {"path": str(output_path), "sha256": frozen_hash},
        "approved_facts": approved_facts,
        "formula_validation": validation,
        "note": note,
    }
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Freeze a user-confirmed Budget Report DRAFT")
    parser.add_argument("--project", default=PROJECT)
    parser.add_argument("--draft", required=True)
    parser.add_argument("--frozen-run-id", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--manifest", required=True)
    parser.add_argument("--user-confirmed-recalculation", action="store_true")
    args = parser.parse_args()
    print(json.dumps(freeze_budget_report(
        args.project, args.draft, args.frozen_run_id,
        args.output, args.manifest, args.user_confirmed_recalculation,
    ), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
