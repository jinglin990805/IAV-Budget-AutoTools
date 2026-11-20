# -*- coding: utf-8 -*-
"""Freeze an explicitly authorized PO Usage Status DRAFT as a snapshot."""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from fact_store import FactStore
from po_usage_fact_audit import _read_json


PROJECT = "CNTS-25165"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def freeze_snapshot(
    project: str,
    draft: str | Path,
    authority: str | Path,
    step2_json: str | Path,
    output: str | Path,
    manifest: str | Path,
) -> dict[str, Any]:
    draft_path = Path(draft).resolve()
    authority_path = Path(authority).resolve()
    output_path = Path(output).resolve()
    manifest_path = Path(manifest).resolve()
    if not draft_path.is_file():
        raise FileNotFoundError(draft_path)
    if not authority_path.is_file():
        raise FileNotFoundError(authority_path)
    if output_path.exists():
        raise FileExistsError(f"Frozen output already exists: {output_path}")

    draft_hash = _sha256(draft_path)
    authority_hash = _sha256(authority_path)
    if draft_hash != authority_hash:
        raise ValueError("DRAFT and authority workbook hashes differ")
    step2 = _read_json(step2_json)
    closing = step2.get("monthly_closing") or {}
    validation = step2.get("timesheet_validation") or {}
    po_master = step2.get("po_master") or {}

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(draft_path, output_path)
    frozen_hash = _sha256(output_path)
    if frozen_hash != draft_hash:
        raise ValueError("Frozen copy hash verification failed")

    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    note = (
        "User-authorized PO Usage Status freeze while REVIEW_REQUIRED; "
        "remaining Monthly Closing risks and deferred item intentionally ignored. "
        "Snapshot is based on the hash-verified customer DRAFT and authority workbook."
    )
    with FactStore() as store:
        store.initialize()
        run_id = store.create_run(
            project,
            data_as_of=closing.get("active_period") or "2026-08",
            source_cutoff=now,
            notes=note,
        )
        source_id = store.record_source_file(
            project,
            "po_usage_status_frozen",
            output_path,
            run_id=run_id,
            metadata={
                "authority_path": str(authority_path),
                "authority_sha256": authority_hash,
                "draft_path": str(draft_path),
                "draft_sha256": draft_hash,
                "risk_override": True,
            },
        )
        approved_facts = store.connection.execute(
            """SELECT COUNT(*) FROM timesheet_facts
               WHERE project_id=? AND approval_status='APPROVED' AND is_current=1""",
            (project,),
        ).fetchone()[0]
        active_candidates = store.connection.execute(
            """SELECT COUNT(*) FROM timesheet_candidates
               WHERE project_id=? AND review_status<>'SUPERSEDED'""",
            (project,),
        ).fetchone()[0]
        approved_candidates = store.connection.execute(
            """SELECT COUNT(*) FROM timesheet_candidates
               WHERE project_id=? AND review_status='APPROVED'""",
            (project,),
        ).fetchone()[0]
        approval_id = f"apr-po-usage-{uuid.uuid4().hex[:20]}"
        store.connection.execute(
            """INSERT INTO approvals(
                   approval_id, project_id, run_id, approval_type, status,
                   approved_by, approved_at, note)
               VALUES(?, ?, ?, 'PO_USAGE_STATUS_FREEZE', 'APPROVED', ?, ?, ?)""",
            (approval_id, project, run_id, "user-authorized", now, note),
        )
        store.connection.execute(
            """UPDATE runs SET status='FROZEN', completed_at=?, notes=?
               WHERE run_id=?""",
            (now, note, run_id),
        )
        store.connection.commit()

    result = {
        "ok": True,
        "project": project,
        "run_id": run_id,
        "approval_id": approval_id,
        "source_id": source_id,
        "status": "FROZEN",
        "approval_type": "PO_USAGE_STATUS_FREEZE",
        "risk_override": True,
        "authority": {"path": str(authority_path), "sha256": authority_hash},
        "draft": {"path": str(draft_path), "sha256": draft_hash},
        "frozen": {"path": str(output_path), "sha256": frozen_hash},
        "approved_facts": approved_facts,
        "active_candidates": active_candidates,
        "approved_candidates": approved_candidates,
        "step2": {
            "candidate_total": validation.get("candidate_total"),
            "candidate_exact": po_master.get("candidate_exact"),
            "validation_blocked": validation.get("blocked_count"),
            "risk_count": closing.get("risk_count"),
            "deferred_count": closing.get("deferred_count"),
        },
        "note": note,
    }
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="Freeze an authorized PO Usage Status DRAFT")
    parser.add_argument("--project", default=PROJECT)
    parser.add_argument("--draft", required=True)
    parser.add_argument("--authority", required=True)
    parser.add_argument("--step2-json", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--manifest", required=True)
    args = parser.parse_args()
    print(json.dumps(freeze_snapshot(
        args.project, args.draft, args.authority, args.step2_json,
        args.output, args.manifest,
    ), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
