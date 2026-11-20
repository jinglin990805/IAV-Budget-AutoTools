# -*- coding: utf-8 -*-
"""Validate and explicitly apply historical Timesheet review decisions.

The default command is a dry-run.  ``--apply`` is required for any mutation,
and even applied PO decisions only update the candidate/review queue; they never
create an approved Timesheet fact.  Historical fact approval remains a separate
explicit workflow.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path
from typing import Any

from fact_store import FactStore
from project_registry import build_profile
from review_scope import path_in_scope, review_scope


DECISIONS = {"CONFIRM_PO", "CORRECT_PO", "DISMISS", "NEEDS_SOURCE"}
REQUIRED_COLUMNS = {"candidate_id", "review_decision", "review_note"}


def read_decision_csv(path: str | Path) -> list[dict[str, str]]:
    with Path(path).open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fields = set(reader.fieldnames or [])
        missing = REQUIRED_COLUMNS - fields
        if missing:
            raise ValueError(f"CSV missing required columns: {', '.join(sorted(missing))}")
        return [{str(key): str(value or "").strip() for key, value in row.items()} for row in reader]


def _digits(value: Any) -> str:
    return re.sub(r"\D", "", str(value or ""))


def validate_decision_rows(
    rows: list[dict[str, str]],
    candidates: dict[str, dict[str, Any]],
    known_po_ids: set[str],
) -> list[dict[str, Any]]:
    """Validate CSV decisions without opening a database or changing state."""
    po_lookup = {_digits(value): str(value) for value in known_po_ids if _digits(value)}
    seen: set[str] = set()
    result = []
    for row_number, row in enumerate(rows, 2):
        candidate_id = row.get("candidate_id", "")
        decision = row.get("review_decision", "").upper().replace(" ", "_")
        note = row.get("review_note", "")
        resolved_input = row.get("resolved_po", "")
        item = {
            "row": row_number,
            "candidate_id": candidate_id,
            "decision": decision,
            "resolved_po": "",
            "note": note,
            "status": "PENDING" if not decision else "VALID",
            "errors": [],
        }
        if not candidate_id:
            item["status"] = "ERROR"
            item["errors"].append("candidate_id is required")
        elif candidate_id in seen:
            item["status"] = "ERROR"
            item["errors"].append("duplicate candidate_id")
        else:
            seen.add(candidate_id)
        candidate = candidates.get(candidate_id)
        if candidate is None and candidate_id:
            item["status"] = "ERROR"
            item["errors"].append("candidate not found in project")
        elif candidate and candidate.get("review_status") in {"APPROVED", "SUPERSEDED", "DISMISSED"}:
            item["status"] = "ERROR"
            item["errors"].append(f"candidate status {candidate['review_status']} cannot be changed")

        if not decision:
            result.append(item)
            continue
        if decision not in DECISIONS:
            item["status"] = "ERROR"
            item["errors"].append(f"unsupported decision: {decision}")
        if decision in {"CONFIRM_PO", "CORRECT_PO"}:
            canonical = po_lookup.get(_digits(resolved_input))
            if not canonical:
                item["status"] = "ERROR"
                item["errors"].append("resolved_po must match a parsed PO Master record")
            else:
                item["resolved_po"] = canonical
        if decision in {"DISMISS", "NEEDS_SOURCE"} and not note:
            item["status"] = "ERROR"
            item["errors"].append("review_note is required for this decision")
        if item["errors"]:
            item["status"] = "ERROR"
        result.append(item)
    return result


def _candidate_rows(store: FactStore, project: str) -> dict[str, dict[str, Any]]:
    scope = review_scope(project)
    rows = [dict(row) for row in store.connection.execute(
        """SELECT c.*, s.path AS source_path
           FROM timesheet_candidates c
           LEFT JOIN source_files s ON s.source_id=c.source_id
           WHERE c.project_id=? AND c.review_status <> 'SUPERSEDED'""",
        (project,),
    )]
    result = {}
    for row in rows:
        current = (
            row.get("year") == scope["year"]
            and row.get("month") == scope["month_number"]
            and path_in_scope(row.get("source_path"), scope)
        )
        row["historical_scope"] = not current
        result[row["candidate_id"]] = row
    return result


def _parsed_po_ids(store: FactStore, project: str) -> set[str]:
    return {
        str(row["po_id"])
        for row in store.connection.execute(
            """SELECT po_id FROM po_master_records
               WHERE project_id=? AND parse_status='PARSED' AND po_id IS NOT NULL AND po_id<>''""",
            (project,),
        )
    }


def apply_decisions(
    store: FactStore,
    project: str,
    decisions: list[dict[str, Any]],
    candidates: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    """Apply only valid historical decisions; never approve a fact."""
    applied = []
    for item in decisions:
        if item["status"] != "VALID":
            continue
        candidate = candidates.get(item["candidate_id"])
        if not candidate or not candidate.get("historical_scope"):
            item["status"] = "ERROR"
            item["errors"].append("only historical candidates may be imported")
            continue
        note = item["note"] or "历史 PO 复核决策"
        if item["decision"] in {"CONFIRM_PO", "CORRECT_PO"}:
            payload = {}
            try:
                payload = json.loads(candidate.get("parse_json") or "{}")
            except (TypeError, ValueError):
                payload = {}
            fields = payload.setdefault("fields", {})
            fields["po_id"] = item["resolved_po"]
            validation = payload.get("validation")
            if isinstance(validation, dict):
                checks = validation.setdefault("checks", {})
                checks["po_digits"] = item["resolved_po"]
                checks["po_id_present"] = True
                validation["issues"] = [
                    issue for issue in (validation.get("issues") or [])
                    if str(issue.get("code") or "") not in {"MISSING_PO_ID", "UNKNOWN_PO"}
                ]
                remaining = validation["issues"]
                has_blocking = any(
                    str(issue.get("severity") or "").upper() == "BLOCKING"
                    for issue in remaining
                    if isinstance(issue, dict)
                )
                validation["status"] = (
                    "BLOCKED" if has_blocking else "WARNING" if remaining else "PASS"
                )
            store.connection.execute(
                """UPDATE timesheet_candidates
                   SET po_id=?, parse_json=?, review_status='REVIEW', notes=?
                   WHERE candidate_id=? AND project_id=?""",
                (item["resolved_po"], json.dumps(payload, ensure_ascii=False, sort_keys=True), note,
                 item["candidate_id"], project),
            )
            store.connection.execute(
                """UPDATE review_queue SET status='RESOLVED', resolved_at=datetime('now'), resolution_note=?
                   WHERE candidate_id=? AND status='OPEN' AND queue_type IN ('MISSING_PO', 'UNKNOWN_PO')""",
                (note, item["candidate_id"]),
            )
        elif item["decision"] == "DISMISS":
            store.dismiss_candidate(item["candidate_id"], note)
        elif item["decision"] == "NEEDS_SOURCE":
            store.connection.execute(
                """UPDATE timesheet_candidates SET notes=?, review_status='REVIEW'
                   WHERE candidate_id=? AND project_id=?""",
                (note, item["candidate_id"], project),
            )
        applied.append({
            "candidate_id": item["candidate_id"],
            "decision": item["decision"],
            "resolved_po": item["resolved_po"],
        })
    store.connection.commit()
    return applied


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Validate/apply historical Timesheet review decisions")
    parser.add_argument("--project", default="CNTS-25165")
    parser.add_argument("--input", required=True, help="review CSV exported by timesheet_review_export.py")
    parser.add_argument("--apply", action="store_true", help="apply valid historical decisions; default is dry-run")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        rows = read_decision_csv(args.input)
        with FactStore() as store:
            store.initialize()
            candidates = _candidate_rows(store, args.project)
            decisions = validate_decision_rows(rows, candidates, _parsed_po_ids(store, args.project))
            applied = apply_decisions(store, args.project, decisions, candidates) if args.apply else []
        summary = {
            "ok": not any(item["status"] == "ERROR" for item in decisions),
            "project": args.project,
            "input": args.input,
            "apply": args.apply,
            "rows": len(decisions),
            "pending": sum(item["status"] == "PENDING" for item in decisions),
            "valid": sum(item["status"] == "VALID" for item in decisions),
            "errors": sum(item["status"] == "ERROR" for item in decisions),
            "applied": applied,
            "mutations": len(applied),
            "fact_approval_performed": False,
            "items": [item for item in decisions if item["status"] != "VALID" or args.apply],
        }
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
        print(json.dumps(summary, ensure_ascii=False))
        return 0 if summary["ok"] else 2
    except (OSError, ValueError) as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
