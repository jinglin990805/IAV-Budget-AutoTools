# -*- coding: utf-8 -*-
"""Export controlled historical Timesheet review batches for business review.

The export is deliberately a review artifact, not an approval import.  It
contains a blank decision column so a reviewer can record the decision without
silently changing the local fact store.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from pathlib import Path
from typing import Any, Iterable

from fact_store import FactStore
from project_registry import build_profile
from review_scope import path_in_scope, review_scope
from timesheet_review_batches import classify_candidate


EXPORT_COLUMNS = (
    "candidate_id",
    "period",
    "raw_name",
    "canonical_name",
    "hours",
    "candidate_po",
    "resolved_po",
    "po_suggestions",
    "batch",
    "validation_status",
    "open_issues",
    "review_status",
    "source_file",
    "source_path",
    "review_decision",
    "review_note",
)


def _decode_json(value: Any) -> Any:
    if isinstance(value, (dict, list)):
        return value
    try:
        return json.loads(value or "{}")
    except (TypeError, ValueError):
        return {}


def _review_items(store: FactStore, candidate_id: str) -> list[dict[str, Any]]:
    rows = [dict(row) for row in store.connection.execute(
        """SELECT review_id, queue_type, severity, status, payload_json,
                  created_at, resolved_at, resolution_note
           FROM review_queue WHERE candidate_id=? ORDER BY created_at""",
        (candidate_id,),
    )]
    for row in rows:
        row["payload_json"] = _decode_json(row.get("payload_json"))
    return rows


def collect_historical_candidates(
    store: FactStore,
    project: str,
    batches: set[str] | None = None,
    period: str = "",
    include_resolved: bool = False,
) -> list[dict[str, Any]]:
    """Collect candidates using the same scope boundary as the Review API."""
    scope = review_scope(project)
    status_clause = "" if include_resolved else " AND c.review_status NOT IN ('APPROVED', 'DISMISSED')"
    rows = [dict(row) for row in store.connection.execute(
        f"""SELECT c.*, s.path AS source_path, s.file_name AS source_file
            FROM timesheet_candidates c
            LEFT JOIN source_files s ON s.source_id=c.source_id
            WHERE c.project_id=? AND c.review_status <> 'SUPERSEDED'{status_clause}
            ORDER BY c.year ASC, c.month ASC, c.created_at ASC, c.candidate_id ASC""",
        (project,),
    )]
    result = []
    for candidate in rows:
        is_current = (
            candidate.get("year") == scope["year"]
            and candidate.get("month") == scope["month_number"]
            and path_in_scope(candidate.get("source_path"), scope)
        )
        if is_current:
            continue
        items = _review_items(store, candidate["candidate_id"])
        classification = classify_candidate(candidate, items)
        if batches and classification["batch"] not in batches:
            continue
        if period and classification["period"] != period:
            continue
        candidate["classification"] = classification
        candidate["review_items"] = items
        result.append(candidate)
    return result


def _suggestions(candidate: dict[str, Any]) -> str:
    values = []
    for item in candidate.get("review_items", []):
        if str(item.get("status") or "OPEN").upper() != "OPEN":
            continue
        if str(item.get("queue_type") or "").upper() != "UNKNOWN_PO":
            continue
        payload = item.get("payload_json") or {}
        for suggestion in payload.get("suggestions") or []:
            po_id = suggestion.get("po_id") or ""
            similarity = suggestion.get("similarity")
            values.append(f"{po_id} ({similarity})" if similarity is not None else str(po_id))
    return "; ".join(dict.fromkeys(values))


def export_row(candidate: dict[str, Any]) -> dict[str, Any]:
    classification = candidate.get("classification") or classify_candidate(
        candidate, candidate.get("review_items") or []
    )
    open_issues = sorted({
        str(item.get("queue_type") or "")
        for item in candidate.get("review_items", [])
        if str(item.get("status") or "OPEN").upper() == "OPEN" and item.get("queue_type")
    })
    return {
        "candidate_id": candidate.get("candidate_id", ""),
        "period": classification.get("period", ""),
        "raw_name": candidate.get("raw_name", ""),
        "canonical_name": candidate.get("canonical_name", ""),
        "hours": candidate.get("total_hours", ""),
        "candidate_po": candidate.get("po_id", ""),
        "resolved_po": "",
        "po_suggestions": _suggestions(candidate),
        "batch": classification.get("batch", ""),
        "validation_status": classification.get("validation_status", ""),
        "open_issues": "; ".join(open_issues),
        "review_status": candidate.get("review_status", ""),
        "source_file": candidate.get("source_file", ""),
        "source_path": candidate.get("source_path", ""),
        "review_decision": "",
        "review_note": "",
    }


def write_csv(rows: Iterable[dict[str, Any]], output: str | Path) -> Path:
    path = Path(output)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=EXPORT_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    return path


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Export historical Timesheet review batches")
    parser.add_argument("--project", default="CNTS-25165")
    parser.add_argument(
        "--batch",
        default="PO_POTENTIAL,PO_UNKNOWN",
        help="comma-separated primary batches; defaults to potential and unknown PO",
    )
    parser.add_argument("--period", default="", help="optional YYYY-MM filter")
    parser.add_argument("--include-resolved", action="store_true")
    parser.add_argument("--output", default="")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    batches = {value.strip().upper() for value in args.batch.split(",") if value.strip()}
    if not args.output:
        label = "-".join(sorted(batches)).lower() or "all"
        args.output = str(Path("output") / f"timesheet_review_{args.project}_{label}.csv")

    with FactStore() as store:
        store.initialize()
        rows = [export_row(candidate) for candidate in collect_historical_candidates(
            store, args.project, batches=batches, period=args.period,
            include_resolved=args.include_resolved,
        )]
    output = write_csv(rows, args.output)
    summary = {
        "ok": True,
        "project": args.project,
        "batches": sorted(batches),
        "period": args.period or None,
        "rows": len(rows),
        "output": str(output),
        "decision_columns_blank": True,
    }
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
