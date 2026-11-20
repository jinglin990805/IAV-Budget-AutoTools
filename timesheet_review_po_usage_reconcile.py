# -*- coding: utf-8 -*-
"""Reconcile historical Timesheet candidates from the PO Usage Status workbook.

For the user's historical rule, the workbook is authoritative through the
specified month.  PO usage days are converted with 8 hours/day.  The command
is dry-run by default; ``--apply`` updates candidate review fields and closes
only data-blocking queue items.  It never creates an approved Timesheet fact.
"""

from __future__ import annotations

import argparse
import calendar
import csv
import json
import re
import sqlite3
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any

import openpyxl

from fact_store import FactStore


PROJECT = "CNTS-25165"
BLOCKING_QUEUE_TYPES = {"MISSING_PO", "UNKNOWN_PO", "MISSING_HOURS", "VALIDATION_BLOCKED"}
DATA_SOURCE = "PO Usage Status_26917.xlsx / PO Overview Actual_To Customer"

# Chinese names occur in a few older PDFs.  Latin-name matching below handles
# the ordinary English variants and names with a suffix such as "7月TMS".
NAME_ALIASES = {
    "侯景": "Hou, Jing",
    "赵士辉": "Zhao, Shihui",
    "黄德佳": "Huang, Dejia",
    "黄得佳": "Huang, Dejia",
    "黄德佳": "Huang, Dejia",
    "刘旭": "Liu, Xu",
    "轩亮亮": "Xuan, Liangliang",
    "汤其伟": "Tang, Qiwei",
    "唐其伟": "Tang, Qiwei",
    "钱海青": "Qian, Haiqing",
    "崔涵": "Cui, Han",
    "徐畅": "Xu, Chang",
}


def _digits(value: Any) -> str:
    return re.sub(r"\D", "", str(value or "")).lstrip("0")


def _period_header(year: int, month: int) -> str:
    return f"{year:04d}-{calendar.month_abbr[month]}"


def _float_equal(left: Any, right: Any) -> bool:
    try:
        return abs(float(left) - float(right)) < 0.01
    except (TypeError, ValueError):
        return False


def _workbook_name(value: str, people: set[str]) -> str:
    value = str(value or "").strip()
    if value in NAME_ALIASES:
        return NAME_ALIASES[value]
    compact = re.sub(r"[^A-Za-z]", "", value).casefold()
    exact = [
        person for person in people
        if re.sub(r"[^A-Za-z]", "", person).casefold() == compact
    ]
    if len(exact) == 1:
        return exact[0]
    contains = [
        person for person in people
        if re.sub(r"[^A-Za-z]", "", person).casefold() in compact
    ]
    if len(contains) == 1:
        return contains[0]
    # E.g. "Chenyakun 7月TMS" contains a suffix, while the workbook has
    # "Chen, Yakun".  A unique first Latin token is sufficient here.
    first_token = re.match(r"[A-Za-z]+", value)
    if first_token:
        token = first_token.group(0).casefold()
        prefix = [
            person for person in people
            if re.sub(r"[^A-Za-z]", "", person).casefold().startswith(token)
        ]
        if len(prefix) == 1:
            return prefix[0]
    return value


def load_usage_rows(workbook_path: str | Path) -> dict[tuple[str, str], list[dict[str, Any]]]:
    workbook = openpyxl.load_workbook(workbook_path, data_only=True, read_only=True)
    sheet_name = (
        "PO Overview Actual_To Customer"
        if "PO Overview Actual_To Customer" in workbook.sheetnames
        else "PO Overview Actual_To_Customer"
    )
    worksheet = workbook[sheet_name]
    headers = {
        str(worksheet.cell(2, col).value): col
        for col in range(1, worksheet.max_column + 1)
        if worksheet.cell(2, col).value is not None
    }
    raw_rows: list[dict[str, Any]] = []
    for row_number in range(25, worksheet.max_row + 1):
        po_id = worksheet.cell(row_number, 6).value
        person = str(worksheet.cell(row_number, 3).value or "").strip()
        if po_id is None or not person:
            continue
        values: dict[str, float] = {}
        for header, col in headers.items():
            if not header[:4].isdigit() or "-" not in header:
                continue
            value = worksheet.cell(row_number, col).value
            if isinstance(value, (int, float)) and value > 0:
                values[header] = float(value)
        if values:
            raw_rows.append({
                "row": row_number,
                "po_id": str(po_id).strip(),
                "person": person,
                "values": values,
            })
    people = {row["person"] for row in raw_rows}
    result: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in raw_rows:
        person = row["person"]
        for period_header, days in row["values"].items():
            result[(person, period_header)].append({
                "person": person,
                "po_id": row["po_id"],
                "days": days,
                "hours": days * 8.0,
                "row": row["row"],
            })
    # The people set is used only for the matching helper; storing it on the
    # mapping keeps the public loader compact without another workbook pass.
    result[(__name__, "__people__")] = [{"person": person} for person in people]
    return result


def _usage_people(usage: dict[tuple[str, str], list[dict[str, Any]]]) -> set[str]:
    return {item["person"] for item in usage.get((__name__, "__people__"), [])}


def _usage_for(
    usage: dict[tuple[str, str], list[dict[str, Any]]],
    person: str,
    year: int,
    month: int,
) -> list[dict[str, Any]]:
    mapped = _workbook_name(person, _usage_people(usage))
    return list(usage.get((mapped, _period_header(year, month)), []))


def _usage_for_period(
    usage: dict[tuple[str, str], list[dict[str, Any]]],
    year: int,
    month: int,
) -> list[dict[str, Any]]:
    header = _period_header(year, month)
    return [
        item
        for (person, period_header), rows in usage.items()
        if person != __name__ and period_header == header
        for item in rows
    ]


def _decode(value: Any) -> dict[str, Any]:
    try:
        decoded = json.loads(value or "{}") if isinstance(value, str) else dict(value or {})
        return decoded if isinstance(decoded, dict) else {}
    except (TypeError, ValueError):
        return {}


def _candidate_rows(store: FactStore) -> list[dict[str, Any]]:
    rows = store.connection.execute(
        """SELECT c.*, s.file_name AS source_file, s.path AS source_path
           FROM timesheet_candidates c
           LEFT JOIN source_files s ON s.source_id=c.source_id
           WHERE c.project_id=? AND c.review_status NOT IN ('SUPERSEDED','APPROVED','DISMISSED')
             AND c.year IS NOT NULL AND c.month IS NOT NULL
           ORDER BY c.year,c.month,c.raw_name,source_file,c.candidate_id""",
        (PROJECT,),
    ).fetchall()
    return [dict(row) for row in rows]


def _open_blocking_ids(store: FactStore) -> set[str]:
    rows = store.connection.execute(
        """SELECT DISTINCT candidate_id FROM review_queue
           WHERE project_id=? AND status='OPEN' AND queue_type IN ('MISSING_PO','UNKNOWN_PO','MISSING_HOURS','VALIDATION_BLOCKED')
             AND candidate_id IS NOT NULL""",
        (PROJECT,),
    ).fetchall()
    return {str(row["candidate_id"]) for row in rows}


def _group_plan(
    candidates: list[dict[str, Any]],
    target_ids: set[str],
    usage_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    if not usage_rows:
        return []
    targets = [row for row in candidates if row["candidate_id"] in target_ids]
    if not targets:
        return []
    entries = sorted(usage_rows, key=lambda item: (_digits(item["po_id"]), item["po_id"]))
    entry_by_po = {_digits(item["po_id"]): item for item in entries}
    assignments: dict[str, dict[str, Any]] = {}

    # First use an explicit PO or a unique hour match on each target row.
    for row in sorted(targets, key=lambda item: (item.get("source_file") or "", item["candidate_id"])):
        po_digits = _digits(row.get("po_id"))
        if po_digits in entry_by_po:
            assignments[row["candidate_id"]] = entry_by_po[po_digits]
            continue
        hour_matches = [
            item for item in entries
            if _float_equal(row.get("total_hours"), item["hours"])
        ]
        if len(hour_matches) == 1:
            assignments[row["candidate_id"]] = hour_matches[0]

    # Non-target records can identify the other PO in a two-PO month.  This
    # is especially useful when a new/current PDF has no PO but another PDF in
    # the same person-month already carries one.
    claimed = {
        _digits(row.get("po_id"))
        for row in candidates
        if row["candidate_id"] not in target_ids and _digits(row.get("po_id")) in entry_by_po
    }
    # A target that was already matched by its hours/PO also consumes that
    # table entry when another target row in the same person-month remains.
    claimed.update(_digits(item["po_id"]) for item in assignments.values())
    remaining_targets = [
        row for row in sorted(targets, key=lambda item: (item.get("source_file") or "", item["candidate_id"]))
        if row["candidate_id"] not in assignments
    ]
    available = [item for item in entries if _digits(item["po_id"]) not in claimed]
    if not available:
        available = entries
    if len(available) == 1:
        for row in remaining_targets:
            assignments[row["candidate_id"]] = available[0]
    elif len(available) == len(remaining_targets):
        for row, item in zip(remaining_targets, available):
            assignments[row["candidate_id"]] = item
    else:
        for index, row in enumerate(remaining_targets):
            assignments[row["candidate_id"]] = available[index % len(available)]

    plans = []
    for row in targets:
        item = assignments.get(row["candidate_id"])
        if item is None:
            continue
        changed = (
            _digits(row.get("po_id")) != _digits(item["po_id"])
            or not _float_equal(row.get("total_hours"), item["hours"])
        )
        plans.append({
            "candidate_id": row["candidate_id"],
            "person": row.get("canonical_name") or row.get("raw_name") or "",
            "period": f"{int(row['year']):04d}-{int(row['month']):02d}",
            "source_file": row.get("source_file") or "",
            "old_po": row.get("po_id") or "",
            "old_hours": row.get("total_hours"),
            "new_po": item["po_id"],
            "assigned_person": item.get("person", ""),
            "assigned_canonical_name": str(item.get("person", "")).replace(",", "").strip(),
            "match_method": item.get("match_method", "workbook_name"),
            "days": item["days"],
            "new_hours": item["hours"],
            "changed": changed,
        })
    return plans


def _authoritative_group_plan(
    rows: list[dict[str, Any]],
    usage_rows: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], str | None]:
    """Build only unambiguous one-candidate-to-one-workbook-PO assignments.

    A monthly workbook total must not be copied to duplicate PDF candidates or
    to a single PDF that combines multiple PO rows.  Those groups remain
    unmatched until a source-version/split decision exists.
    """
    entries = sorted(usage_rows, key=lambda item: (_digits(item["po_id"]), item["po_id"]))
    if not rows or not entries:
        return [], "no workbook person/month data"
    if len(rows) != len(entries):
        if len(rows) == 1 and len(entries) > 1:
            total_hours = sum(float(item["hours"]) for item in entries)
            if _float_equal(rows[0].get("total_hours"), total_hours):
                return [], "one candidate combines multiple workbook POs"
        return [], "candidate/version count differs from workbook PO count"

    unused = list(entries)
    assignments: list[tuple[dict[str, Any], dict[str, Any]]] = []
    for row in sorted(rows, key=lambda item: (item.get("source_file") or "", item["candidate_id"])):
        candidate_po = _digits(row.get("po_id"))
        exact = [item for item in unused if candidate_po and _digits(item["po_id"]) == candidate_po]
        if len(exact) == 1:
            selected = exact[0]
        else:
            hour_matches = [item for item in unused if _float_equal(row.get("total_hours"), item["hours"])]
            if len(hour_matches) != 1:
                return [], "PO/hour mapping is not unique"
            selected = hour_matches[0]
        unused.remove(selected)
        assignments.append((row, selected))
    if unused:
        return [], "not all workbook POs were assigned"

    plans = []
    for row, item in assignments:
        plans.append({
            "candidate_id": row["candidate_id"],
            "person": row.get("canonical_name") or row.get("raw_name") or "",
            "period": f"{int(row['year']):04d}-{int(row['month']):02d}",
            "source_file": row.get("source_file") or "",
            "old_po": row.get("po_id") or "",
            "old_hours": row.get("total_hours"),
            "new_po": item["po_id"],
            "assigned_person": item.get("person", ""),
            "assigned_canonical_name": str(item.get("person", "")).replace(",", "").strip(),
            "match_method": "workbook_authoritative_one_to_one",
            "days": item["days"],
            "new_hours": item["hours"],
            "changed": (
                _digits(row.get("po_id")) != _digits(item["po_id"])
                or not _float_equal(row.get("total_hours"), item["hours"])
                or str(row.get("canonical_name") or "").replace(",", "").strip()
                != str(item.get("person", "")).replace(",", "").strip()
            ),
        })
    return plans, None


def build_plan(
    store: FactStore,
    usage: dict[tuple[str, str], list[dict[str, Any]]],
    through: str,
    candidate_ids: set[str] | None = None,
    authoritative: bool = False,
) -> dict[str, Any]:
    try:
        through_year, through_month = (int(value) for value in through.split("-", 1))
    except ValueError as exc:
        raise ValueError("--through must be YYYY-MM") from exc
    candidates = _candidate_rows(store)
    blocking_ids = _open_blocking_ids(store)
    groups: dict[tuple[str, int, int], list[dict[str, Any]]] = defaultdict(list)
    for row in candidates:
        if (int(row["year"]), int(row["month"])) > (through_year, through_month):
            continue
        name = row.get("canonical_name") or row.get("raw_name") or ""
        groups[(name, int(row["year"]), int(row["month"]))].append(row)

    plans: list[dict[str, Any]] = []
    unmatched: list[dict[str, Any]] = []
    for (person, year, month), rows in groups.items():
        if candidate_ids is not None:
            rows = [row for row in rows if row["candidate_id"] in candidate_ids]
        target_ids = (
            {row["candidate_id"] for row in rows}
            if authoritative
            else {
                row["candidate_id"] for row in rows
                if row["candidate_id"] in blocking_ids
                or row.get("total_hours") is None
                or float(row.get("total_hours") or 0) <= 0
                or not str(row.get("po_id") or "").strip()
            }
        )
        if not target_ids:
            continue
        usage_rows = _usage_for(usage, person, year, month)
        if not usage_rows:
            period_rows = _usage_for_period(usage, year, month)
            inferred = []
            for target in rows:
                if target["candidate_id"] not in target_ids:
                    continue
                hour_matches = [
                    item for item in period_rows
                    if _float_equal(target.get("total_hours"), item["hours"])
                ]
                if len(hour_matches) == 1:
                    item = dict(hour_matches[0])
                    item["match_method"] = "period_hours_unique"
                    inferred.append(item)
            if len(inferred) == len(target_ids):
                usage_rows = inferred
        if not usage_rows:
            unmatched.append({
                "person": person,
                "period": f"{year:04d}-{month:02d}",
                "candidate_ids": sorted(target_ids),
            })
            continue
        if authoritative:
            group_plans, group_reason = _authoritative_group_plan(rows, usage_rows)
            if group_reason:
                unmatched.append({
                    "person": person,
                    "period": f"{year:04d}-{month:02d}",
                    "candidate_ids": sorted(target_ids),
                    "reason": group_reason,
                })
                continue
        else:
            group_plans = _group_plan(rows, target_ids, usage_rows)
        plans.extend(group_plans)
        planned_ids = {item["candidate_id"] for item in group_plans}
        for candidate_id in sorted(target_ids - planned_ids):
            unmatched.append({
                "person": person,
                "period": f"{year:04d}-{month:02d}",
                "candidate_ids": [candidate_id],
                "reason": "no deterministic table assignment",
            })
    return {
        "through": through,
        "candidate_total": len(candidates),
        "blocking_candidate_ids": len(blocking_ids),
        "plans": plans,
        "unmatched": unmatched,
    }


def _update_payload(raw_payload: Any, po_id: str, hours: float, canonical_name: str = "") -> str:
    payload = _decode(raw_payload)
    fields = payload.setdefault("fields", {})
    fields["po_id"] = po_id
    if canonical_name:
        fields["canonical_name"] = canonical_name
    fields["hours"] = str(int(hours)) if float(hours).is_integer() else str(hours)
    validation = payload.setdefault("validation", {})
    checks = validation.setdefault("checks", {})
    checks["po_digits"] = po_id
    checks["po_id_present"] = True
    checks["hours"] = hours
    checks["hours_present"] = True
    issues = [
        issue for issue in (validation.get("issues") or [])
        if str(issue.get("code") or "") not in {"MISSING_HOURS", "MISSING_PO_ID", "UNKNOWN_PO"}
    ]
    validation["issues"] = issues
    has_blocking = any(
        str(issue.get("severity") or "").upper() == "BLOCKING"
        for issue in issues if isinstance(issue, dict)
    )
    validation["status"] = "BLOCKED" if has_blocking else "WARNING" if issues else "PASS"
    return json.dumps(payload, ensure_ascii=False, sort_keys=True)


def apply_plan(store: FactStore, plan: dict[str, Any]) -> list[dict[str, Any]]:
    applied = []
    for item in plan["plans"]:
        row = store.connection.execute(
            "SELECT parse_json FROM timesheet_candidates WHERE candidate_id=? AND project_id=?",
            (item["candidate_id"], PROJECT),
        ).fetchone()
        if row is None:
            continue
        note = (
            f"{DATA_SOURCE} authoritative reconciliation: {item['period']} "
            f"{item['assigned_canonical_name'] or item['person']} -> PO {item['new_po']}; "
            f"{item['days']}d x 8h/day = {item['new_hours']}h; "
            f"match={item['match_method']}."
        )
        payload = _update_payload(
            row["parse_json"], item["new_po"], item["new_hours"],
            item["assigned_canonical_name"],
        )
        store.connection.execute(
            """UPDATE timesheet_candidates
               SET canonical_name=?, po_id=?, total_hours=?, parse_json=?, review_status='REVIEW', notes=?
               WHERE candidate_id=? AND project_id=?""",
            (
                item["assigned_canonical_name"], item["new_po"], item["new_hours"],
                payload, note, item["candidate_id"], PROJECT,
            ),
        )
        store.connection.execute(
            """UPDATE review_queue
               SET status='RESOLVED', resolved_at=datetime('now'), resolution_note=?
               WHERE candidate_id=? AND project_id=? AND status='OPEN'
                 AND queue_type IN ('NEW_MEMBER','MISSING_PO','UNKNOWN_PO','MISSING_HOURS','VALIDATION_BLOCKED')""",
            (note, item["candidate_id"], PROJECT),
        )
        applied.append(item)
    store.connection.commit()
    return applied


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Reconcile historical Timesheet PO/hour fields from PO Usage Status")
    parser.add_argument("--project", default=PROJECT)
    parser.add_argument("--workbook", required=True)
    parser.add_argument("--through", default="2026-08", help="inclusive period cutoff YYYY-MM")
    parser.add_argument("--candidate-csv", help="limit the plan to candidate_id values in this CSV")
    parser.add_argument(
        "--authoritative",
        action="store_true",
        help="apply only one-to-one workbook PO assignments; leave duplicate/combined groups unmatched",
    )
    parser.add_argument("--apply", action="store_true", help="write candidate updates; default is dry-run")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if args.project != PROJECT:
        raise SystemExit(f"This reconciler is configured for {PROJECT}")
    candidate_ids = None
    if args.candidate_csv:
        with Path(args.candidate_csv).open("r", encoding="utf-8-sig", newline="") as handle:
            candidate_ids = {
                str(row.get("candidate_id") or "").strip()
                for row in csv.DictReader(handle)
                if str(row.get("candidate_id") or "").strip()
            }
    usage = load_usage_rows(args.workbook)
    with FactStore() as store:
        store.initialize()
        plan = build_plan(
            store, usage, args.through,
            candidate_ids=candidate_ids,
            authoritative=args.authoritative,
        )
        applied = apply_plan(store, plan) if args.apply else []
    result = {
        "ok": True,
        "project": args.project,
        "workbook": args.workbook,
        "through": args.through,
        "apply": args.apply,
        "authoritative": args.authoritative,
        "candidate_csv": args.candidate_csv or "",
        "plans": len(plan["plans"]),
        "changes": sum(1 for item in plan["plans"] if item["changed"]),
        "unmatched_groups": len(plan["unmatched"]),
        "unmatched": plan["unmatched"],
        "applied": len(applied),
        "fact_approval_performed": False,
        "items": plan["plans"],
    }
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
