# -*- coding: utf-8 -*-
"""Read-only dry-run audit for a CNTS-25165 Cost Tracking workbook."""

from __future__ import annotations

import argparse
import hashlib
import json
import sqlite3
import sys
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


PROJECT = "CNTS-25165"
WBS_MAP = {
    "10026": "CNTS-25165-10000",
    "11026": "CNTS-25165-11000",
}
MONTHS = ["2026-" + f"{month:02d}" for month in range(1, 13)]
MONTH_NUMBERS = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
}


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _number(value: Any) -> float | None:
    if value is None or value == "" or value == "-":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _parse_sheet(ws: Any, sheet_name: str) -> dict[str, Any]:
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError(f"Empty Cost Tracking sheet: {sheet_name}")
    header = [str(value).strip() if value is not None else "" for value in rows[0]]
    month_indexes = {header[index]: index for index in range(len(header)) if header[index] in {
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
    }}
    if len(month_indexes) != 12:
        raise ValueError(f"Expected Jan-Dec columns in {sheet_name}, got {sorted(month_indexes)}")

    hour_start = next(
        (index for index, row in enumerate(rows) if str(row[0]).strip().lower() == "hours"),
        None,
    )
    if hour_start is None:
        raise ValueError(f"Hours section not found in {sheet_name}")

    cost_rows: list[dict[str, Any]] = []
    for row in rows[1:hour_start]:
        employee_id = str(row[0]).strip() if row[0] is not None else ""
        name = str(row[1]).strip() if row[1] is not None else ""
        if not name or name.upper() == "TOTAL":
            continue
        monthly = {
            f"2026-{MONTH_NUMBERS[month_name]:02d}": _number(row[index])
            for month_name, index in month_indexes.items()
        }
        cost_rows.append({
            "employee_id": employee_id,
            "name": name,
            "sartyp": str(row[2]).strip() if row[2] is not None else "",
            "hourly_rate": _number(row[3]),
            "monthly_cost": monthly,
            "total_cost": _number(row[16]),
        })

    hours_rows: list[dict[str, Any]] = []
    for row in rows[hour_start + 1:]:
        employee_id = str(row[0]).strip() if row[0] is not None else ""
        name = str(row[1]).strip() if row[1] is not None else ""
        if not name or name.upper() == "TOTAL":
            continue
        monthly = {
            f"2026-{MONTH_NUMBERS[month_name]:02d}": _number(row[index])
            for month_name, index in month_indexes.items()
        }
        hours_rows.append({
            "employee_id": employee_id,
            "name": name,
            "monthly_hours": monthly,
            "total_hours": _number(row[16]),
        })

    cost_by_employee = {(item["employee_id"], item["name"]): item for item in cost_rows}
    hours_by_employee = {(item["employee_id"], item["name"]): item for item in hours_rows}
    checks: list[dict[str, Any]] = []
    records: list[dict[str, Any]] = []
    for key, cost in cost_by_employee.items():
        if cost["name"].lower() in {"other expenses", "subcontr. gmbh"} or not cost["employee_id"]:
            continue
        hours = hours_by_employee.get(key)
        if not hours:
            checks.append({"type": "MISSING_HOURS_ROW", "employee": cost["name"]})
            continue
        for month in MONTHS:
            amount = cost["monthly_cost"].get(month)
            hours_value = hours["monthly_hours"].get(month)
            if amount is None and hours_value is None:
                continue
            if amount is not None and hours_value is not None and cost["hourly_rate"] is not None:
                expected = round(cost["hourly_rate"] * hours_value, 2)
                if abs(expected - amount) > 0.02:
                    checks.append({
                        "type": "RATE_HOURS_MISMATCH",
                        "employee": cost["name"],
                        "month": month,
                        "amount": amount,
                        "hours": hours_value,
                        "hourly_rate": cost["hourly_rate"],
                        "expected": expected,
                    })
            records.append({
                "wbs": WBS_MAP.get(sheet_name),
                "sheet": sheet_name,
                "month": month,
                "cost_layer": "internal_labor_actual",
                "employee_id": cost["employee_id"],
                "employee_name": cost["name"],
                "hours": hours_value or 0.0,
                "amount": amount or 0.0,
                "currency": "CNY",
            })

    other = next((item for item in cost_rows if item["name"].lower() == "other expenses"), None)
    if other:
        for month, amount in other["monthly_cost"].items():
            if amount is not None:
                records.append({
                    "wbs": WBS_MAP.get(sheet_name),
                    "sheet": sheet_name,
                    "month": month,
                    "cost_layer": "other_expenses_actual",
                    "employee_id": None,
                    "employee_name": "Other Expenses",
                    "hours": 0.0,
                    "amount": amount,
                    "currency": "CNY",
                })

    duplicate_keys: dict[tuple[Any, ...], int] = {}
    for record in records:
        key = (
            record["wbs"], record["month"], record["cost_layer"],
            record["employee_id"], record["employee_name"],
        )
        duplicate_keys[key] = duplicate_keys.get(key, 0) + 1
    duplicates = [
        {"key": list(key), "count": count}
        for key, count in duplicate_keys.items() if count > 1
    ]
    return {
        "sheet": sheet_name,
        "wbs": WBS_MAP.get(sheet_name),
        "cost_rows": cost_rows,
        "hours_rows": hours_rows,
        "records": records,
        "record_count": len(records),
        "duplicate_count": len(duplicates),
        "duplicates": duplicates,
        "checks": checks,
        "cost_total": sum(item["total_cost"] or 0.0 for item in cost_rows),
        "hours_total": sum(item["total_hours"] or 0.0 for item in hours_rows),
    }


def audit(source: str | Path, po_file: str | Path | None = None) -> dict[str, Any]:
    source_path = Path(source).resolve()
    if not source_path.is_file():
        raise FileNotFoundError(source_path)
    wb = load_workbook(source_path, data_only=True, read_only=True)
    try:
        sheets = [_parse_sheet(wb[sheet], sheet) for sheet in wb.sheetnames]
    finally:
        wb.close()

    db_path = Path(__file__).resolve().parent / "data" / "iav_facts.sqlite3"
    with sqlite3.connect(db_path) as connection:
        connection.row_factory = sqlite3.Row
        existing_facts = connection.execute(
            "SELECT COUNT(*) FROM cost_facts WHERE project_id=?", (PROJECT,)
        ).fetchone()[0]
        existing_sources = connection.execute(
            "SELECT source_id, path, sha256 FROM source_files WHERE project_id=? AND sha256=?",
            (PROJECT, _sha256(source_path)),
        ).fetchall()

    po_comparison: list[dict[str, Any]] = []
    if po_file:
        po_path = Path(po_file).resolve()
        po_wb = load_workbook(po_path, data_only=True, read_only=True)
        try:
            ws = po_wb["Sheet1"]
            for row in ws.iter_rows(min_row=5, values_only=True):
                if str(row[10]).strip() not in {"CNTS-25165"}:
                    continue
                wbs = str(row[11]).strip()
                po_comparison.append({
                    "wbs": wbs,
                    "actual_expenses": _number(row[23]),
                    "actual_material": _number(row[24]),
                    "actual_activity_allocation": _number(row[25]),
                })
        finally:
            po_wb.close()

    total_cost = sum(item["cost_total"] for item in sheets)
    total_hours = sum(item["hours_total"] for item in sheets)
    all_records = [record for item in sheets for record in item["records"]]
    all_checks = [check | {"sheet": item["sheet"]} for item in sheets for check in item["checks"]]
    all_duplicates = [duplicate | {"sheet": item["sheet"]} for item in sheets for duplicate in item["duplicates"]]
    return {
        "ok": True,
        "project": PROJECT,
        "source": {
            "path": str(source_path),
            "sha256": _sha256(source_path),
            "size_bytes": source_path.stat().st_size,
            "sheet_names": wb.sheetnames if False else [item["sheet"] for item in sheets],
        },
        "sheets": sheets,
        "record_count": len(all_records),
        "total_cost": total_cost,
        "total_hours": total_hours,
        "arithmetic_check_count": len(all_checks),
        "arithmetic_checks": all_checks,
        "duplicate_count": len(all_duplicates),
        "duplicates": all_duplicates,
        "database": {
            "existing_cost_facts": existing_facts,
            "same_hash_sources": [dict(row) for row in existing_sources],
        },
        "po_file_comparison": po_comparison,
        "status": "DRY_RUN_READY_FOR_REVIEW" if not all_checks and not all_duplicates else "DRY_RUN_REVIEW_REQUIRED",
        "write_applied": False,
    }


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Audit Cost Tracking workbook without writing facts")
    parser.add_argument("--source", required=True)
    parser.add_argument("--po-file")
    args = parser.parse_args()
    print(json.dumps(audit(args.source, args.po_file), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
