# -*- coding: utf-8 -*-
"""Create a COM-free PO Usage Status fact audit workbook.

This is a working/audit artifact only.  It does not replace the customer
template and does not modify the fact store or the authoritative workbook.
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path
from typing import Any

import xlsxwriter

from fact_store import FactStore


PROJECT = "CNTS-25165"
HOURS_PER_DAY = 8.0


def _rows(store: FactStore, query: str, params: tuple[Any, ...]) -> list[dict[str, Any]]:
    return [dict(row) for row in store.connection.execute(query, params).fetchall()]


def _read_json(path: str | Path) -> dict[str, Any]:
    raw = Path(path).read_bytes()
    for encoding in ("utf-8-sig", "utf-16", "utf-8"):
        try:
            return json.loads(raw.decode(encoding))
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue
    raise ValueError(f"Unable to decode JSON: {path}")


def _write_table(workbook: Any, name: str, headers: list[str], rows: list[list[Any]]) -> None:
    worksheet = workbook.add_worksheet(name)
    header_format = workbook.add_format({
        "bold": True, "bg_color": "#2F5496", "font_color": "white",
        "border": 1, "align": "center", "valign": "vcenter",
    })
    cell_format = workbook.add_format({"border": 1, "valign": "vcenter"})
    number_format = workbook.add_format({"border": 1, "num_format": "#,##0.00"})
    for column, value in enumerate(headers):
        worksheet.write(0, column, value, header_format)
    for row_number, row in enumerate(rows, 1):
        for column, value in enumerate(row):
            fmt = number_format if isinstance(value, (int, float)) and not isinstance(value, bool) else cell_format
            worksheet.write(row_number, column, value, fmt)
    worksheet.freeze_panes(1, 0)
    worksheet.set_column(0, max(0, len(headers) - 1), 18)
    for column, header in enumerate(headers):
        if "source" in header.lower() or "note" in header.lower():
            worksheet.set_column(column, column, 42)
        elif "person" in header.lower() or "name" in header.lower():
            worksheet.set_column(column, column, 24)
    if rows:
        worksheet.add_table(
            0, 0, len(rows), len(headers) - 1,
            {"name": f"{name.replace('_', '')}Table", "style": "Table Style Medium 2"},
        )
    else:
        worksheet.autofilter(0, 0, 0, max(0, len(headers) - 1))


def build_workbook(
    project: str,
    authority_workbook: str,
    step2_json: str,
    output: str,
) -> dict[str, Any]:
    step2 = _read_json(step2_json)
    with FactStore() as store:
        store.initialize()
        candidates = _rows(
            store,
            """SELECT candidate_id, raw_name, canonical_name, year, month,
                      po_id, total_hours, parse_method, review_status, is_preferred
               FROM timesheet_candidates
               WHERE project_id=? AND review_status<>'SUPERSEDED'
               ORDER BY year, month, canonical_name, candidate_id""",
            (project,),
        )
        facts = _rows(
            store,
            """SELECT fact_id, person_key, month, po_id, hours, raw_unit,
                      approval_status, is_current, source_id, notes
               FROM timesheet_facts
               WHERE project_id=? AND approval_status='APPROVED' AND is_current=1
               ORDER BY month, person_key, po_id, fact_id""",
            (project,),
        )
        po_master = _rows(
            store,
            """SELECT po_id, parse_status, po_date, sap_po_number,
                      total_excl_vat, delivery_date, level_quotas_json
               FROM po_master_records
               WHERE project_id=? ORDER BY po_id, record_id""",
            (project,),
        )
        monthly = _rows(
            store,
            """SELECT period, sheet_name, row_number, raw_name,
                      canonical_name, project_number, hours, mapping_status
               FROM monthly_closing_facts
               WHERE project_id=? ORDER BY period, sheet_name, row_number""",
            (project,),
        )
        open_queue = _rows(
            store,
            """SELECT queue_type, severity, COUNT(*) AS count
               FROM review_queue
               WHERE project_id=? AND status='OPEN'
               GROUP BY queue_type, severity ORDER BY queue_type, severity""",
            (project,),
        )

    closing = step2.get("monthly_closing") or {}
    validation = step2.get("timesheet_validation") or {}
    po_summary = step2.get("po_master") or {}
    mismatches = closing.get("mismatches") or []
    risk_rows = [item for item in mismatches if item.get("risk") or item.get("deferred")]
    risk_count = int(closing.get("risk_count", sum(1 for item in risk_rows if item.get("risk"))))
    deferred_count = int(closing.get("deferred_count", sum(1 for item in risk_rows if item.get("deferred"))))
    gate_status = "READY_FOR_FREEZE" if (
        not validation.get("blocked_count") and not risk_count
        and not deferred_count and not [row for row in candidates if row["review_status"] != "APPROVED"]
    ) else "REVIEW_REQUIRED"

    audit = {
        "Project": project,
        "Generated at": datetime.now().isoformat(timespec="seconds"),
        "Authority workbook": authority_workbook,
        "Authority workbook exists": Path(authority_workbook).is_file(),
        "Hours per day": HOURS_PER_DAY,
        "Approved current facts": len(facts),
        "Active candidates": len(candidates),
        "Approved candidates": sum(1 for row in candidates if row["review_status"] == "APPROVED"),
        "Unapproved candidates": sum(1 for row in candidates if row["review_status"] != "APPROVED"),
        "Step 2 candidate total": validation.get("candidate_total", ""),
        "Step 2 PO exact": po_summary.get("candidate_exact", ""),
        "Step 2 validation blocked": validation.get("blocked_count", ""),
        "Monthly Closing risk": risk_count,
        "Monthly Closing deferred": deferred_count,
        "Gate status": gate_status,
        "Customer DRAFT": "NOT GENERATED",
    }

    workbook = xlsxwriter.Workbook(output)
    title = workbook.add_format({"bold": True, "font_size": 14, "font_color": "#2F5496"})
    label = workbook.add_format({"bold": True, "bg_color": "#D9E2F3", "border": 1})
    value = workbook.add_format({"border": 1})
    warning = workbook.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1})
    audit_sheet = workbook.add_worksheet("_Audit")
    audit_sheet.write("A1", f"{project} - PO Usage Status Fact Audit", title)
    audit_sheet.set_column("A:A", 34)
    audit_sheet.set_column("B:B", 110)
    for index, (key, item) in enumerate(audit.items(), 2):
        audit_sheet.write(index, 0, key, label)
        audit_sheet.write(index, 1, item, warning if key in {"Gate status", "Customer DRAFT"} else value)
    audit_sheet.freeze_panes(2, 0)

    _write_table(workbook, "Approved_Facts", [
        "Fact_ID", "Person", "Month", "PO_ID", "Hours", "Days",
        "Raw_Unit", "Approval_Status", "Source_ID", "Notes",
    ], [[
        row["fact_id"], row["person_key"], row["month"], row["po_id"],
        row["hours"], round(float(row["hours"] or 0) / HOURS_PER_DAY, 4),
        row["raw_unit"], row["approval_status"], row["source_id"], row["notes"],
    ] for row in facts])
    _write_table(workbook, "Candidate_Status", [
        "Candidate_ID", "Raw_Name", "Canonical_Name", "Year", "Month",
        "PO_ID", "Hours", "Parse_Method", "Review_Status", "Preferred",
    ], [[
        row["candidate_id"], row["raw_name"], row["canonical_name"], row["year"],
        row["month"], row["po_id"], row["total_hours"], row["parse_method"],
        row["review_status"], row["is_preferred"],
    ] for row in candidates])
    _write_table(workbook, "PO_Master", [
        "PO_ID", "Parse_Status", "PO_Date", "SAP_PO_Number", "Total_Excl_VAT",
        "Delivery_Date", "Level_Quotas",
    ], [[
        row["po_id"], row["parse_status"], row["po_date"], row["sap_po_number"],
        row["total_excl_vat"], row["delivery_date"], row["level_quotas_json"],
    ] for row in po_master])
    _write_table(workbook, "Monthly_Closing", [
        "Period", "Sheet", "Row", "Raw_Name", "Canonical_Name",
        "Project_Number", "Hours", "Mapping_Status",
    ], [[
        row["period"], row["sheet_name"], row["row_number"], row["raw_name"],
        row["canonical_name"], row["project_number"], row["hours"], row["mapping_status"],
    ] for row in monthly])
    _write_table(workbook, "Risk_Summary", [
        "Person", "Period", "PDF_Hours", "Monthly_Closing_Hours", "Difference",
        "Risk", "Deferred", "Relationship",
    ], [[
        item.get("person", ""), item.get("period", ""), item.get("pdf_hours", ""),
        item.get("monthly_closing_hours", ""), item.get("difference", ""),
        item.get("risk", False), item.get("deferred", False), item.get("relationship", ""),
    ] for item in risk_rows])
    _write_table(workbook, "Open_Queue", ["Queue_Type", "Severity", "Count"], [[
        row["queue_type"], row["severity"], row["count"],
    ] for row in open_queue])
    workbook.close()
    return {"ok": True, "output": output, "audit": audit, "risk_rows": len(risk_rows)}


def main() -> int:
    parser = argparse.ArgumentParser(description="Create a COM-free PO Usage Status fact audit workbook")
    parser.add_argument("--project", default=PROJECT)
    parser.add_argument("--authority-workbook", required=True)
    parser.add_argument("--step2-json", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    result = build_workbook(args.project, args.authority_workbook, args.step2_json, args.output)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
