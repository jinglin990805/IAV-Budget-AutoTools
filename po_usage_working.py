# -*- coding: utf-8 -*-
"""Create a formula-preserving PO Usage Status working copy.

The customer-facing template is copied first and never overwritten.  Audit
tables are added as helper sheets so that formulas and layout in the approved
manual template remain intact while the user reviews the automated inputs.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

import config
from fact_store import FactStore
from project_registry import build_profile


NAVY = "2F5496"
LIGHT_BLUE = "D9E2F3"
LIGHT_YELLOW = "FFF2CC"
LIGHT_RED = "FCE4D6"
WHITE = "FFFFFF"


def _latest_audit(store: FactStore, project_id: str) -> Any:
    return store.connection.execute(
        """SELECT * FROM runs
           WHERE project_id=? AND notes LIKE 'Step 2%'
           ORDER BY started_at DESC LIMIT 1""",
        (project_id,),
    ).fetchone()


def _style_sheet(ws: Any, widths: dict[str, float] | None = None) -> None:
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    if widths:
        for column, width in widths.items():
            ws.column_dimensions[column].width = width


def _write_table(ws: Any, title: str, headers: list[str], rows: list[list[Any]], table_name: str) -> None:
    ws.delete_rows(1, ws.max_row)
    ws.append(headers)
    for row in rows:
        ws.append(row)
    header_fill = PatternFill("solid", fgColor=NAVY)
    header_font = Font(color=WHITE, bold=True)
    thin = Side(style="thin", color="D9E1F2")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=False)
    if rows:
        ref = f"A1:{openpyxl.utils.get_column_letter(len(headers))}{len(rows) + 1}"
        table = Table(displayName=table_name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showFirstColumn=False,
            showLastColumn=False, showRowStripes=True, showColumnStripes=False,
        )
        ws.add_table(table)
    ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(headers))}{max(1, len(rows) + 1)}"
    ws.title = title


def _ensure_sheet(wb: Any, name: str) -> Any:
    if name in wb.sheetnames:
        del wb[name]
    return wb.create_sheet(name)


def _audit_rows(store: FactStore, project_id: str, run_id: str | None) -> tuple[list[list[Any]], list[list[Any]], list[list[Any]], dict[str, Any]]:
    po_rows = []
    for row in store.connection.execute(
        """SELECT p.*, s.file_name FROM po_master_records p
           LEFT JOIN source_files s ON s.source_id=p.source_id
           WHERE p.project_id=? ORDER BY p.po_id, p.record_id""",
        (project_id,),
    ):
        quotas = []
        try:
            quotas = json.loads(row["level_quotas_json"] or "[]")
        except json.JSONDecodeError:
            pass
        quota_text = "; ".join(
            f"{q.get('description', '')}: {q.get('qty_days', '')}d × {q.get('unit_price', '')}"
            for q in quotas
        )
        po_rows.append([
            project_id, row["po_id"], row["parse_status"], row["po_date"],
            row["sap_po_number"], row["total_excl_vat"], row["delivery_date"],
            row["file_name"] or row["source_id"], quota_text,
            "=IF(C2=\"PARSED\",\"OK\",\"REVIEW\")",
        ])

    candidate_rows = []
    for row in store.connection.execute(
        """SELECT c.*, s.file_name FROM timesheet_candidates c
           LEFT JOIN source_files s ON s.source_id=c.source_id
           WHERE c.project_id=? AND c.review_status<>'SUPERSEDED'
           ORDER BY c.year DESC, c.month DESC, c.canonical_name, c.candidate_id""",
        (project_id,),
    ):
        candidate_rows.append([
            row["candidate_id"], row["raw_name"], row["canonical_name"], row["year"],
            row["month"], "", row["po_id"], row["total_hours"], "",
            row["parse_method"], row["review_status"], row["file_name"] or row["source_id"], "",
        ])

    monthly_rows = []
    monthly_query = """SELECT m.*, s.file_name FROM monthly_closing_facts m
                       LEFT JOIN source_files s ON s.source_id=m.source_id
                       WHERE m.project_id=?"""
    params: tuple[Any, ...] = (project_id,)
    if run_id:
        monthly_query += " AND m.run_id=?"
        params = (project_id, run_id)
    monthly_query += " ORDER BY m.period, m.sheet_name, m.row_number"
    for row in store.connection.execute(monthly_query, params):
        monthly_rows.append([
            row["period"], row["sheet_name"], row["row_number"], row["raw_name"],
            row["canonical_name"], row["project_number"], row["hours"],
            row["mapping_status"], "",
        ])

    audit = {
        "run_id": run_id or "",
        "run_status": "",
        "po_files": len(po_rows),
        "po_parsed": sum(1 for row in po_rows if row[2] == "PARSED"),
        "candidate_count": len(candidate_rows),
        "candidate_approved": sum(1 for row in candidate_rows if row[10] == "APPROVED"),
        "monthly_rows": len(monthly_rows),
        "monthly_unresolved": sum(1 for row in monthly_rows if row[7] == "UNRESOLVED_PROJECT"),
        "approved_fact_count": store.connection.execute(
            "SELECT COUNT(1) FROM timesheet_facts WHERE project_id=? AND approval_status='APPROVED' AND is_current=1",
            (project_id,),
        ).fetchone()[0],
    }
    if run_id:
        run = store.connection.execute("SELECT status FROM runs WHERE run_id=?", (run_id,)).fetchone()
        audit["run_status"] = run[0] if run else ""
    return po_rows, candidate_rows, monthly_rows, audit


def _add_formula_columns(ws: Any) -> None:
    # Formula rows are deliberately relative to the row they occupy.
    for row in range(2, ws.max_row + 1):
        ws.cell(row, 6).value = f'=IF(AND(D{row}<>"",E{row}<>""),TEXT(D{row},"0000")&"-"&TEXT(E{row},"00"),"")'
        ws.cell(row, 9).value = f'=IFERROR(H{row}/\'_Audit\'!$B$6,"")'
        ws.cell(row, 13).value = f'=IF(OR(C{row}="",G{row}="",H{row}=""),"REVIEW",IF(K{row}="APPROVED","APPROVED","PENDING"))'
        ws.cell(row, 9).number_format = "0.00"


def _build_audit_sheet(wb: Any, audit: dict[str, Any], project_id: str) -> None:
    ws = _ensure_sheet(wb, "_Audit")
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"{project_id} — PO Usage Status Working Gate"
    ws["A1"].font = Font(bold=True, size=14, color=NAVY)
    ws["A3"] = "Project"
    ws["B3"] = project_id
    ws["A4"] = "Latest audit Run ID"
    ws["B4"] = audit["run_id"]
    ws["A5"] = "Audit status"
    ws["B5"] = "=IF(OR(B10>0,B11>0,B12>0,B13>0,B14=0),\"REVIEW_REQUIRED\",\"READY_FOR_FREEZE\")"
    ws["A6"] = "Hours per day"
    ws["B6"] = config.HOURS_PER_DAY
    metrics = [
        ("PO PDF records", "=COUNTA(_PO_Master!B2:B10000)"),
        ("PO parsed records", "=COUNTIF(_PO_Master!C:C,\"PARSED\")"),
        ("PO parse review", "=COUNTIF(_PO_Master!C:C,\"<>PARSED\")"),
        ("Candidates missing PO", "=COUNTIF(_Timesheet_Candidates!G:G,\"\")"),
        ("Candidates requiring review", "=COUNTIF(_Timesheet_Candidates!M:M,\"REVIEW\")"),
        ("Monthly Closing unresolved mapping", "=COUNTIF(_Monthly_Closing!H:H,\"UNRESOLVED_PROJECT\")"),
        ("Approved timesheet facts", "=COUNTIF(_Timesheet_Candidates!K:K,\"APPROVED\")"),
    ]
    for index, (label, formula) in enumerate(metrics, 8):
        ws.cell(index, 1).value = label
        ws.cell(index, 2).value = formula
    ws["A16"] = "Gate meaning"
    ws["B16"] = "Only READY_FOR_FREEZE may feed the official PO Usage Status; REVIEW_REQUIRED is a working copy."
    ws["A18"] = "Source rule"
    ws["B18"] = "PDF is the primary timesheet source after approval; Monthly Closing is reconciliation only and is never added again."
    for cell in ["A3", "A4", "A5", "A6", "A16", "A18"]:
        ws[cell].font = Font(bold=True, color=NAVY)
        ws[cell].fill = PatternFill("solid", fgColor=LIGHT_BLUE)
    ws["B5"].fill = PatternFill("solid", fgColor=LIGHT_YELLOW)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 105
    ws.freeze_panes = "A8"


def create_working_copy(project_id: str, template_path: str, output_path: str) -> str:
    template = Path(template_path).resolve()
    output = Path(output_path).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    with FactStore() as store:
        store.initialize()
        audit_run = _latest_audit(store, project_id)
        run_id = audit_run["run_id"] if audit_run else None
        po_rows, candidate_rows, monthly_rows, audit = _audit_rows(store, project_id, run_id)

    wb = load_workbook(template, data_only=False, keep_links=True)
    original_formula_count = sum(
        1 for ws in wb.worksheets for row in ws.iter_rows()
        if row for cell in row if isinstance(cell.value, str) and cell.value.startswith("=")
    )
    wb.close()

    # openpyxl cannot round-trip this template because it contains external
    # workbook links.  Use native Excel to copy the original and append sheets,
    # preserving the original formula XML and cached values.
    payload = {
        "project": project_id,
        "audit": audit,
        "po_rows": po_rows,
        "candidate_rows": candidate_rows,
        "monthly_rows": monthly_rows,
    }
    ps_script = Path(__file__).with_name("po_usage_working_excel.ps1")
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".json", delete=False) as handle:
        json.dump(payload, handle, ensure_ascii=False)
        payload_path = handle.name
    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps_script),
             "-Template", str(template), "-Output", str(output), "-Payload", payload_path],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180,
        )
    finally:
        try:
            os.unlink(payload_path)
        except OSError:
            pass
    if result.returncode != 0:
        raise RuntimeError(f"Excel native working-copy generation failed: {result.stderr[-3000:] or result.stdout[-3000:]}")

    # Re-open once to verify that the template formulas survived the round trip.
    check = load_workbook(output, data_only=False, read_only=False, keep_links=True)
    final_formula_count = sum(
        1 for ws in check.worksheets for row in ws.iter_rows()
        if row for cell in row if isinstance(cell.value, str) and cell.value.startswith("=")
    )
    if final_formula_count < original_formula_count:
        raise RuntimeError(f"Formula preservation check failed: {original_formula_count} -> {final_formula_count}")
    data_only = load_workbook(output, data_only=True, read_only=False, keep_links=True)
    cached_formula_count = sum(
        1 for ws in check.worksheets[:2] for row in ws.iter_rows()
        if row for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
        and data_only[ws.title][cell.coordinate].value is not None
    )
    check.close()
    data_only.close()
    print(json.dumps({
        "ok": True, "output": str(output), "template": str(template),
        "original_formula_count": original_formula_count,
        "final_formula_count": final_formula_count,
        "customer_sheet_cached_values": cached_formula_count,
        "audit": audit,
    }, ensure_ascii=False))
    return str(output)


def main() -> int:
    parser = argparse.ArgumentParser(description="Create formula-preserving PO Usage Status working copy")
    parser.add_argument("--project", default="CNTS-25165")
    parser.add_argument("--template", required=True)
    parser.add_argument("--output", default="")
    args = parser.parse_args()
    output = args.output or os.path.join(
        config.BASE_DIR if hasattr(config, "BASE_DIR") else os.path.dirname(__file__),
        "output", f"PO_Usage_Status_{args.project}_Working_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    )
    create_working_copy(args.project, args.template, output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
