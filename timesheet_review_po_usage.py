# -*- coding: utf-8 -*-
"""Auto-fill historical Timesheet PO review decisions from PO Usage Status.

The Usage Status workbook is evidence for the person/month-to-PO mapping.  A
candidate still stores one ``po_id``; when the workbook shows multiple PO rows
for the same person/month, all matched POs are preserved in the review note.
Only an exact or unambiguous prefix match to the candidate PO is selected in a
multi-PO month.  Otherwise the row is left as NEEDS_SOURCE.
"""

from __future__ import annotations

import argparse
import calendar
import csv
import io
import re
import sys
from pathlib import Path
from typing import Any

import openpyxl


NAME_MAP = {
    "侯景": "Hou, Jing",
    "Hou Jing": "Hou, Jing",
    "赵士辉": "Zhao, Shihui",
    "Zhao Shihui": "Zhao, Shihui",
    "黄得佳": "Huang, Dejia",
    "黄德佳": "Huang, Dejia",
    "Huang Dejia": "Huang, Dejia",
    "刘旭": "Liu, Xu",
    "Liu Xu": "Liu, Xu",
    "轩亮亮": "Xuan, Liangliang",
    "Xuan Liangliang": "Xuan, Liangliang",
    "Yin Zicheng": "Yin, Zicheng",
    "Tang Qiwei": "Tang, Qiwei",
    "唐其伟": "Tang, Qiwei",
    "Qian Haiqing": "Qian, Haiqing",
    "钱海青": "Qian, Haiqing",
    "Cui Han": "Cui, Han",
    "崔涵": "Cui, Han",
    "Xu Chang": "Xu, Chang",
    "徐畅": "Xu, Chang",
}


def _digits(value: Any) -> str:
    return re.sub(r"\D", "", str(value or "")).lstrip("0")


def _period_header(period: str) -> str:
    year, month = period.split("-", 1)
    return f"{year}-{calendar.month_abbr[int(month)]}"


def _canonical_name(row: dict[str, str]) -> str:
    return NAME_MAP.get(row.get("canonical_name", "").strip()) or NAME_MAP.get(
        row.get("raw_name", "").strip(), ""
    )


def _load_usage_rows(workbook_path: str | Path) -> tuple[dict[str, int], list[dict[str, Any]]]:
    workbook = openpyxl.load_workbook(workbook_path, data_only=True, read_only=True)
    worksheet = workbook["PO Overview Actual_To Customer"]
    headers = {
        str(worksheet.cell(2, col).value): col
        for col in range(1, worksheet.max_column + 1)
        if worksheet.cell(2, col).value is not None
    }
    usage_rows = []
    for row_number in range(25, worksheet.max_row + 1):
        po_id = worksheet.cell(row_number, 6).value
        person = str(worksheet.cell(row_number, 3).value or "").strip()
        if po_id is None or not person:
            continue
        usage_rows.append({
            "row": row_number,
            "po_id": str(po_id).strip(),
            "person": person,
            "values": {
                header: worksheet.cell(row_number, col).value
                for header, col in headers.items()
                if header[:4].isdigit() and "-" in header
            },
        })
    return headers, usage_rows


def _matches(
    row: dict[str, str],
    usage_rows: list[dict[str, Any]],
) -> tuple[str, list[dict[str, Any]]]:
    person = _canonical_name(row)
    period = row.get("period", "").strip()
    header = _period_header(period) if re.fullmatch(r"\d{4}-\d{2}", period) else ""
    matches = []
    if not person or not header:
        return person, matches
    for usage in usage_rows:
        hours = usage["values"].get(header)
        if usage["person"] != person or not isinstance(hours, (int, float)) or hours <= 0:
            continue
        matches.append({"po_id": usage["po_id"], "hours": hours, "row": usage["row"]})
    return person, matches


def _auto_decision(row: dict[str, str], usage_rows: list[dict[str, Any]]) -> dict[str, str]:
    person, matches = _matches(row, usage_rows)
    candidate_po = _digits(row.get("candidate_po", ""))
    source = "PO Usage Status_26917.xlsx / PO Overview Actual_To Customer"
    po_summary = ", ".join(f"{item['po_id']} ({item['hours']}d)" for item in matches)

    if not person:
        return {
            "resolved_po": "",
            "review_decision": "NEEDS_SOURCE",
            "review_note": f"{source} 无法唯一匹配人员：{row.get('raw_name', '')}；需要确认姓名。",
        }
    if not matches:
        return {
            "resolved_po": "",
            "review_decision": "NEEDS_SOURCE",
            "review_note": f"{source} 未找到 {person} / {row.get('period', '')} 的有效 PO 月度数据；需要补充来源。",
        }

    exact = [item for item in matches if _digits(item["po_id"]) == candidate_po and candidate_po]
    prefix = [
        item for item in matches
        if candidate_po and (_digits(item["po_id"]).startswith(candidate_po) or candidate_po.startswith(_digits(item["po_id"])))
    ]
    selected = exact[0] if len(exact) == 1 else prefix[0] if len(prefix) == 1 else None

    if selected is None and len(matches) == 1:
        selected = matches[0]

    if selected is None:
        return {
            "resolved_po": "",
            "review_decision": "NEEDS_SOURCE",
            "review_note": f"{source} 显示同一人同一月有两张 PO：{po_summary}；候选 PO 无法唯一对应，需要补充工时单/PO 来源。",
        }

    decision = "CONFIRM_PO" if _digits(selected["po_id"]) == candidate_po else "CORRECT_PO"
    if len(matches) > 1:
        note = f"{source} 显示同一人同一月有两张 PO：{po_summary}；本候选按 PO {selected['po_id']} 复核。"
    elif decision == "CONFIRM_PO":
        note = f"{source} 匹配 {person} / {row.get('period', '')}：PO {selected['po_id']}。"
    else:
        note = f"{source} 匹配 {person} / {row.get('period', '')}：候选 PO 修正为 {selected['po_id']}。"
    return {
        "resolved_po": selected["po_id"],
        "review_decision": decision,
        "review_note": note,
    }


def auto_review_rows(
    rows: list[dict[str, str]], usage_rows: list[dict[str, Any]]
) -> list[dict[str, str]]:
    output = []
    for row in rows:
        updated = dict(row)
        updated.update(_auto_decision(row, usage_rows))
        output.append(updated)
    return output


def _read_csv(path: str | Path) -> tuple[list[str], list[dict[str, str]]]:
    with Path(path).open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fields = list(reader.fieldnames or [])
        if not fields:
            raise ValueError("CSV has no header")
        return fields, [{key: str(value or "") for key, value in row.items()} for row in reader]


def _write_csv(path: str | Path, fields: list[str], rows: list[dict[str, str]]) -> None:
    buffer = io.StringIO(newline="")
    writer = csv.DictWriter(buffer, fieldnames=fields, extrasaction="ignore", lineterminator="\n")
    writer.writeheader()
    writer.writerows(rows)
    Path(path).write_text(buffer.getvalue(), encoding="utf-8-sig", newline="")


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Fill historical PO review decisions from PO Usage Status")
    parser.add_argument("--workbook", required=True, help="PO Usage Status workbook")
    parser.add_argument("--input", required=True, help="Historical PO review CSV")
    parser.add_argument("--output", required=True, help="Decision CSV to write")
    parser.add_argument("--write", action="store_true", help="write the output CSV; default is preview")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    fields, rows = _read_csv(args.input)
    _, usage_rows = _load_usage_rows(args.workbook)
    updated = auto_review_rows(rows, usage_rows)
    counts = {}
    for row in updated:
        decision = row.get("review_decision", "")
        counts[decision] = counts.get(decision, 0) + 1
    if args.write:
        _write_csv(args.output, fields, updated)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    print({"rows": len(updated), "decisions": counts, "output": args.output, "written": args.write})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
