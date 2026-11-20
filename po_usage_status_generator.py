# -*- coding: utf-8 -*-
"""Template-preserving PO Usage Status writer.

The customer workbook contains formulas, external links, formatting and
PivotTables.  This writer patches only the needed SpreadsheetML cells in a
copy of the latest valid customer workbook.  It never rebuilds or resaves the
workbook through a library that would discard those Excel features.
"""

from __future__ import annotations

import glob
import math
import os
import posixpath
import re
import warnings
import zipfile
from datetime import datetime
from typing import Any
from xml.etree import ElementTree as ET

import config
from utils import setup_utf8_stdout


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
Q = lambda name: f"{{{MAIN_NS}}}{name}"
CUSTOMER_SHEETS = (
    "PO Overview Actual_To Customer",
    "PO Overview Plan_To Customer",
)


def _register_namespaces(xml_bytes: bytes) -> None:
    """Keep existing namespace prefixes stable when serializing edited XML."""
    for prefix, uri in re.findall(
        rb'xmlns(?::([A-Za-z_][\w.-]*))?=["\']([^"\']+)["\']', xml_bytes
    ):
        ET.register_namespace(prefix.decode("ascii") if prefix else "", uri.decode("utf-8"))


def _parse_xml(xml_bytes: bytes) -> ET.Element:
    _register_namespaces(xml_bytes)
    return ET.fromstring(xml_bytes)


def _serialize_xml(root: ET.Element, original_xml: bytes | None = None) -> bytes:
    serialized = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    if original_xml is None:
        return serialized

    # ElementTree omits namespace declarations that are only referenced by
    # mc:Ignorable (for example xr2/xr3). Excel treats those undeclared
    # prefixes as a corrupt worksheet and may discard all its cells.
    tag_pattern = re.compile(rb"<[A-Za-z_][\w:.-]*(?:\s+[^<>]*?)?>", re.DOTALL)
    source_match = tag_pattern.search(original_xml)
    output_match = tag_pattern.search(serialized)
    if source_match is None or output_match is None:
        return serialized

    declaration_pattern = re.compile(rb"\sxmlns(?::([A-Za-z_][\w.-]*))?=(['\"])(.*?)\2")
    source_declarations = {
        match.group(1) or b"": match.group(3)
        for match in declaration_pattern.finditer(source_match.group(0))
    }
    output_declarations = {
        match.group(1) or b""
        for match in declaration_pattern.finditer(output_match.group(0))
    }
    missing = [
        (prefix, uri)
        for prefix, uri in source_declarations.items()
        if prefix not in output_declarations
    ]
    if not missing:
        return serialized

    additions = b"".join(
        b" xmlns" + (b":" + prefix if prefix else b"") + b"=\"" + uri + b"\""
        for prefix, uri in missing
    )
    output_tag = output_match.group(0)
    repaired_tag = output_tag[:-1] + additions + output_tag[-1:]
    return serialized[:output_match.start()] + repaired_tag + serialized[output_match.end():]


def _excel_structure(path: str) -> dict[str, int]:
    formula_count = 0
    pivot_count = 0
    with zipfile.ZipFile(path, "r") as archive:
        names = archive.namelist()
        pivot_count = sum(
            1 for name in names
            if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)
        )
        for name in names:
            if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name):
                root = _parse_xml(archive.read(name))
                formula_count += sum(1 for node in root.iter(Q("f")))
    return {"formulas": formula_count, "pivots": pivot_count}


def _is_customer_template(path: str) -> bool:
    try:
        structure = _excel_structure(path)
    except (OSError, zipfile.BadZipFile, ET.ParseError):
        return False
    return structure["pivots"] >= 2 and structure["formulas"] >= 100


def select_customer_template(
    directory: str,
    output_path: str,
    configured_path: str | None = None,
) -> str:
    """Choose the configured customer template, then a structurally valid fallback."""
    candidates: list[tuple[float, str]] = []
    output_abs = os.path.normcase(os.path.abspath(output_path))
    if configured_path:
        configured_abs = os.path.abspath(configured_path)
        if os.path.normcase(configured_abs) == output_abs:
            raise ValueError("PO Usage Status output path cannot overwrite its source template")
        if os.path.isfile(configured_abs):
            if not _is_customer_template(configured_abs):
                raise ValueError(
                    f"Configured PO Usage Status template is missing formulas or PivotTables: {configured_abs}"
                )
            return configured_abs
        print(f"  [WARN] Configured PO Usage Status template not found: {configured_abs}; searching valid templates")

    for path in glob.glob(os.path.join(directory, "PO Usage Status_*.xlsx")):
        if os.path.normcase(os.path.abspath(path)) == output_abs:
            continue
        base_name = os.path.basename(path).casefold()
        if any(token in base_name for token in ("backup", "working", "fixed", "repaired", "preview", "~$")):
            continue
        if not _is_customer_template(path):
            continue
        candidates.append((os.path.getmtime(path), path))
    if not candidates:
        raise FileNotFoundError(
            f"No formula- and PivotTable-preserving PO Usage Status workbook found in {directory}"
        )
    return max(candidates)[1]


_select_customer_template = select_customer_template


def _finite_number(value: Any) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return 0.0
    return number if math.isfinite(number) else 0.0


def _summary_months(allocation: dict) -> list[str]:
    return sorted({
        str(month)
        for person_months in (allocation or {}).values()
        for month in (person_months or {})
    })


def _summary_people(personnel: dict, allocation: dict) -> list[dict[str, Any]]:
    people = []
    for name, info in sorted((personnel or {}).items()):
        months = {}
        for month, po_days in (allocation or {}).get(name, {}).items():
            days = sum(_finite_number(value) for value in (po_days or {}).values())
            months[str(month)] = round(days, 4) if days else 0
        people.append({
            "name": str(name),
            "supervisor": str(info.get("supervisor", "") or ""),
            "level": str(info.get("level", "") or ""),
            "months": months,
        })
    return people


def _detail_rows(personnel: dict, allocation: dict, inventory: dict) -> list[dict[str, Any]]:
    rows = []
    configured_po_levels = set()
    rate_by_level = {}
    for po_id in sorted((inventory or {}).keys(), key=str):
        for level in sorted((inventory[po_id] or {}).keys(), key=str):
            inv = inventory[po_id][level] or {}
            configured_po_levels.add((str(po_id), str(level)))
            rate_by_level.setdefault(str(level), _finite_number(inv.get("unit_price", 0)))
            for name, info in sorted((personnel or {}).items()):
                if str(info.get("level", "")) != str(level):
                    continue
                monthly = {}
                for month, po_days in (allocation or {}).get(name, {}).items():
                    days = _finite_number((po_days or {}).get(po_id, 0))
                    if days > 0:
                        monthly[str(month)] = round(days, 4)
                if not monthly:
                    continue
                rows.append({
                    "name": str(name),
                    "supervisor": str(info.get("supervisor", "") or ""),
                    "level": str(level),
                    "po_id": str(po_id),
                    "unit_price": _finite_number(inv.get("unit_price", 0)),
                    "quota": _finite_number(inv.get("quota", 0)),
                    "remaining": _finite_number(inv.get("remaining", 0)),
                    "start_month": str(inv.get("start_month", "") or ""),
                    "monthly": monthly,
                })

    # Approved Review Center facts can reference a PO that has not yet been
    # added to the quota configuration. Keep its actual usage visible instead
    # of silently dropping it; leave quota/remaining blank until configured.
    for name, info in sorted((personnel or {}).items()):
        level = str(info.get("level", ""))
        po_ids = {
            str(po_id)
            for po_days in (allocation or {}).get(name, {}).values()
            for po_id, days in (po_days or {}).items()
            if _finite_number(days) > 0
            and (str(po_id), level) not in configured_po_levels
        }
        for po_id in sorted(po_ids):
            monthly = {
                str(month): round(_finite_number((po_days or {}).get(po_id, 0)), 4)
                for month, po_days in (allocation or {}).get(name, {}).items()
                if _finite_number((po_days or {}).get(po_id, 0)) > 0
            }
            if not monthly:
                continue
            rows.append({
                "name": str(name),
                "supervisor": str(info.get("supervisor", "") or ""),
                "level": level,
                "po_id": po_id,
                "unit_price": rate_by_level.get(level),
                "quota": None,
                "remaining": None,
                "start_month": "",
                "monthly": monthly,
                "quota_configured": False,
            })
    return rows


def _payload(
    personnel: dict,
    merged_allocation: dict,
    inventory: dict,
    person_forecast: dict,
    as_of_month: str,
) -> dict[str, Any]:
    actual_allocation = merged_allocation or {}
    plan_allocation = person_forecast or {}
    return {
        "as_of_month": as_of_month,
        "actual": {
            "months": _summary_months(actual_allocation),
            "people": _summary_people(personnel, actual_allocation),
            "details": _detail_rows(personnel, actual_allocation, inventory or {}),
        },
        "plan": {
            "months": _summary_months(plan_allocation),
            "people": _summary_people(personnel, plan_allocation),
            "details": _detail_rows(personnel, plan_allocation, inventory or {}),
        },
    }


def _normalize(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").casefold())


def _month_key(value: Any) -> str | None:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m")
    if value is None:
        return None
    text = str(value).strip()
    match = re.fullmatch(r"(\d{4})[-/](\d{1,2})", text)
    if match and 1 <= int(match.group(2)) <= 12:
        return f"{int(match.group(1)):04d}-{int(match.group(2)):02d}"
    match = re.fullmatch(r"(\d{4})[-/]([A-Za-z]{3,9})", text)
    if match:
        month = {
            "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
            "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
        }.get(match.group(2)[:3].casefold())
        if month:
            return f"{int(match.group(1)):04d}-{month:02d}"
    return None


def _month_label(month: str) -> str:
    return datetime.strptime(month, "%Y-%m").strftime("%Y-%b")


def _col_index(cell_ref: str) -> int:
    result = 0
    for char in re.match(r"[A-Z]+", cell_ref.upper()).group(0):
        result = result * 26 + ord(char) - 64
    return result


def _col_letter(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _detail_key(name: Any, level: Any, po_id: Any) -> tuple[str, str, str]:
    return _normalize(name), _normalize(level), _normalize(po_id)


def _sheet_parts(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook = _parse_xml(archive.read("xl/workbook.xml"))
    rels = _parse_xml(archive.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    result = {}
    for sheet in workbook.findall(f".//{Q('sheet')}"):
        rel_id = sheet.attrib.get(f"{{{REL_NS}}}id")
        target = rel_targets[rel_id]
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        result[sheet.attrib["name"]] = part
    return result


class _SheetXml:
    def __init__(self, xml_bytes: bytes):
        self.original_xml = xml_bytes
        self.root = _parse_xml(xml_bytes)
        self.sheet_data = self.root.find(Q("sheetData"))
        if self.sheet_data is None:
            raise ValueError("Worksheet XML has no sheetData")
        self.rows = {
            int(row.attrib["r"]): row
            for row in self.sheet_data.findall(Q("row"))
        }
        self.cells_written = 0
        self.formula_cells_preserved = 0
        self.formulas_changed = 0
        self.rows_added = 0

    def _row(self, row_num: int) -> ET.Element:
        row = self.rows.get(row_num)
        if row is not None:
            return row
        row = ET.Element(Q("row"), {"r": str(row_num)})
        following = next((item for item in self.sheet_data.findall(Q("row"))
                          if int(item.attrib["r"]) > row_num), None)
        if following is None:
            self.sheet_data.append(row)
        else:
            self.sheet_data.insert(list(self.sheet_data).index(following), row)
        self.rows[row_num] = row
        return row

    def cell(self, row_num: int, col_num: int, style_id: int | None = None) -> ET.Element:
        row = self._row(row_num)
        ref = f"{_col_letter(col_num)}{row_num}"
        for cell in row.findall(Q("c")):
            if cell.attrib.get("r") == ref:
                if style_id is not None:
                    cell.set("s", str(style_id))
                return cell
        attrs = {"r": ref}
        if style_id is not None:
            attrs["s"] = str(style_id)
        new_cell = ET.Element(Q("c"), attrs)
        following = next((item for item in row.findall(Q("c"))
                          if _col_index(item.attrib["r"]) > col_num), None)
        if following is None:
            row.append(new_cell)
        else:
            row.insert(list(row).index(following), new_cell)
        return new_cell

    def style_id(self, row_num: int, col_num: int) -> int | None:
        """Return a style reference already present in the source worksheet XML.

        Do not synthesize cells here: OpenPyXL can assign an in-memory style ID
        for a blank cell that has no corresponding entry in styles.xml.  Such
        an ID is invalid when copied into the untouched OOXML package.
        """
        row = self.rows.get(row_num)
        if row is None:
            return None
        ref = f"{_col_letter(col_num)}{row_num}"
        cell = next((item for item in row.findall(Q("c"))
                     if item.attrib.get("r") == ref), None)
        if cell is None:
            return None
        value = cell.attrib.get("s")
        return int(value) if value is not None else None

    def set_column_width(self, col_num: int, width: float | None) -> None:
        if width is None:
            return
        cols = self.root.find(Q("cols"))
        if cols is None:
            cols = ET.Element(Q("cols"))
            self.root.insert(list(self.root).index(self.sheet_data), cols)
        existing = next((
            col for col in cols.findall(Q("col"))
            if int(col.attrib.get("min", "0")) == col_num
            and int(col.attrib.get("max", "0")) == col_num
        ), None)
        if existing is None:
            existing = ET.Element(Q("col"), {"min": str(col_num), "max": str(col_num)})
            following = next((
                col for col in cols.findall(Q("col"))
                if int(col.attrib.get("min", "0")) > col_num
            ), None)
            if following is None:
                cols.append(existing)
            else:
                cols.insert(list(cols).index(following), existing)
        existing.set("width", format(float(width), ".12g"))
        existing.set("customWidth", "1")

    def reset_view(self) -> None:
        """Open the worksheet at its populated top-left area."""
        views = self.root.find(Q("sheetViews"))
        if views is None:
            return
        for view in views.findall(Q("sheetView")):
            view.set("topLeftCell", "A1")
            selections = view.findall(Q("selection"))
            if not selections:
                selections = [ET.SubElement(view, Q("selection"))]
            for selection in selections:
                selection.set("activeCell", "A1")
                selection.set("sqref", "A1")

    def write_value(
        self,
        row_num: int,
        col_num: int,
        value: Any,
        style_id: int | None = None,
    ) -> bool:
        cell = self.cell(row_num, col_num, style_id)
        if cell.find(Q("f")) is not None:
            self.formula_cells_preserved += 1
            return False
        for child in list(cell):
            if child.tag in (Q("v"), Q("is"), Q("f")):
                cell.remove(child)
        cell.attrib.pop("t", None)
        if value is None or value == "":
            self.cells_written += 1
            return True
        if isinstance(value, bool):
            cell.set("t", "b")
            ET.SubElement(cell, Q("v")).text = "1" if value else "0"
        elif isinstance(value, (int, float)):
            number = float(value)
            if not math.isfinite(number):
                number = 0.0
            cell.set("t", "n")
            ET.SubElement(cell, Q("v")).text = format(number, ".15g")
        else:
            cell.set("t", "inlineStr")
            inline = ET.SubElement(cell, Q("is"))
            text = ET.SubElement(inline, Q("t"))
            text.text = str(value)
            if text.text != text.text.strip():
                text.set(f"{{{XML_NS}}}space", "preserve")
        self.cells_written += 1
        return True

    def write_formula(
        self,
        row_num: int,
        col_num: int,
        formula: str,
        style_id: int | None = None,
    ) -> None:
        cell = self.cell(row_num, col_num, style_id)
        for child in list(cell):
            if child.tag in (Q("v"), Q("is"), Q("f")):
                cell.remove(child)
        cell.attrib.pop("t", None)
        ET.SubElement(cell, Q("f")).text = formula.lstrip("=")
        ET.SubElement(cell, Q("v"))
        self.formulas_changed += 1

    def clone_detail_row(
        self,
        source_row: int,
        target_row: int,
        formula_cols: set[int],
        last_col: int,
    ) -> None:
        source = self.rows.get(source_row)
        if source is None:
            raise ValueError(f"Cannot clone missing style row {source_row}")
        new_row = ET.Element(source.tag, dict(source.attrib))
        new_row.set("r", str(target_row))
        for source_cell in source.findall(Q("c")):
            ref = source_cell.attrib.get("r", "")
            col = _col_index(ref)
            if col > last_col:
                continue
            target_ref = f"{_col_letter(col)}{target_row}"
            attrs = {k: v for k, v in source_cell.attrib.items() if k in ("s", "cm", "vm", "ph")}
            attrs["r"] = target_ref
            target_cell = ET.Element(Q("c"), attrs)
            source_formula = source_cell.find(Q("f"))
            if source_formula is not None and col in formula_cols and source_formula.text:
                from openpyxl.formula.translate import Translator
                translated = Translator(
                    "=" + source_formula.text, origin=ref
                ).translate_formula(target_ref)
                formula = ET.SubElement(target_cell, Q("f"))
                formula.text = translated.lstrip("=")
            new_row.append(target_cell)
        following = next((item for item in self.sheet_data.findall(Q("row"))
                          if int(item.attrib["r"]) > target_row), None)
        if following is None:
            self.sheet_data.append(new_row)
        else:
            self.sheet_data.insert(list(self.sheet_data).index(following), new_row)
        self.rows[target_row] = new_row
        self.rows_added += 1

    def update_dimension(self) -> None:
        max_row = max(self.rows, default=1)
        max_col = 1
        for row in self.rows.values():
            for cell in row.findall(Q("c")):
                max_col = max(max_col, _col_index(cell.attrib["r"]))
        dimension = self.root.find(Q("dimension"))
        if dimension is None:
            dimension = ET.Element(Q("dimension"))
            self.root.insert(0, dimension)
        dimension.set("ref", f"A1:{_col_letter(max_col)}{max_row}")

    def to_bytes(self) -> bytes:
        self.update_dimension()
        return _serialize_xml(self.root, self.original_xml)


def _template_layout(ws: Any) -> dict[str, Any]:
    summary_rows = {}
    summary_end = 2
    for row in range(3, ws.max_row + 1):
        if str(ws.cell(row, 1).value or "").strip() == "PO CBD":
            break
        name = ws.cell(row, 3).value
        if name:
            summary_rows[_normalize(name)] = row
            summary_end = max(summary_end, row)

    def find_header(row: int, value: str) -> int:
        for col in range(1, ws.max_column + 1):
            if str(ws.cell(row, col).value or "").strip() == value:
                return col
        return 0

    summary_total_col = find_header(2, "Total mandays")
    detail_total_col = find_header(24, "Total GR mandays")
    remaining_col = find_header(24, "PO Remaining")
    amount_col = find_header(24, "GR Amount")

    # PivotTables to the right of the customer tables also contain date cells.
    # Do not treat those as customer month columns or use their unformatted
    # header style for a missing month in the main Actual/Plan table.
    month_columns = {}
    summary_month_columns = {}
    detail_month_columns = {}
    for header_row, end_col, month_map in (
        (2, summary_total_col, summary_month_columns),
        (24, detail_total_col, detail_month_columns),
    ):
        for col in range(1, end_col):
            month = _month_key(ws.cell(header_row, col).value)
            if month:
                month_columns.setdefault(month, col)
                month_map[month] = col

    detail_end = 24
    details: dict[tuple[str, str, str], list[int]] = {}
    for row in range(25, ws.max_row + 1):
        name = ws.cell(row, 3).value
        level = ws.cell(row, 4).value
        po_id = ws.cell(row, 6).value
        if name and level and po_id:
            details.setdefault(_detail_key(name, level, po_id), []).append(row)
            detail_end = max(detail_end, row)

    if not summary_rows or not month_columns or not summary_total_col or not detail_total_col:
        raise ValueError(f"Unrecognized customer workbook layout: {ws.title}")

    summary_month_style_col = max(summary_month_columns.values()) if summary_month_columns else 11
    detail_month_style_col = max(detail_month_columns.values()) if detail_month_columns else 11
    formula_source_row = 25
    formula_cols = {7, 10, detail_total_col}
    if remaining_col:
        formula_cols.add(remaining_col)
    best_score = -1
    for row in range(25, detail_end + 1):
        score = sum(1 for col in formula_cols if ws.cell(row, col).data_type == "f")
        if score > best_score:
            best_score = score
            formula_source_row = row

    return {
        "summary_rows": summary_rows,
        "summary_end": summary_end,
        "month_columns": month_columns,
        "summary_month_style_col": summary_month_style_col,
        "detail_month_style_col": detail_month_style_col,
        "summary_total_col": summary_total_col,
        "detail_total_col": detail_total_col,
        "remaining_col": remaining_col,
        "detail_style_end_col": max(detail_total_col, remaining_col, amount_col),
        "detail_rows": details,
        "detail_end": detail_end,
        "formula_cols": formula_cols,
        "formula_source_row": formula_source_row,
        "max_column": ws.max_column,
    }


def _person_month_lookup(dataset: dict) -> dict[str, dict[str, float]]:
    result = {}
    for person in dataset.get("people", []):
        name = _normalize(person.get("name"))
        result[name] = {
            str(month): _finite_number(days)
            for month, days in (person.get("months") or {}).items()
        }
    return result


def _record_month_lookup(dataset: dict) -> dict[tuple[str, str, str], dict[str, float]]:
    return {
        _detail_key(record.get("name"), record.get("level"), record.get("po_id")): {
            str(month): _finite_number(days)
            for month, days in (record.get("monthly") or {}).items()
        }
        for record in dataset.get("details", [])
    }


def _excel_date_serial(month: str) -> float | None:
    if not re.fullmatch(r"\d{4}-\d{2}", month):
        return None
    value = datetime.strptime(month, "%Y-%m")
    return float((value - datetime(1899, 12, 30)).days)


def _write_dataset(
    ws: Any,
    xml: _SheetXml,
    dataset: dict,
    layout: dict[str, Any],
    *,
    allow_plan_extension: bool,
    cached_ws: Any = None,
) -> tuple[int, dict[str, int]]:
    stats = {
        "cells_written": 0,
        "formula_cells_preserved": 0,
        "new_detail_rows": 0,
        "unconfigured_po_rows": 0,
    }
    month_columns = dict(layout["month_columns"])
    last_existing_month_col = max(month_columns.values()) if month_columns else 11
    months = sorted({str(month) for month in dataset.get("months", [])})
    missing = [month for month in months if month not in month_columns]
    if missing and not allow_plan_extension:
        raise ValueError(
            "Actual month columns are missing from the source customer workbook: "
            + ", ".join(missing)
        )

    extension_start = None
    extension_total_col = None
    extension_remaining_col = None
    if missing:
        extension_start = layout["max_column"] + 2  # leave one blank separator column
        extension_total_col = extension_start + len(missing)
        extension_remaining_col = extension_total_col + 1
        xml.write_value(1, extension_start, "Plan Forecast Extension", xml.style_id(1, 1))
        summary_style_col = layout["summary_month_style_col"]
        detail_style_col = layout["detail_month_style_col"]
        summary_month_header_style = xml.style_id(2, summary_style_col)
        detail_month_header_style = xml.style_id(24, detail_style_col)
        summary_total_header_style = xml.style_id(2, layout["summary_total_col"])
        detail_total_header_style = xml.style_id(24, layout["detail_total_col"])
        remaining_header_style = (
            xml.style_id(24, layout["remaining_col"])
            if layout["remaining_col"] else detail_total_header_style
        )
        for offset, month in enumerate(missing):
            col = extension_start + offset
            month_columns[month] = col
            xml.write_value(2, col, _month_label(month), summary_month_header_style)
            xml.write_value(24, col, _month_label(month), detail_month_header_style)
            summary_width = ws.column_dimensions[_col_letter(layout["summary_month_style_col"])].width
            xml.set_column_width(col, summary_width)
        xml.write_value(2, extension_total_col, "Forecast Extension Total", summary_total_header_style)
        xml.write_value(24, extension_total_col, "Forecast Extension Total", detail_total_header_style)
        xml.write_value(24, extension_remaining_col, "Remaining After Forecast", remaining_header_style)
        xml.set_column_width(
            extension_total_col,
            ws.column_dimensions[_col_letter(layout["summary_total_col"])].width,
        )
        if layout["remaining_col"]:
            xml.set_column_width(
                extension_remaining_col,
                ws.column_dimensions[_col_letter(layout["remaining_col"])].width,
            )

    # Ensure months that exist only in the detail header (September 2026 in the
    # current customer file) also have a matching summary header.
    for month in months:
        col = month_columns[month]
        if not _month_key(ws.cell(2, col).value):
            xml.write_value(
                2, col, _month_label(month),
                xml.style_id(2, layout["summary_month_style_col"]),
            )

    summary_rows: dict[str, int] = layout["summary_rows"]
    for person in dataset.get("people", []):
        if _normalize(person.get("name")) not in summary_rows:
            raise ValueError(f"Person row is missing in customer template: {person.get('name')}")

    # Clear and refill only non-formula input cells for the months in this run.
    for month in months:
        col = month_columns[month]
        for row in range(3, layout["summary_end"] + 1):
            xml.write_value(row, col, None)
        for row in range(25, layout["detail_end"] + 1):
            xml.write_value(row, col, None)

    people_months = _person_month_lookup(dataset)
    for key, row in summary_rows.items():
        month_values = people_months.get(key, {})
        for month in months:
            value = month_values.get(month, 0.0)
            xml.write_value(row, month_columns[month], value if value else None)

    # The Plan template can carry a customer-approved monthly PO allocation
    # even when the current person forecast does not include that month. In
    # that case, keep the detail plan as the source of truth and roll its
    # per-PO values up into the person summary instead of leaving the summary
    # column blank (September 2026 is the current example).
    if allow_plan_extension:
        for month, col in month_columns.items():
            if month in months:
                continue
            person_detail_rows: dict[str, list[int]] = {}
            for detail_row in range(25, layout["detail_end"] + 1):
                person_key = _normalize(ws.cell(detail_row, 3).value)
                if person_key and ws.cell(detail_row, col).value not in (None, ""):
                    person_detail_rows.setdefault(person_key, []).append(detail_row)
            for name_key, summary_row in summary_rows.items():
                rows_for_person = person_detail_rows.get(name_key, [])
                if not rows_for_person:
                    continue
                values = [
                    cached_ws.cell(detail_row, col).value
                    if cached_ws is not None else ws.cell(detail_row, col).value
                    for detail_row in rows_for_person
                ]
                has_uncached_formula = any(
                    ws.cell(detail_row, col).data_type == "f"
                    and not isinstance(value, (int, float))
                    for detail_row, value in zip(rows_for_person, values)
                )
                if has_uncached_formula:
                    month_letter = _col_letter(col)
                    xml.write_formula(
                        summary_row,
                        col,
                        f"=SUMIF($C$25:$C${layout['detail_end']},C{summary_row},"
                        f"${month_letter}$25:${month_letter}${layout['detail_end']})",
                    )
                else:
                    total = sum(
                        _finite_number(value)
                        for value in values
                        if isinstance(value, (int, float))
                    )
                    xml.write_value(summary_row, col, total if total else None)

    detail_rows = {key: list(rows) for key, rows in layout["detail_rows"].items()}
    used_rows: set[int] = set()
    records_by_row: list[tuple[int, dict[str, Any], bool]] = []
    detail_end = layout["detail_end"]
    for record in dataset.get("details", []):
        key = _detail_key(record.get("name"), record.get("level"), record.get("po_id"))
        row = next((candidate for candidate in detail_rows.get(key, []) if candidate not in used_rows), None)
        is_new = row is None
        if is_new:
            row = detail_end + 1
            xml.clone_detail_row(
                layout["formula_source_row"], row, layout["formula_cols"] - {10},
                layout["detail_style_end_col"],
            )
            detail_end = row
            detail_rows.setdefault(key, []).append(row)
            stats["new_detail_rows"] += 1
        used_rows.add(row)
        records_by_row.append((row, record, is_new))

    display_names = {
        key: str(ws.cell(row, 3).value or "")
        for key, row in summary_rows.items()
    }
    for row, record, _is_new in records_by_row:
        name = str(record.get("name", ""))
        level = str(record.get("level", ""))
        po_id = str(record.get("po_id", ""))
        preserve_template_metadata = record.get("quota_configured") is False and not _is_new
        if not preserve_template_metadata:
            xml.write_value(row, 1, record.get("supervisor", ""))
            xml.write_value(row, 2, "IAV")
            xml.write_value(row, 3, display_names.get(_normalize(name), name))
            xml.write_value(row, 4, level)
            unit_price = record.get("unit_price")
            xml.write_value(row, 5, _finite_number(unit_price) if unit_price is not None else None)
            po_numeric: Any = float(po_id) if po_id.isdigit() else po_id
            xml.write_value(row, 6, po_numeric)
            date_serial = _excel_date_serial(str(record.get("start_month", "")))
            if date_serial is not None:
                xml.write_value(row, 7, date_serial)
            quota = record.get("quota")
            xml.write_value(row, 8, _finite_number(quota) if quota is not None else None)
            xml.write_value(row, 9, 0)
            remaining = record.get("remaining")
            xml.write_value(row, 10, _finite_number(remaining) if remaining is not None else None)
        if record.get("quota_configured") is False:
            template_quota = ws.cell(row, 8).value if not _is_new else None
            if template_quota in (None, "") or _is_new:
                stats["unconfigured_po_rows"] += 1
        month_values = record.get("monthly") or {}
        for month in months:
            days = _finite_number(month_values.get(month, 0))
            xml.write_value(row, month_columns[month], days if days else None)
        if _is_new:
            xml.write_formula(
                row,
                layout["detail_total_col"],
                f"=SUM(K{row}:{_col_letter(last_existing_month_col)}{row})",
            )

    # Give appended detail rows the same existing borders, then add their input
    # styles in any forecast extension columns.
    if missing:
        summary_style_col = layout["summary_month_style_col"]
        detail_style_col = layout["detail_month_style_col"]
        summary_total_style_col = layout["summary_total_col"]
        detail_total_style_col = layout["detail_total_col"]
        remaining_style_col = layout["remaining_col"] or layout["detail_total_col"]
        for row in range(3, layout["summary_end"] + 1):
            style_id = xml.style_id(row, summary_style_col)
            total_style = xml.style_id(row, summary_total_style_col)
            for offset, month in enumerate(missing):
                xml.cell(row, extension_start + offset, style_id)
            xml.write_formula(
                row,
                extension_total_col,
                f"=SUM({ _col_letter(extension_start)}{row}:{_col_letter(extension_total_col - 1)}{row})",
                total_style,
            )

        last_style_row = layout["detail_end"]
        for row in range(25, detail_end + 1):
            style_row = row if row <= layout["detail_end"] else last_style_row
            month_style = xml.style_id(style_row, detail_style_col)
            total_style = xml.style_id(style_row, detail_total_style_col)
            remaining_style = xml.style_id(style_row, remaining_style_col)
            for offset, month in enumerate(missing):
                xml.cell(row, extension_start + offset, month_style)
            xml.write_formula(
                row,
                extension_total_col,
                f"=SUM({_col_letter(extension_start)}{row}:{_col_letter(extension_total_col - 1)}{row})",
                total_style,
            )
            if layout["remaining_col"]:
                xml.write_formula(
                    row,
                    extension_remaining_col,
                    f"=J{row}-{_col_letter(layout['detail_total_col'])}{row}-{_col_letter(extension_total_col)}{row}",
                    remaining_style,
                )

    # The Actual summary/detail total formulas need to include the newly used
    # last actual month. Plan's existing totals remain intact; its extension has
    # separate totals to the right of the original PivotTable.
    if dataset is not None and not allow_plan_extension:
        last_month_col = max(month_columns[month] for month in months) if months else 0
        end_letter = _col_letter(last_month_col) if last_month_col else "K"
        total_col = layout["summary_total_col"]
        for row in range(3, layout["summary_end"] + 1):
            if ws.cell(row, total_col).data_type == "f":
                xml.write_formula(row, total_col, f"=SUM(K{row}:{end_letter}{row})")
        for row in range(25, detail_end + 1):
            if ws.cell(row, layout["detail_total_col"]).data_type == "f" or row > layout["detail_end"]:
                xml.write_formula(
                    row,
                    layout["detail_total_col"],
                    f"=SUM(K{row}:{end_letter}{row})",
                )

    # Recompute write counters from the live XML writer.
    stats["cells_written"] = xml.cells_written
    stats["formula_cells_preserved"] = xml.formula_cells_preserved
    return detail_end, stats


def _sheet_formula_refs(archive: zipfile.ZipFile) -> dict[str, set[str]]:
    parts = _sheet_parts(archive)
    result = {}
    for sheet_name, part in parts.items():
        root = _parse_xml(archive.read(part))
        refs = set()
        for cell in root.iter(Q("c")):
            if cell.find(Q("f")) is not None:
                refs.add(cell.attrib.get("r", ""))
        result[sheet_name] = refs
    return result


def _workbook_sheet_names(archive: zipfile.ZipFile) -> list[str]:
    root = _parse_xml(archive.read("xl/workbook.xml"))
    return [sheet.attrib["name"] for sheet in root.findall(f".//{Q('sheet')}")]


def _patch_workbook_calc(xml_bytes: bytes) -> bytes:
    root = _parse_xml(xml_bytes)
    calc = root.find(Q("calcPr"))
    if calc is None:
        calc = ET.Element(Q("calcPr"))
        insert_before = {"oleSize", "customWorkbookViews", "pivotCaches", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"}
        children = list(root)
        index = next((i for i, child in enumerate(children) if child.tag.split("}")[-1] in insert_before), len(children))
        root.insert(index, calc)
    calc.set("calcMode", "auto")
    calc.set("fullCalcOnLoad", "1")
    calc.set("forceFullCalc", "1")
    return _serialize_xml(root, xml_bytes)


def _patch_pivot_cache(xml_bytes: bytes, actual_sheet: str, last_row: int, end_col: int) -> bytes:
    root = _parse_xml(xml_bytes)
    root.set("refreshOnLoad", "1")
    root.set("enableRefresh", "1")
    source = root.find(f".//{Q('worksheetSource')}")
    if source is not None and source.attrib.get("sheet") == actual_sheet:
        source.set("ref", f"A24:{_col_letter(end_col)}{last_row}")
    return _serialize_xml(root, xml_bytes)


def _validate_output(source: str, output: str) -> dict[str, int]:
    with zipfile.ZipFile(source, "r") as before, zipfile.ZipFile(output, "r") as after:
        before_names = set(before.namelist())
        after_names = set(after.namelist())
        if before_names != after_names:
            raise RuntimeError("Workbook package parts changed while updating PO Usage Status")
        if before.read("xl/styles.xml") != after.read("xl/styles.xml"):
            raise RuntimeError("Excel cell styles changed while updating PO Usage Status")
        styles_root = _parse_xml(after.read("xl/styles.xml"))
        cell_xfs = styles_root.find(Q("cellXfs"))
        style_count = len(cell_xfs) if cell_xfs is not None else 0
        for part in sorted(name for name in after_names
                           if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)):
            sheet_xml = after.read(part)
            sheet_root = _parse_xml(sheet_xml)
            root_tag = re.search(rb"<worksheet\b[^>]*>", sheet_xml, re.DOTALL)
            namespace_prefixes = set()
            if root_tag is not None:
                namespace_prefixes = {
                    (match.group(1) or b"").decode("ascii")
                    for match in re.finditer(
                        rb"\sxmlns(?::([A-Za-z_][\w.-]*))?=(['\"])(.*?)\2",
                        root_tag.group(0),
                    )
                }
            ignorable = sheet_root.attrib.get(
                "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable", ""
            ).split()
            undefined = [prefix for prefix in ignorable if prefix not in namespace_prefixes]
            if undefined:
                raise RuntimeError(
                    f"Undefined mc:Ignorable namespace prefixes in {part}: {', '.join(undefined)}"
                )
            sheet_views = sheet_root.find(Q("sheetViews"))
            if sheet_views is not None:
                for view in sheet_views.findall(Q("sheetView")):
                    selections = view.findall(Q("selection"))
                    if view.attrib.get("topLeftCell", "A1") != "A1" or any(
                        selection.attrib.get("activeCell") != "A1"
                        or selection.attrib.get("sqref") != "A1"
                        for selection in selections
                    ):
                        raise RuntimeError(
                            f"Worksheet view does not open at A1 in {part}"
                        )
            invalid = [
                (cell.attrib.get("r", "?"), cell.attrib["s"])
                for cell in sheet_root.iter(Q("c"))
                if "s" in cell.attrib and int(cell.attrib["s"]) >= style_count
            ]
            if invalid:
                sample = ", ".join(f"{ref}:s={style}" for ref, style in invalid[:8])
                raise RuntimeError(
                    f"Invalid style references in {part} (cellXfs={style_count}): {sample}"
                )
        if _workbook_sheet_names(before) != _workbook_sheet_names(after):
            raise RuntimeError("Worksheet names changed while updating PO Usage Status")
        pivot_parts = sorted(name for name in before_names if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name))
        if not pivot_parts or any(before.read(name) != after.read(name) for name in pivot_parts):
            raise RuntimeError("PivotTable definitions changed while updating PO Usage Status")
        external_parts = sorted(name for name in before_names if name.startswith("xl/externalLinks/"))
        if any(before.read(name) != after.read(name) for name in external_parts):
            raise RuntimeError("External workbook links changed while updating PO Usage Status")
        formulas_before = _sheet_formula_refs(before)
        formulas_after = _sheet_formula_refs(after)
        for sheet_name, refs in formulas_before.items():
            missing = refs - formulas_after.get(sheet_name, set())
            if missing:
                sample = ", ".join(sorted(missing)[:10])
                raise RuntimeError(f"Original formulas were removed from {sheet_name}: {sample}")
        structure = _excel_structure(output)
        if structure["pivots"] != len(pivot_parts):
            raise RuntimeError("PivotTable count changed while updating PO Usage Status")
        return structure


def _write_template_copy(template_path: str, output_path: str, data: dict) -> dict[str, Any]:
    from openpyxl import load_workbook

    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        workbook = load_workbook(template_path, data_only=False, read_only=False, keep_links=True)
        cached_workbook = load_workbook(template_path, data_only=True, read_only=False, keep_links=True)
    try:
        missing_sheets = [name for name in CUSTOMER_SHEETS if name not in workbook.sheetnames]
        if missing_sheets:
            raise ValueError(f"Customer template is missing sheets: {', '.join(missing_sheets)}")
        layouts = {name: _template_layout(workbook[name]) for name in CUSTOMER_SHEETS}

        with zipfile.ZipFile(template_path, "r") as source_zip:
            parts = _sheet_parts(source_zip)
            changed: dict[str, bytes] = {}
            actual_last_row = layouts[CUSTOMER_SHEETS[0]]["detail_end"]
            actual_end_col = layouts[CUSTOMER_SHEETS[0]]["detail_total_col"]
            summaries = {}

            for sheet_name, dataset, allow_extension in (
                (CUSTOMER_SHEETS[0], data["actual"], False),
                (CUSTOMER_SHEETS[1], data["plan"], True),
            ):
                part = parts[sheet_name]
                xml = _SheetXml(source_zip.read(part))
                last_row, sheet_stats = _write_dataset(
                    workbook[sheet_name], xml, dataset, layouts[sheet_name],
                    allow_plan_extension=allow_extension,
                    cached_ws=cached_workbook[sheet_name],
                )
                xml.reset_view()
                changed[part] = xml.to_bytes()
                summaries[sheet_name] = sheet_stats
                if sheet_name == CUSTOMER_SHEETS[0]:
                    actual_last_row = last_row
                    actual_end_col = layouts[sheet_name]["detail_total_col"]

            changed["xl/workbook.xml"] = _patch_workbook_calc(source_zip.read("xl/workbook.xml"))
            for name in source_zip.namelist():
                if re.fullmatch(r"xl/pivotCache/pivotCacheDefinition\d+\.xml", name):
                    changed[name] = _patch_pivot_cache(
                        source_zip.read(name),
                        CUSTOMER_SHEETS[0],
                        actual_last_row,
                        actual_end_col,
                    )

            with zipfile.ZipFile(output_path, "w") as output_zip:
                for info in source_zip.infolist():
                    payload = changed.get(info.filename)
                    if payload is None:
                        payload = source_zip.read(info.filename)
                    output_zip.writestr(info, payload)
    finally:
        workbook.close()
        cached_workbook.close()
    return summaries


def generate_po_usage_status(
    budget_path: str,
    output_path: str,
    personnel: dict,
    merged_allocation: dict = None,
    inventory: dict = None,
    forward_projection: dict = None,
    person_forecast: dict = None,
    as_of_month: str = None,
    template_path: str | None = None,
) -> str:
    """Create a new customer workbook while preserving its Excel structure."""
    del forward_projection  # person_forecast contains the month-by-month projection.
    as_of_month = as_of_month or datetime.now().strftime("%Y-%m")
    budget_dir = os.path.dirname(os.path.abspath(budget_path))
    template_path = select_customer_template(budget_dir, output_path, template_path)
    if os.path.normcase(os.path.abspath(template_path)) == os.path.normcase(os.path.abspath(output_path)):
        raise ValueError("Output path must not overwrite the source PO Usage Status workbook")

    output_parent = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(output_parent, exist_ok=True)
    if os.path.exists(output_path):
        stem, extension = os.path.splitext(output_path)
        suffix = 2
        while os.path.exists(f"{stem}_{suffix}{extension}"):
            suffix += 1
        output_path = f"{stem}_{suffix}{extension}"

    payload = _payload(
        personnel=personnel,
        merged_allocation=merged_allocation or {},
        inventory=inventory or {},
        person_forecast=person_forecast or {},
        as_of_month=as_of_month,
    )
    try:
        print(f"  PO Usage Status template: {template_path}")
        stats = _write_template_copy(template_path, os.path.abspath(output_path), payload)
        structure = _validate_output(template_path, output_path)
        for sheet_name in CUSTOMER_SHEETS:
            sheet_stats = stats[sheet_name]
            print(
                f"  {sheet_name}: {sheet_stats['cells_written']} cells updated, "
                f"{sheet_stats['formula_cells_preserved']} formula cells kept, "
                f"{sheet_stats['new_detail_rows']} detail rows added"
            )
            if sheet_stats["unconfigured_po_rows"]:
                print(
                    f"  [WARN] {sheet_name}: {sheet_stats['unconfigured_po_rows']} detail rows "
                    "use approved facts for PO(s) without quota configuration; usage is shown, "
                    "quota and remaining are blank"
                )
        print(
            "  Template structure preserved: "
            f"{structure['formulas']} formulas, {structure['pivots']} PivotTables"
        )
        print(f"  PO Usage Status updated: {output_path}")
        return output_path
    except Exception:
        if os.path.isfile(output_path):
            try:
                os.remove(output_path)
            except OSError:
                pass
        raise


def main() -> None:
    """Compatibility CLI: update from Budget Report values using the same template."""
    setup_utf8_stdout()
    import argparse
    from budget_report_writer import BudgetReportReader

    parser = argparse.ArgumentParser(description="Template-preserving PO Usage Status updater")
    parser.add_argument("--project", default="CNTS-25165", help="Project code")
    parser.add_argument("--as-of", default=None, help="As-of month (YYYY-MM)")
    parser.add_argument("--output", default=None, help="Output path")
    args = parser.parse_args()

    project_config = config.PROJECTS.get(args.project, {})
    budget_dir = project_config.get("budget_report_dir", "")
    if not budget_dir:
        raise RuntimeError(f"Budget Report directory not configured for {args.project}")
    reports = [
        path for path in glob.glob(os.path.join(budget_dir, "IAV_Budget_Report_*.xlsx"))
        if "TEMPLATE_V2" not in os.path.basename(path).upper()
    ]
    if not reports:
        raise FileNotFoundError(f"No Budget Report found in {budget_dir}")
    budget_path = max(reports, key=os.path.getmtime)
    personnel = config.CNTS25165_PERSONNEL if args.project == "CNTS-25165" else {}
    if not personnel:
        raise RuntimeError(f"No personnel configuration for {args.project}")
    reader = BudgetReportReader(budget_path, personnel)
    actuals = reader.read_actuals()
    forecast = reader.read_forecast()

    def to_po_allocations(source: dict) -> dict:
        return {
            name: {month: {"Budget_Report": days} for month, days in months.items()}
            for name, months in source.items()
        }

    output_dir = project_config.get("po_usage_dir", budget_dir)
    output_path = args.output or os.path.join(
        output_dir, f"PO_Usage_Status_{datetime.now().strftime('%y%m%d_%H%M%S')}.xlsx"
    )
    generate_po_usage_status(
        budget_path=budget_path,
        output_path=output_path,
        personnel=personnel,
        merged_allocation=to_po_allocations(actuals),
        person_forecast=to_po_allocations(forecast),
        as_of_month=args.as_of,
    )


if __name__ == "__main__":
    main()
