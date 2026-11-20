"""Template adapters and safe, read-only Excel validation for document outputs.

The adapters deliberately own template shape and diff concerns.  Business
parsers should provide structured facts, never hard-coded Excel coordinates.
"""

from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Optional

from openpyxl import load_workbook


def _normalise_header(value: Any) -> str:
    return " ".join(str(value or "").replace("\n", " ").split()).casefold()


def _serialise_value(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


@dataclass
class TemplateValidation:
    """JSON-safe result shared by the API, preview UI and tests."""

    adapter: str
    path: str
    status: str = "PASS"
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        return self.status != "BLOCKED" and not self.errors

    def as_dict(self) -> dict[str, Any]:
        return {
            "adapter": self.adapter,
            "path": self.path,
            "status": self.status,
            "ok": self.ok,
            "errors": list(self.errors),
            "warnings": list(self.warnings),
            "metadata": self.metadata,
        }


class TemplateAdapter:
    """Small common interface for deterministic Excel template adapters."""

    adapter_name = "TemplateAdapter"

    def validate_template(self, path: str) -> TemplateValidation:
        raise NotImplementedError

    def validate_output(self, path: str) -> TemplateValidation:
        return self.validate_template(path)


class TravelExpenseTemplateAdapter(TemplateAdapter):
    """Adapter for the Vendor Weekly Business Trip Expense Claim workbook."""

    adapter_name = "TravelExpenseTemplateAdapter"
    HEADER_ALIASES = {
        "date": {"date"},
        "po_no": {"travel cost po no.", "travel cost po no", "po no.", "po no"},
        "name": {"name"},
        "city": {"city"},
        "city_layer": {"city layer"},
        "ticket": {"flight/train ticket", "flight / train ticket"},
        "taxi": {"taxi"},
        "accommodation": {"accomodation", "accommodation"},
        "allowance": {"allowance"},
        "sum_without_tax": {"sum w/o tax", "sum without tax"},
        "sum_with_tax": {"sum with tax"},
    }

    def _load(self, path: str):
        suffix = Path(path).suffix.lower()
        return load_workbook(path, data_only=False, keep_vba=suffix == ".xlsm")

    def _find_summary_row(self, ws) -> Optional[int]:
        for row in range(2, ws.max_row + 1):
            value = ws.cell(row, 15).value
            if isinstance(value, str) and "SUMIF" in value.upper():
                return row
        return None

    def validate_template(self, path: str) -> TemplateValidation:
        result = TemplateValidation(self.adapter_name, os.path.abspath(path))
        if not os.path.isfile(path):
            result.status = "BLOCKED"
            result.errors.append(f"模板不存在: {path}")
            return result
        if Path(path).suffix.lower() not in {".xlsx", ".xlsm"}:
            result.status = "BLOCKED"
            result.errors.append("Travel 模板必须是 .xlsx 或 .xlsm 文件。")
            return result

        workbook = None
        try:
            workbook = self._load(path)
            if not workbook.worksheets:
                result.status = "BLOCKED"
                result.errors.append("工作簿没有可用工作表。")
                return result
            ws = workbook.worksheets[0]
            headers: dict[str, str] = {}
            for col in range(1, ws.max_column + 1):
                header = _normalise_header(ws.cell(1, col).value)
                if header:
                    headers[header] = ws.cell(1, col).coordinate
            missing = []
            for field_name, aliases in self.HEADER_ALIASES.items():
                if not any(alias in headers for alias in aliases):
                    missing.append(field_name)
            if missing:
                result.status = "BLOCKED"
                result.errors.append(f"差旅模板缺少字段列: {', '.join(missing)}")
            summary_row = self._find_summary_row(ws)
            if summary_row is None:
                result.status = "BLOCKED"
                result.errors.append("差旅模板缺少 Sum w/o tax 的 SUMIF 汇总行。")
            if ws.max_column < 16:
                result.status = "BLOCKED"
                result.errors.append("差旅模板列数不足，至少需要 P 列。")
            result.metadata.update({
                "sheet": ws.title,
                "sheet_count": len(workbook.worksheets),
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "header_cells": headers,
                "summary_row": summary_row,
                "formula_count": sum(
                    1 for row in ws.iter_rows() for cell in row
                    if isinstance(cell.value, str) and cell.value.startswith("=")
                ),
            })
        except Exception as exc:
            result.status = "BLOCKED"
            result.errors.append(f"无法读取差旅模板: {exc}")
        finally:
            if workbook is not None:
                workbook.close()
        return result

    def validate_output(self, path: str) -> TemplateValidation:
        result = self.validate_template(path)
        if not result.ok:
            return result
        if result.metadata.get("formula_count", 0) == 0:
            result.status = "BLOCKED"
            result.errors.append("输出文件没有保留任何公式，疑似模板写入破坏。")
        return result

    def audit_formula_calculation(
        self,
        path: str,
        expected_cells: Iterable[dict[str, Any]],
    ) -> dict[str, Any]:
        """Check formula definitions and cached values after a spreadsheet recalculation.

        openpyxl can preserve or write formulas but cannot calculate them. Missing
        cached values therefore remain PENDING; actual formula errors or incorrect
        values are BLOCKED.
        """
        formula_book = value_book = None
        cells: list[dict[str, Any]] = []
        formula_errors: list[dict[str, Any]] = []
        uncached_formulas: list[dict[str, Any]] = []
        blank_formula_results: list[dict[str, Any]] = []
        mismatches: list[dict[str, Any]] = []
        pending = False
        formula_count = 0
        suffix = Path(path).suffix.lower()
        try:
            formula_book = load_workbook(path, data_only=False, keep_vba=suffix == ".xlsm")
            value_book = load_workbook(path, data_only=True, keep_vba=suffix == ".xlsm")
            formula_count = sum(
                1 for ws in formula_book.worksheets for row in ws.iter_rows()
                for cell in row if isinstance(cell.value, str) and cell.value.startswith("=")
            )
            for formula_ws in formula_book.worksheets:
                value_ws = value_book[formula_ws.title]
                for row in formula_ws.iter_rows():
                    for formula_cell in row:
                        if not (isinstance(formula_cell.value, str) and formula_cell.value.startswith("=")):
                            continue
                        cached_value = value_ws[formula_cell.coordinate].value
                        if isinstance(cached_value, str) and cached_value.startswith("#"):
                            formula_errors.append({
                                "sheet": formula_ws.title,
                                "cell": formula_cell.coordinate,
                                "formula": formula_cell.value,
                                "cached_value": cached_value,
                            })
                        elif cached_value is None:
                            item = {
                                "sheet": formula_ws.title,
                                "cell": formula_cell.coordinate,
                                "formula": formula_cell.value,
                            }
                            if '""' in formula_cell.value:
                                blank_formula_results.append(item)
                            else:
                                uncached_formulas.append(item)

            for expected in expected_cells:
                sheet_name = str(expected.get("sheet") or formula_book.worksheets[0].title)
                coordinate = str(expected["cell"])
                item = {
                    "sheet": sheet_name,
                    "cell": coordinate,
                    "expected_formula": expected.get("formula"),
                    "expected_value": expected.get("value"),
                    "tolerance": float(expected.get("tolerance", 0.005)),
                }
                if sheet_name not in formula_book.sheetnames:
                    item.update({"status": "BLOCKED", "error": "工作表不存在"})
                    cells.append(item)
                    mismatches.append(item)
                    continue
                formula_value = formula_book[sheet_name][coordinate].value
                cached_value = value_book[sheet_name][coordinate].value
                item["formula"] = formula_value
                item["cached_value"] = _serialise_value(cached_value)
                expected_formula = expected.get("formula")
                if not isinstance(formula_value, str) or not formula_value.startswith("="):
                    item.update({"status": "BLOCKED", "error": "目标单元格没有公式"})
                    mismatches.append(item)
                elif expected_formula and re.sub(r"\s+", "", formula_value).casefold() != re.sub(
                    r"\s+", "", str(expected_formula),
                ).casefold():
                    item.update({"status": "BLOCKED", "error": "公式与预期不一致"})
                    mismatches.append(item)
                elif isinstance(cached_value, str) and cached_value.startswith("#"):
                    item.update({"status": "BLOCKED", "error": "公式缓存为 Excel 错误值"})
                    mismatches.append(item)
                elif cached_value is None:
                    item.update({"status": "PENDING", "error": "缺少 Excel 公式缓存值"})
                    pending = True
                elif not isinstance(cached_value, (int, float)):
                    item.update({"status": "BLOCKED", "error": "公式缓存不是数值"})
                    mismatches.append(item)
                elif abs(float(cached_value) - float(expected["value"])) > item["tolerance"]:
                    item.update({"status": "BLOCKED", "error": "公式缓存值与独立期望值不符"})
                    mismatches.append(item)
                else:
                    item["status"] = "PASS"
                cells.append(item)
        except Exception as exc:
            return {
                "status": "BLOCKED",
                "formula_count": formula_count,
                "formula_error_count": len(formula_errors),
                "formula_errors": formula_errors,
                "uncached_formula_count": len(uncached_formulas),
                "uncached_formulas": uncached_formulas[:100],
                "blank_formula_result_count": len(blank_formula_results),
                "mismatches": mismatches,
                "cells": cells,
                "error": f"无法读取重算后的工作簿: {type(exc).__name__}: {exc}",
            }
        finally:
            if formula_book is not None:
                formula_book.close()
            if value_book is not None:
                value_book.close()
        pending = pending or bool(uncached_formulas)
        status = "BLOCKED" if formula_errors or mismatches else (
            "PENDING_EXCEL_RECALCULATION" if pending else "PASS"
        )
        return {
            "status": status,
            "formula_count": formula_count,
            "formula_error_count": len(formula_errors),
            "formula_errors": formula_errors,
            "uncached_formula_count": len(uncached_formulas),
            "uncached_formulas": uncached_formulas[:100],
            "blank_formula_result_count": len(blank_formula_results),
            "mismatch_count": len(mismatches),
            "mismatches": mismatches,
            "cells": cells,
        }

    def audit_preserved_content(
        self,
        template_path: str,
        output_path: str,
        *,
        insertions: Iterable[dict[str, int]] = (),
        writable_row_ranges: Iterable[tuple[int, int]] = (),
        summary_row: Optional[int] = None,
    ) -> dict[str, Any]:
        """Verify original non-output cells/formulas survived template insertion."""
        template_book = output_book = None
        mismatches: list[dict[str, Any]] = []
        checked = 0
        skipped = 0
        insertions = list(insertions)
        writable_row_ranges = list(writable_row_ranges)
        try:
            template_book = self._load(template_path)
            output_book = self._load(output_path)
            missing_sheets = [
                name for name in template_book.sheetnames
                if name not in output_book.sheetnames
            ]
            if missing_sheets:
                mismatches.append({
                    "sheet": "<workbook>",
                    "template": template_book.sheetnames,
                    "output": output_book.sheetnames,
                    "missing_sheets": missing_sheets,
                })
            for template_ws in template_book.worksheets:
                if template_ws.title not in output_book.sheetnames:
                    continue
                output_ws = output_book[template_ws.title]
                for row in range(1, template_ws.max_row + 1):
                    output_row = row
                    for insertion in insertions:
                        if output_row >= int(insertion["row"]):
                            output_row += int(insertion["count"])
                    if any(first <= output_row <= last for first, last in writable_row_ranges):
                        skipped += template_ws.max_column
                        continue
                    for column in range(1, template_ws.max_column + 1):
                        original = template_ws.cell(row, column).value
                        if (
                            summary_row is not None
                            and (
                                (row == summary_row and column in {15, 16})
                                or (
                                    row == summary_row + 1
                                    and column == 15
                                    and isinstance(original, str)
                                    and original.startswith("=")
                                )
                            )
                        ):
                            # These formulas are intentionally extended after trip rows are inserted.
                            skipped += 1
                            continue
                        checked += 1
                        template_cell = template_ws.cell(row, column)
                        output_cell = output_ws.cell(output_row, column)
                        actual = output_cell.value
                        value_changed = original != actual
                        style_changed = template_cell.style_id != output_cell.style_id
                        if value_changed or style_changed:
                            mismatches.append({
                                "sheet": template_ws.title,
                                "template_cell": template_cell.coordinate,
                                "output_cell": output_cell.coordinate,
                                "template_value": _serialise_value(original),
                                "output_value": _serialise_value(actual),
                                "style_changed": style_changed,
                                "template_style_id": template_cell.style_id,
                                "output_style_id": output_cell.style_id,
                            })
        except Exception as exc:
            return {
                "status": "BLOCKED",
                "cells_checked": checked,
                "cells_skipped": skipped,
                "mismatch_count": len(mismatches),
                "mismatches": mismatches,
                "error": f"无法核对模板保留区域: {type(exc).__name__}: {exc}",
            }
        finally:
            if template_book is not None:
                template_book.close()
            if output_book is not None:
                output_book.close()
        return {
            "status": "BLOCKED" if mismatches else "PASS",
            "cells_checked": checked,
            "cells_skipped": skipped,
            "mismatch_count": len(mismatches),
            "mismatches": mismatches[:100],
        }

    @staticmethod
    def compare_workbooks(before_path: str, after_path: str, max_changes: int = 500) -> dict[str, Any]:
        """Compare formula-preserving workbook values without modifying either file."""
        before = load_workbook(
            before_path, data_only=False, keep_vba=Path(before_path).suffix.lower() == ".xlsm"
        )
        after = load_workbook(
            after_path, data_only=False, keep_vba=Path(after_path).suffix.lower() == ".xlsm"
        )
        changes: list[dict[str, Any]] = []
        try:
            sheet_names = sorted(set(before.sheetnames) | set(after.sheetnames))
            for sheet_name in sheet_names:
                if sheet_name not in before.sheetnames or sheet_name not in after.sheetnames:
                    changes.append({"sheet": sheet_name, "cell": "<sheet>", "before": sheet_name in before.sheetnames, "after": sheet_name in after.sheetnames})
                    continue
                left = before[sheet_name]
                right = after[sheet_name]
                max_row = max(left.max_row, right.max_row)
                max_col = max(left.max_column, right.max_column)
                for row in range(1, max_row + 1):
                    for col in range(1, max_col + 1):
                        old = left.cell(row, col).value
                        new = right.cell(row, col).value
                        if old == new:
                            continue
                        changes.append({
                            "sheet": sheet_name,
                            "cell": right.cell(row, col).coordinate,
                            "before": _serialise_value(old),
                            "after": _serialise_value(new),
                        })
                        if len(changes) >= max_changes:
                            return {
                                "changes": changes,
                                "change_count": len(changes),
                                "truncated": True,
                                "sheets": sheet_names,
                            }
            return {
                "changes": changes,
                "change_count": len(changes),
                "truncated": False,
                "sheets": sheet_names,
            }
        finally:
            before.close()
            after.close()


class AcceptanceTemplateAdapter(TemplateAdapter):
    """Adapter for the existing Overview/Breakdown acceptance family.

    Acceptance workbooks differ by customer and period.  This adapter validates
    the stable anchors and only writes metadata when a project profile provides
    explicit label-based mappings; it never guesses customer-specific cells.
    """

    adapter_name = "AcceptanceTemplateAdapter"

    def _find_label_cell(self, sheet, label: str):
        expected = _normalise_header(label)
        for row in sheet.iter_rows():
            for cell in row:
                if _normalise_header(cell.value) == expected:
                    return cell
        return None

    def _mapping_specs(self, mapping: dict[str, Any]) -> dict[str, Any]:
        metadata = mapping.get("metadata", {}) if isinstance(mapping, dict) else {}
        return metadata if isinstance(metadata, dict) else {}

    def validate_mapping(self, path: str, mapping: dict[str, Any]) -> TemplateValidation:
        """Validate configured labels before any Acceptance write is allowed."""
        result = self.validate_template(path)
        specs = self._mapping_specs(mapping)
        if not specs:
            result.status = "BLOCKED"
            result.errors.append("Acceptance 缺少显式 metadata 字段映射。")
            return result
        workbook = None
        try:
            workbook = load_workbook(path, data_only=False, keep_vba=Path(path).suffix.lower() == ".xlsm")
            missing = []
            targets = {}
            default_sheet = str(mapping.get("sheet", "Overview"))
            default_offset = int(mapping.get("value_column_offset", 1))
            default_row_offset = int(mapping.get("value_row_offset", 0))
            for field_name, spec in specs.items():
                if isinstance(spec, dict):
                    label = spec.get("label", field_name)
                    sheet_name = str(spec.get("sheet", default_sheet))
                    col_offset = int(spec.get("value_column_offset", default_offset))
                    row_offset = int(spec.get("value_row_offset", default_row_offset))
                else:
                    label = spec
                    sheet_name = default_sheet
                    col_offset = default_offset
                    row_offset = default_row_offset
                if sheet_name not in workbook.sheetnames:
                    missing.append(f"{field_name}（工作表 {sheet_name} 不存在）")
                    continue
                label_cell = self._find_label_cell(workbook[sheet_name], str(label))
                if label_cell is None:
                    missing.append(f"{field_name}（标签 {label} 不存在）")
                    continue
                target = workbook[sheet_name].cell(
                    row=label_cell.row + row_offset,
                    column=label_cell.column + col_offset,
                )
                targets[field_name] = {
                    "sheet": sheet_name,
                    "label_cell": label_cell.coordinate,
                    "target_cell": target.coordinate,
                    "label": str(label),
                }
            if missing:
                result.status = "BLOCKED"
                result.errors.extend([f"Acceptance 映射目标不存在: {item}" for item in missing])
            result.metadata["mapping_targets"] = targets
        except Exception as exc:
            result.status = "BLOCKED"
            result.errors.append(f"无法校验 Acceptance 映射: {exc}")
        finally:
            if workbook is not None:
                workbook.close()
        return result

    def preview_metadata(
        self,
        path: str,
        values: dict[str, Any],
        mapping: dict[str, Any],
    ) -> dict[str, Any]:
        """Return metadata cell changes without writing the workbook."""
        validation = self.validate_mapping(path, mapping)
        if not validation.ok:
            return {"ok": False, "validation": validation.as_dict(), "changes": []}
        workbook = load_workbook(path, data_only=False, keep_vba=Path(path).suffix.lower() == ".xlsm")
        changes = []
        try:
            for field_name, target in validation.metadata["mapping_targets"].items():
                if field_name not in values:
                    continue
                cell = workbook[target["sheet"]][target["target_cell"]]
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    validation.status = "BLOCKED"
                    validation.errors.append(
                        f"Acceptance 映射目标 {target['sheet']}!{target['target_cell']} 是公式，禁止覆盖。"
                    )
                    continue
                new_value = values[field_name]
                if cell.value != new_value:
                    changes.append({
                        "sheet": target["sheet"],
                        "cell": target["target_cell"],
                        "field": field_name,
                        "before": _serialise_value(cell.value),
                        "after": _serialise_value(new_value),
                    })
        finally:
            workbook.close()
        if not validation.ok:
            return {"ok": False, "validation": validation.as_dict(), "changes": changes}
        return {
            "ok": True,
            "validation": validation.as_dict(),
            "changes": changes,
            "change_count": len(changes),
        }

    def write_metadata(
        self,
        template_path: str,
        output_path: str,
        values: dict[str, Any],
        mapping: dict[str, Any],
    ) -> dict[str, Any]:
        """Write only explicitly mapped metadata to a versioned copy."""
        template = os.path.abspath(template_path)
        output = os.path.abspath(output_path)
        if template == output:
            raise ValueError("Acceptance 输出不能覆盖原始模板")
        preview = self.preview_metadata(template, values, mapping)
        if not preview["ok"]:
            raise ValueError("Acceptance 模板或映射校验失败: " + "；".join(preview["validation"]["errors"]))
        os.makedirs(os.path.dirname(output), exist_ok=True)
        shutil.copy2(template, output)
        workbook = load_workbook(output, data_only=False, keep_vba=Path(output).suffix.lower() == ".xlsm")
        try:
            targets = preview["validation"]["metadata"]["mapping_targets"]
            for field_name, target in targets.items():
                if field_name in values:
                    workbook[target["sheet"]][target["target_cell"]] = values[field_name]
            workbook.save(output)
        finally:
            workbook.close()
        output_validation = self.validate_mapping(output, mapping)
        if not output_validation.ok:
            raise ValueError("Acceptance 输出校验失败: " + "；".join(output_validation.errors))
        return {
            "ok": True,
            "output_path": output,
            "changes": preview["changes"],
            "change_count": preview["change_count"],
            "validation": output_validation.as_dict(),
        }

    def validate_template(self, path: str) -> TemplateValidation:
        result = TemplateValidation(self.adapter_name, os.path.abspath(path))
        if not os.path.isfile(path):
            result.status = "BLOCKED"
            result.errors.append(f"验收模板不存在: {path}")
            return result
        workbook = None
        try:
            workbook = load_workbook(path, data_only=False, keep_vba=Path(path).suffix.lower() == ".xlsm")
            names = set(workbook.sheetnames)
            if "Overview" not in names:
                result.status = "BLOCKED"
                result.errors.append("验收模板缺少 Overview 工作表。")
            detail_sheet = "Breakdown" if "Breakdown" in names else None
            if detail_sheet is None:
                result.status = "WARNING"
                result.warnings.append("验收模板没有 Breakdown 工作表，需要项目级明细映射。")
            overview = workbook["Overview"] if "Overview" in names else None
            labels = set()
            if overview is not None:
                for row in overview.iter_rows(min_row=1, max_row=min(overview.max_row, 40), max_col=min(overview.max_column, 12)):
                    for cell in row:
                        label = _normalise_header(cell.value)
                        if label:
                            labels.add(label)
            for required in ("po number", "service period"):
                if required not in labels:
                    result.warnings.append(f"Overview 未发现稳定字段: {required}")
            result.metadata.update({
                "sheet_count": len(names),
                "sheet_names": sorted(names),
                "detail_sheet": detail_sheet,
                "overview_labels": sorted(labels),
            })
        except Exception as exc:
            result.status = "BLOCKED"
            result.errors.append(f"无法读取验收模板: {exc}")
        finally:
            if workbook is not None:
                workbook.close()
        return result
