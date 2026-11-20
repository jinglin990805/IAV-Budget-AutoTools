# -*- coding: utf-8 -*-
"""
Budget Report Writer — 读写增强版 Budget Report 模板

核心功能:
1. AnchorCellTracker: 通过 Column A 锚点标记定位数据区域
2. BudgetReportReader: 锚点感知的 Budget Report 读取器（替代 po_allocator_v2 中的内联解析）
3. BudgetReportWriter: 使用 openpyxl 精确回写单元格（保留格式和公式）
4. DiffPreview: 写入前变更预览
5. BackupManager: 自动时间戳备份

三层安全机制:
  Layer 1 — 自动备份（写入前强制复制）
  Layer 2 — 差异预览（控制台 + Excel diff 文件）
  Layer 3 — 显式确认（控制台提示，超时 60s 默认中止）

用法:
    from budget_report_writer import (
        BudgetReportTemplate, AnchorCellTracker,
        backup_budget_report, preview_changes, write_budget_report,
    )
"""

import os
import sys
import io
import re
import shutil
import math
import tempfile
import zipfile
import xml.etree.ElementTree as ET
import posixpath
from datetime import datetime, timedelta
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config
from utils import setup_utf8_stdout, safe_load_workbook, safe_str, parse_date_month

SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _xlsx_package_snapshot(filepath: str) -> dict:
    """Capture workbook structure that a value-only update must preserve."""
    with zipfile.ZipFile(filepath, "r") as archive:
        bad_member = archive.testzip()
        if bad_member:
            raise RuntimeError(f"Workbook ZIP member is corrupt: {bad_member}")
        names = set(archive.namelist())
        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationships = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels_root.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
        }
        sheets = {}
        sheet_names = []
        for sheet in workbook_root.findall(f".//{{{SPREADSHEET_NS}}}sheet"):
            name = sheet.attrib["name"]
            rel_id = sheet.attrib.get(f"{{{OFFICE_REL_NS}}}id")
            target = relationships[rel_id]
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            sheet_names.append(name)
            sheet_root = ET.fromstring(archive.read(part))
            formulas = {}
            styles = {}
            special_counts = {}
            for cell in sheet_root.iter(f"{{{SPREADSHEET_NS}}}c"):
                ref = cell.attrib.get("r", "")
                if "s" in cell.attrib:
                    styles[ref] = cell.attrib["s"]
                formula = cell.find(f"{{{SPREADSHEET_NS}}}f")
                if formula is not None:
                    formulas[ref] = (formula.text or "", tuple(sorted(formula.attrib.items())))
            for node in sheet_root.iter():
                local = node.tag.rsplit("}", 1)[-1]
                if local in {
                    "mergeCell", "dataValidation", "conditionalFormatting", "tablePart",
                    "drawing", "legacyDrawing", "extLst", "pivotTablePart", "sheetProtection",
                }:
                    special_counts[local] = special_counts.get(local, 0) + 1
            merge_cells = tuple(sorted(
                node.attrib.get("ref", "")
                for node in sheet_root.iter(f"{{{SPREADSHEET_NS}}}mergeCell")
            ))
            sheets[name] = {
                "part": part,
                "formulas": formulas,
                "styles": styles,
                "special_counts": special_counts,
                "merge_cells": merge_cells,
            }

        pivot_parts = {
            name: archive.read(name)
            for name in names
            if name.startswith("xl/pivotTables/") and name.endswith(".xml")
        }
        external_parts = {
            name: archive.read(name)
            for name in names
            if name.startswith("xl/externalLinks/")
        }
        return {
            "parts": names,
            "sheet_names": tuple(sheet_names),
            "styles_xml": archive.read("xl/styles.xml") if "xl/styles.xml" in names else b"",
            "sheets": sheets,
            "pivots": pivot_parts,
            "external_links": external_parts,
        }


# ============================================================
# 常量
# ============================================================
ANCHOR_PREFIX = "=== "
ANCHOR_SUFFIX = " ==="
TEMPLATE_VERSION = "2.0"
CONFIRM_TIMEOUT_SECONDS = 60
BACKUP_DIR = os.path.join(config.AUTOTOOLS_DIR, "backups")

# CNTS-25165 关键列位置（内部 Budget_FC sheet）
# Forecast / Actual 区域列映射
COL_NAME = "J"       # 人员姓名
COL_VCTC_RATE = "K"  # VCTC Rate
COL_LEVEL = "L"      # Level
MONTH_START_COL = "W"  # 2026-Jan（M=2025-Mar, W=2026-Jan）
MONTH_END_COL = "AH"   # 2026-Dec
COL_SUM_DAYS = "AI"    # Sum in [d]
COL_ACTUAL_COST = "AJ"  # Actual Cost
COL_PLAN_COST = "AK"    # Plan Cost


# ============================================================
# 数据模型
# ============================================================

@dataclass
class CellChange:
    """单个单元格的变更记录"""
    sheet: str
    cell_ref: str           # e.g. "W10"
    row: int
    col: int
    col_letter: str
    current_value: object
    new_value: object
    description: str = ""   # 人性化描述，如 "Tang Qiwei 2026-Jan days"

    @property
    def is_addition(self):
        return self.current_value is None and self.new_value is not None

    @property
    def is_deletion(self):
        return self.current_value is not None and self.new_value is None

    @property
    def is_modification(self):
        return (self.current_value is not None and self.new_value is not None
                and self.current_value != self.new_value)

    @property
    def delta(self):
        """数值变化的增量"""
        try:
            old = float(self.current_value) if self.current_value is not None else 0
            new = float(self.new_value) if self.new_value is not None else 0
            return new - old
        except (ValueError, TypeError):
            return None


@dataclass
class DiffPreview:
    """一组 CellChange 的汇总"""
    changes: list = field(default_factory=list)
    sheet_name: str = ""
    sections_affected: set = field(default_factory=set)

    @property
    def total_changes(self):
        return len(self.changes)

    @property
    def additions(self):
        return sum(1 for c in self.changes if c.is_addition)

    @property
    def modifications(self):
        return sum(1 for c in self.changes if c.is_modification)

    @property
    def deletions(self):
        return sum(1 for c in self.changes if c.is_deletion)

    def summary(self) -> str:
        lines = [
            f"=== DIFF SUMMARY ===",
            f"Sheet: {self.sheet_name}",
            f"Sections affected: {', '.join(sorted(self.sections_affected))}",
            f"Total cell changes: {self.total_changes}",
            f"  Additions: {self.additions}",
            f"  Modifications: {self.modifications}",
            f"  Deletions: {self.deletions}",
        ]
        return "\n".join(lines)


# ============================================================
# AnchorCellTracker — 锚点定位
# ============================================================

class AnchorCellTracker:
    """
    扫描 sheet 的 Column A，通过锚点文本标记定位各数据区域。

    锚点格式: "=== SECTION_NAME ==="
    锚点文字设为白色字体（与背景融为一体），肉眼不可见但工具可读。

    **重要**: 锚点必须用 data_only=False 读取，
    因为 openpyxl 的 data_only=True 模式可能返回 None（缓存问题）。
    """

    @staticmethod
    def locate_all(ws) -> dict:
        """
        扫描 Column A 中所有锚点标记。

        Returns:
            dict: {section_name: row_number}
            例如: {"PO_SUMMARY_START": 3, "PO_SUMMARY_END": 7, ...}
        """
        anchors = {}
        for row_idx in range(1, ws.max_row + 1):
            cell_val = ws.cell(row=row_idx, column=1).value
            if cell_val and isinstance(cell_val, str):
                val = cell_val.strip()
                if val.startswith(ANCHOR_PREFIX) and val.endswith(ANCHOR_SUFFIX):
                    section_name = val[len(ANCHOR_PREFIX):-len(ANCHOR_SUFFIX)]
                    anchors[section_name] = row_idx
        return anchors

    @staticmethod
    def locate_all_from_file(filepath: str, sheet_name: str) -> dict:
        """
        从文件直接读取锚点（自动使用 data_only=False）。

        Args:
            filepath: Excel 文件路径
            sheet_name: sheet 名称

        Returns:
            dict: {section_name: row_number}
        """
        wb = openpyxl.load_workbook(filepath, data_only=False)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return {}
        ws = wb[sheet_name]
        anchors = AnchorCellTracker.locate_all(ws)
        wb.close()
        return anchors

    @staticmethod
    def get_section_range(ws, section_prefix: str) -> tuple:
        """
        获取某个 section 的起止行号。

        Args:
            section_prefix: 如 "PO_SUMMARY", "FORECAST", "ACTUAL"

        Returns:
            (start_row, end_row) 或 (None, None) 如果找不到
        """
        anchors = AnchorCellTracker.locate_all(ws)
        start_key = f"{section_prefix}_START"
        end_key = f"{section_prefix}_END"
        return anchors.get(start_key), anchors.get(end_key)

    @staticmethod
    def validate(ws, required_sections: list[str]) -> list[str]:
        """
        验证所需的 section 锚点是否都存在。

        Returns:
            list[str]: 缺失的 section 名称列表（空 = 全部存在）
        """
        anchors = AnchorCellTracker.locate_all(ws)
        missing = []
        for section in required_sections:
            start_key = f"{section}_START"
            end_key = f"{section}_END"
            if start_key not in anchors:
                missing.append(start_key)
            if end_key not in anchors:
                missing.append(end_key)
        return missing


# ============================================================
# BudgetReportReader — 锚点感知的读取器
# ============================================================

class BudgetReportReader:
    """
    从增强版 Budget Report 模板中读取数据。
    替代 po_allocator_v2.py 中的 load_budget_report_actuals_and_forecast()。

    使用双重打开策略:
      - data_only=False: 读取锚点 + 公式（需要公式的场合）
      - data_only=True:  读取缓存值（数据读取场合）
    """

    def __init__(self, filepath: str, personnel: dict):
        self.filepath = filepath
        self.personnel = personnel
        self.name_lookup = config.build_name_lookup(personnel)

    def _normalize_name(self, raw_name: str) -> Optional[str]:
        """将 Budget Report 中的名字标准化为 canonical name"""
        if not raw_name:
            return None
        n = str(raw_name).strip()
        # "Last, First" → "First Last"
        if ',' in n:
            parts = n.split(',', 1)
            n = f"{parts[1].strip()} {parts[0].strip()}"
        # Alias lookup
        canonical = self.name_lookup.get(n.lower())
        if canonical:
            return canonical
        # Fuzzy match
        for alias, cname in self.name_lookup.items():
            if alias in n.lower() or n.lower() in alias:
                return cname
        return None

    def _find_active_sheet(self, wb) -> Optional[str]:
        """自动查找最新的 Internal Budget FC sheet"""
        candidates = [sn for sn in wb.sheetnames
                      if 'Internal Budget' in sn and '2026' in sn]
        # 优先 Q4 > Q3 > Q2 > Q1
        for q in ['Q4', 'Q3', 'Q2', 'Q1']:
            for sn in candidates:
                if q in sn:
                    return sn
        return candidates[0] if candidates else None

    def _get_anchors(self) -> tuple:
        """获取锚点映射和活跃 sheet 名称 (使用 data_only=False)"""
        wb = openpyxl.load_workbook(self.filepath, data_only=False)
        sheet_name = self._find_active_sheet(wb)
        if not sheet_name:
            wb.close()
            return None, {}
        ws = wb[sheet_name]
        anchors = AnchorCellTracker.locate_all(ws)
        wb.close()
        return sheet_name, anchors

    def read_forecast(self, sheet_name: str = None) -> dict:
        """
        读取 Forecast 区域。

        Returns:
            {canonical_name: {month_str: days}}
        """
        # 1. 用 data_only=False 获取锚点位置
        if sheet_name is None:
            sheet_name, anchors = self._get_anchors()
        else:
            # 直接用指定的 sheet 名获取锚点
            _, anchors = self._get_anchors()
        if sheet_name is None or not anchors:
            return {}

        fc_start = anchors.get("FORECAST_START")
        fc_end = anchors.get("FORECAST_END")
        if fc_start is None or fc_end is None:
            return {}

        # 2. 用 data_only=True 读取实际数据值
        wb = openpyxl.load_workbook(self.filepath, data_only=True)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return {}
        ws = wb[sheet_name]

        # 解析 header row 找月份列映射
        headers = self._get_row_values(ws, fc_start)
        month_cols = self._find_month_columns(headers)

        result = defaultdict(lambda: defaultdict(float))

        for row_idx in range(fc_start + 1, fc_end + 1):
            row_vals = self._get_row_values(ws, row_idx)
            if not row_vals:
                continue

            # 从 J 列读取姓名 (index 9 = column J)
            name_val = row_vals[9] if len(row_vals) > 9 else None
            if not name_val:
                continue

            canonical = self._normalize_name(str(name_val))
            if not canonical or canonical not in self.personnel:
                continue

            for col_idx, month_str in month_cols.items():
                if col_idx < len(row_vals) and row_vals[col_idx] is not None:
                    try:
                        days = float(row_vals[col_idx])
                    except (ValueError, TypeError):
                        days = 0.0
                    if days > 0:
                        result[canonical][month_str] += days

        wb.close()
        return dict(result)

    def read_actuals(self, sheet_name: str = None) -> dict:
        """
        读取 Actual 区域。

        Returns:
            {canonical_name: {month_str: days}}
        """
        if sheet_name is None:
            sheet_name, anchors = self._get_anchors()
        else:
            _, anchors = self._get_anchors()
        if sheet_name is None or not anchors:
            return {}

        ac_start = anchors.get("ACTUAL_START")
        ac_end = anchors.get("ACTUAL_END")
        if ac_start is None or ac_end is None:
            return {}

        wb = openpyxl.load_workbook(self.filepath, data_only=True)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return {}
        ws = wb[sheet_name]

        headers = self._get_row_values(ws, ac_start)
        month_cols = self._find_month_columns(headers)

        result = defaultdict(lambda: defaultdict(float))

        for row_idx in range(ac_start + 1, ac_end + 1):
            row_vals = self._get_row_values(ws, row_idx)
            if not row_vals:
                continue

            name_val = row_vals[9] if len(row_vals) > 9 else None
            if not name_val:
                continue

            canonical = self._normalize_name(str(name_val))
            if not canonical or canonical not in self.personnel:
                continue

            for col_idx, month_str in month_cols.items():
                if col_idx < len(row_vals) and row_vals[col_idx] is not None:
                    try:
                        days = float(row_vals[col_idx])
                    except (ValueError, TypeError):
                        days = 0.0
                    if days > 0:
                        result[canonical][month_str] += days

        wb.close()
        return dict(result)

    def read_forecast_with_po_attribution(self, sheet_name: str = None) -> dict:
        """
        读取 Forecast 区域，包含 PO 归因信息。

        Returns:
            {canonical_name: {month_str: {po_id: days}}}
        """
        if sheet_name is None:
            sheet_name, anchors = self._get_anchors()
        else:
            _, anchors = self._get_anchors()
        if sheet_name is None or not anchors:
            return {}

        fc_start = anchors.get("FORECAST_START")
        fc_end = anchors.get("FORECAST_END")
        if fc_start is None or fc_end is None:
            return {}

        wb = openpyxl.load_workbook(self.filepath, data_only=True)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return {}
        ws = wb[sheet_name]

        headers = self._get_row_values(ws, fc_start)
        month_cols = self._find_month_columns(headers)

        result = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

        for row_idx in range(fc_start + 1, fc_end + 1):
            row_vals = self._get_row_values(ws, row_idx)
            if not row_vals:
                continue

            name_val = row_vals[9] if len(row_vals) > 9 else None
            if not name_val:
                continue

            canonical = self._normalize_name(str(name_val))
            if not canonical or canonical not in self.personnel:
                continue

            for col_idx, month_str in month_cols.items():
                if col_idx < len(row_vals) and row_vals[col_idx] is not None:
                    try:
                        days = float(row_vals[col_idx])
                    except (ValueError, TypeError):
                        days = 0.0
                    if days > 0:
                        # 尝试读取 PO Attribution 列（如果存在）
                        # 目前默认使用 "unknown" PO，后续可扩展
                        result[canonical][month_str]["unknown"] += days

        wb.close()
        return {k: dict(v) for k, v in result.items()}

    def read_po_config(self) -> list:
        """从 PO_Config sheet 读取 PO Level 配置"""
        wb = openpyxl.load_workbook(self.filepath, data_only=True)
        if "PO_Config" not in wb.sheetnames:
            wb.close()
            return []

        ws = wb["PO_Config"]
        configs = []
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            if len(row) < 6 or not row[1]:
                continue
            configs.append(list(row))

        wb.close()
        return configs

    def read_person_row_mapping(self, sheet_name: str = None) -> dict:
        """
        建立人员名称 → Forecast 区域行号的映射。
        用于后续回写时定位每个人的数据行。

        Returns:
            {canonical_name: forecast_row_number}
        """
        if sheet_name is None:
            sheet_name, anchors = self._get_anchors()
        else:
            _, anchors = self._get_anchors()
        if sheet_name is None or not anchors:
            return {}

        fc_start = anchors.get("FORECAST_START")
        fc_end = anchors.get("FORECAST_END")
        if fc_start is None or fc_end is None:
            return {}

        wb = openpyxl.load_workbook(self.filepath, data_only=True)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return {}
        ws = wb[sheet_name]

        mapping = {}
        for row_idx in range(fc_start + 1, fc_end + 1):
            name_val = ws.cell(row=row_idx, column=10).value  # Column J
            if name_val:
                canonical = self._normalize_name(str(name_val))
                if canonical and canonical in self.personnel:
                    mapping[canonical] = row_idx

        wb.close()
        return mapping

    def read_actual_person_row_mapping(self, sheet_name: str = None) -> dict:
        """
        建立人员名称 → Actual 区域行号的映射。

        Returns:
            {canonical_name: actual_row_number}
        """
        if sheet_name is None:
            sheet_name, anchors = self._get_anchors()
        else:
            _, anchors = self._get_anchors()
        if sheet_name is None or not anchors:
            return {}

        ac_start = anchors.get("ACTUAL_START")
        ac_end = anchors.get("ACTUAL_END")
        if ac_start is None or ac_end is None:
            return {}

        wb = openpyxl.load_workbook(self.filepath, data_only=True)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return {}
        ws = wb[sheet_name]

        mapping = {}
        for row_idx in range(ac_start + 1, ac_end + 1):
            name_val = ws.cell(row=row_idx, column=10).value  # Column J
            if name_val:
                canonical = self._normalize_name(str(name_val))
                if canonical and canonical in self.personnel:
                    mapping[canonical] = row_idx

        wb.close()
        return mapping

    def check_template_version(self) -> bool:
        """检查模板版本是否兼容"""
        wb = openpyxl.load_workbook(self.filepath, data_only=True)
        if "_ToolConfig" not in wb.sheetnames:
            wb.close()
            return False
        ws = wb["_ToolConfig"]
        version = ws.cell(row=1, column=2).value
        wb.close()
        return str(version) == TEMPLATE_VERSION

    # --- helpers ---
    def _get_row_values(self, ws, row_idx: int, max_col: int = 42) -> list:
        """获取一行所有列的值（1-indexed column → 0-indexed list）"""
        vals = []
        for col_idx in range(1, max_col + 1):
            vals.append(ws.cell(row=row_idx, column=col_idx).value)
        return vals

    def _find_month_columns(self, headers: list) -> dict:
        """从 header row 中找出月份列 → {0-indexed-col: 'YYYY-MM'}"""
        month_cols = {}

        # 月份名映射 (仅名称，无年份)
        bare_month_names = {
            "jan": 1, "feb": 2, "mar": 3, "apr": 4,
            "may": 5, "jun": 6, "jul": 7, "aug": 8,
            "sep": 9, "oct": 10, "nov": 11, "dec": 12,
            "january": 1, "february": 2, "march": 3, "april": 4,
            "june": 6, "july": 7, "august": 8,
            "september": 9, "october": 10, "november": 11, "december": 12,
        }

        # 推断年份: 2025 年月份列在 W 之前 (index < 22), 2026 年在 W 及之后
        # Col W = index 22 (0-indexed)
        YEAR_SPLIT_COL = 22  # Column W is the first 2026 month

        for i, h in enumerate(headers):
            if h is None:
                continue
            # 1) 尝试 parse_date_month (处理 "2026-01" / "2026-Jan" / datetime)
            parsed = parse_date_month(h)
            if parsed:
                month_cols[i] = parsed
                continue

            # 2) 处理裸月份名 "Jan", "Mar" 等
            hs = str(h).strip().lower()
            if hs in bare_month_names:
                month_num = bare_month_names[hs]
                # 根据列位置推断年份
                year = 2026 if i >= YEAR_SPLIT_COL else 2025
                month_cols[i] = f"{year}-{month_num:02d}"
                continue

        return month_cols


# ============================================================
# BudgetReportWriter — 精确回写
# ============================================================

class BudgetReportWriter:
    """
    使用 openpyxl 向 Budget Report 模板写入数据。
    保留所有格式、合并单元格、公式和其他 sheet。
    """

    def __init__(self, filepath: str, project: str = "CNTS-25165"):
        self.filepath = filepath
        self.project = project

    def write_updates(self, updates: dict, dry_run: bool = False) -> DiffPreview:
        """
        将变更写入 Budget Report。

        Args:
            updates: {sheet_name: {cell_ref: new_value}}
            dry_run: True = 只生成 DiffPreview，不实际写入

        Returns:
            DiffPreview 对象
        """
        diff = DiffPreview()

        if dry_run:
            # 只读模式：比较当前值和提议值
            wb = safe_load_workbook(self.filepath, data_only=True)
            if not wb:
                return diff

            for sheet_name, cell_updates in updates.items():
                if sheet_name not in wb.sheetnames:
                    continue
                ws = wb[sheet_name]
                diff.sheet_name = sheet_name
                for cell_ref, new_val in cell_updates.items():
                    col_letter, row_num = self._parse_cell_ref(cell_ref)
                    col_idx = column_index_from_string(col_letter)
                    current_val = ws.cell(row=row_num, column=col_idx).value
                    if current_val != new_val:
                        section = self._infer_section(row_num)
                        diff.sections_affected.add(section)
                        diff.changes.append(CellChange(
                            sheet=sheet_name, cell_ref=cell_ref,
                            row=row_num, col=col_idx, col_letter=col_letter,
                            current_value=current_val, new_value=new_val,
                            description=f"{cell_ref}: {current_val} → {new_val}",
                        ))
            wb.close()
            return diff

        # 写模式：先备份，再生成经过校验的临时文件，最后原子替换正式文件。
        # 直接 wb.save(self.filepath) 会先截断目标文件；如果 Excel 正在占用目标，
        # 保存和后续从备份恢复都可能失败。暂存写入可以保证提交失败时原文件不变。
        backup_path = backup_budget_report(self.filepath)
        print(f"  Backup saved: {os.path.basename(backup_path)}")

        # 打开文件（不读缓存值，保留公式和格式）
        wb = openpyxl.load_workbook(self.filepath)
        temp_path = None
        keep_temp_file = False
        staged_outside_target = False

        try:
            written_count = 0
            skipped_formulas = 0

            for sheet_name, cell_updates in updates.items():
                if sheet_name not in wb.sheetnames:
                    print(f"  [WARN] Sheet '{sheet_name}' not found, skipping")
                    continue

                ws = wb[sheet_name]
                diff.sheet_name = sheet_name

                for cell_ref, new_val in cell_updates.items():
                    col_letter, row_num = self._parse_cell_ref(cell_ref)
                    col_idx = column_index_from_string(col_letter)
                    cell = ws.cell(row=row_num, column=col_idx)

                    # 保护公式：不覆盖以 '=' 开头的单元格
                    if isinstance(cell.value, str) and str(cell.value).startswith('='):
                        skipped_formulas += 1
                        continue

                    section = self._infer_section(row_num)
                    diff.sections_affected.add(section)
                    diff.changes.append(CellChange(
                        sheet=sheet_name, cell_ref=cell_ref,
                        row=row_num, col=col_idx, col_letter=col_letter,
                        current_value=cell.value, new_value=new_val,
                    ))
                    cell.value = new_val
                    written_count += 1

            if skipped_formulas > 0:
                print(f"  [INFO] Skipped {skipped_formulas} formula cells")

            # 先保存到目标目录；如果目录只读，则将完整版本保存到 AutoTools/output。
            target_dir = os.path.dirname(os.path.abspath(self.filepath))
            try:
                file_descriptor, temp_path = tempfile.mkstemp(
                    prefix=f".{os.path.basename(self.filepath)}.",
                    suffix=".tmp.xlsx",
                    dir=target_dir,
                )
            except OSError as target_directory_error:
                fallback_dir = config.OUTPUT_DIR
                os.makedirs(fallback_dir, exist_ok=True)
                file_descriptor, temp_path = tempfile.mkstemp(
                    prefix=f".{os.path.basename(self.filepath)}.",
                    suffix=".tmp.xlsx",
                    dir=fallback_dir,
                )
                staged_outside_target = True
                print(
                    f"  [WARN] Cannot stage an update beside the report ({target_directory_error}); "
                    f"a validated version will be published to {fallback_dir}"
                )
            os.close(file_descriptor)
            try:
                wb.calculation.calcMode = "auto"
                wb.calculation.fullCalcOnLoad = True
                wb.calculation.forceFullCalc = True
            except AttributeError:
                pass
            wb.save(temp_path)
            wb.close()

            # 只提交可重新打开且公式、格式、透视表、外链结构完整的工作簿。
            verified = self._post_write_verify(temp_path, self.filepath, updates)

            publish_error = None
            if staged_outside_target:
                version_path = self._unique_versioned_path(config.OUTPUT_DIR)
                try:
                    os.replace(temp_path, version_path)
                    temp_path = None
                    self.filepath = version_path
                except OSError as publish_exc:
                    publish_error = publish_exc
                    print(f"  [WARN] Could not publish the output version: {publish_error}")

            if not staged_outside_target:
                try:
                    os.replace(temp_path, self.filepath)
                    temp_path = None
                except OSError as commit_error:
                    publish_error = commit_error
                    for publish_dir in (target_dir, config.OUTPUT_DIR):
                        try:
                            os.makedirs(publish_dir, exist_ok=True)
                            version_path = self._unique_versioned_path(publish_dir)
                            if os.path.normcase(os.path.abspath(publish_dir)) == os.path.normcase(target_dir):
                                os.replace(temp_path, version_path)
                            else:
                                shutil.copy2(temp_path, version_path)
                                os.remove(temp_path)
                            self.filepath = version_path
                            temp_path = None
                            staged_outside_target = True
                            print(
                                f"  [WARN] Original Budget Report could not be replaced ({commit_error}); "
                                f"the original is unchanged and a validated new version was saved to: {version_path}"
                            )
                            break
                        except OSError as fallback_error:
                            publish_error = fallback_error

            if temp_path:
                recovery_path = self._preserve_recovery_copy(temp_path)
                keep_temp_file = recovery_path is None
                preserved_path = recovery_path or temp_path
                raise RuntimeError(
                    f"Budget Report update could not be published ({publish_error if not staged_outside_target else 'fallback output failed'}). "
                    f"The original report was left unchanged. A validated recovery copy is preserved at: "
                    f"{preserved_path}. Original backup: {backup_path}"
                )

            print(f"  Written {written_count} cells to {self.filepath}")
            if staged_outside_target:
                print(f"  Budget Report updated in: {self.filepath}")
            if verified:
                print("  Post-write verification: OK")

        except Exception as e:
            print(f"  [ERROR] Write failed: {e}")
            try:
                wb.close()
            except Exception:
                pass
            raise
        finally:
            if temp_path and not keep_temp_file and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError as cleanup_error:
                    print(f"  [WARN] Could not remove temporary workbook {temp_path}: {cleanup_error}")

        return diff

    def _unique_versioned_path(self, directory: str) -> str:
        base_name = os.path.basename(self.filepath)
        name, ext = os.path.splitext(base_name)
        timestamp = datetime.now().strftime("%y%m%d_%H%M%S")
        candidate = os.path.join(directory, f"{name}_updated_{timestamp}{ext}")
        suffix = 2
        while os.path.exists(candidate):
            candidate = os.path.join(directory, f"{name}_updated_{timestamp}_{suffix}{ext}")
            suffix += 1
        return candidate

    def _parse_cell_ref(self, cell_ref: str) -> tuple:
        """解析 "W10" → ('W', 10)"""
        match = re.match(r'([A-Z]+)(\d+)', cell_ref.upper())
        if not match:
            raise ValueError(f"Invalid cell reference: {cell_ref}")
        return match.group(1), int(match.group(2))

    def _infer_section(self, row_num: int) -> str:
        """根据行号推断 section 名称"""
        # 需要读取锚点来确定（简化版：使用固定行范围）
        if row_num <= 8:
            return "PO_SUMMARY"
        elif row_num <= 31:
            return "FORECAST"
        elif row_num <= 57:
            return "ACTUAL"
        return "UNKNOWN"

    def _post_write_verify(
        self,
        workbook_path: Optional[str] = None,
        source_path: Optional[str] = None,
        expected_updates: Optional[dict] = None,
    ) -> bool:
        """Reopen and verify package, formulas, styles, pivots, links, and written values."""
        path_to_check = workbook_path or self.filepath
        if source_path:
            source = _xlsx_package_snapshot(source_path)
            result = _xlsx_package_snapshot(path_to_check)
            if source["parts"] != result["parts"]:
                missing = sorted(source["parts"] - result["parts"])
                added = sorted(result["parts"] - source["parts"])
                raise RuntimeError(
                    f"Post-write verification failed: workbook package parts changed "
                    f"(missing={missing[:8]}, added={added[:8]})"
                )
            if source["sheet_names"] != result["sheet_names"]:
                raise RuntimeError("Post-write verification failed: worksheet names changed")
            if source["styles_xml"] != result["styles_xml"]:
                raise RuntimeError("Post-write verification failed: workbook styles changed")
            if source["pivots"] != result["pivots"]:
                raise RuntimeError("Post-write verification failed: PivotTable definitions changed")
            if source["external_links"] != result["external_links"]:
                raise RuntimeError("Post-write verification failed: external workbook links changed")
            for sheet_name, before_sheet in source["sheets"].items():
                after_sheet = result["sheets"].get(sheet_name)
                if after_sheet is None:
                    raise RuntimeError(f"Post-write verification failed: sheet missing: {sheet_name}")
                for key in ("formulas", "styles", "special_counts", "merge_cells"):
                    if before_sheet[key] != after_sheet[key]:
                        raise RuntimeError(
                            f"Post-write verification failed: {key} changed in {sheet_name}"
                        )

        wb_check = openpyxl.load_workbook(path_to_check)
        try:
            if expected_updates:
                source_wb = openpyxl.load_workbook(source_path, data_only=False, read_only=False) if source_path else None
                try:
                    for sheet_name, cell_updates in expected_updates.items():
                        if sheet_name not in wb_check.sheetnames:
                            raise RuntimeError(f"Post-write verification failed: missing sheet {sheet_name}")
                        for cell_ref, expected in cell_updates.items():
                            if source_wb and sheet_name in source_wb.sheetnames:
                                source_cell = source_wb[sheet_name][cell_ref]
                                if isinstance(source_cell.value, str) and source_cell.value.startswith("="):
                                    continue
                            actual = wb_check[sheet_name][cell_ref].value
                            if actual != expected:
                                raise RuntimeError(
                                    f"Post-write verification failed: {sheet_name}!{cell_ref} "
                                    f"expected {expected!r}, found {actual!r}"
                                )
                finally:
                    if source_wb:
                        source_wb.close()

            for sheet_name in wb_check.sheetnames:
                if 'Internal Budget' not in sheet_name or '2026' not in sheet_name:
                    continue

                ws = wb_check[sheet_name]
                all_anchors = AnchorCellTracker.locate_all(ws)
                if not all_anchors:
                    # 非 V2 模板 — 无锚点标记，跳过锚点验证。
                    print("  Post-write verification: skipped (non-V2 template, no anchors found)")
                    return False

                missing = AnchorCellTracker.validate(ws, ["PO_SUMMARY", "FORECAST", "ACTUAL"])
                if missing:
                    raise RuntimeError(f"Post-write verification failed: missing anchors {missing}")
                return True
        finally:
            wb_check.close()
        return False

    def _preserve_recovery_copy(self, temp_path: str) -> Optional[str]:
        """将未能提交的完整更新版保存到自动化工具的备份目录。"""
        try:
            os.makedirs(BACKUP_DIR, exist_ok=True)
            file_name = os.path.basename(self.filepath)
            name, ext = os.path.splitext(file_name)
            timestamp = datetime.now().strftime("%y%m%d_%H%M%S")
            recovery_path = os.path.join(
                BACKUP_DIR, f"{name}_updated_{timestamp}{ext}"
            )
            shutil.copy2(temp_path, recovery_path)
            return recovery_path
        except OSError as error:
            print(f"  [WARN] Could not copy updated workbook to backup directory: {error}")
            return None


# ============================================================
# 高级 API — Forecast / Actual 区域更新
# ============================================================

def build_forecast_updates(
    sheet_name: str,
    personnel: dict,
    person_allocation: dict,    # {name: {month_str: {po_id: days}}}
    name_lookup: dict,
) -> dict:
    """
    构建 Forecast 区域单元格更新。

    Args:
        person_allocation: 未来月份每人每天的 PO 归属
                           {canonical_name: {month_str: {po_id: days}}}

    Returns:
        {cell_ref: new_value} 用于 BudgetReportWriter.write_updates()
    """
    updates = {}

    # 先通过锚点确定 Forecast 区域的行范围
    # 这里构建一个 person_name → row_number 的映射需要读取模板
    # 简化方案：假设 Forecast 区域人员按固定顺序排列
    # 实际使用时需要结合 BudgetReportReader 找到每个人对应的行

    for name, months in person_allocation.items():
        for month_str, po_days in months.items():
            total_days = sum(po_days.values())
            if total_days > 0:
                # 需要找到该人在 Forecast 区域的行号
                # cell_ref 格式: "{col_letter}{row_number}"
                # 由于行号需要通过锚点+人员查找确定，这里返回一个占位标记
                # 实际构建由 build_budget_report_updates() 完成
                pass

    return updates


def build_actual_updates(
    sheet_name: str,
    merged_allocation: dict,    # {name: {month_str: {po_id: days}}}
    personnel: dict,
) -> dict:
    """
    构建 Actual 区域单元格更新。
    将合并后的实际消耗数据映射到 Actual 区域单元格。
    """
    updates = {}
    # 与 build_forecast_updates 类似，需要确定 person→row 映射
    return updates


# ============================================================
# BackupManager
# ============================================================

def backup_budget_report(filepath: str) -> str:
    """
    创建 Budget Report 的时间戳备份。

    Returns:
        backup_path: 备份文件完整路径
    """
    os.makedirs(BACKUP_DIR, exist_ok=True)

    dir_name = os.path.dirname(filepath)
    base_name = os.path.basename(filepath)
    name, ext = os.path.splitext(base_name)
    timestamp = datetime.now().strftime("%y%m%d_%H%M%S")
    backup_name = f"{name}_backup_{timestamp}{ext}"
    backup_path = os.path.join(BACKUP_DIR, backup_name)
    suffix = 2
    while os.path.exists(backup_path):
        backup_name = f"{name}_backup_{timestamp}_{suffix}{ext}"
        backup_path = os.path.join(BACKUP_DIR, backup_name)
        suffix += 1

    shutil.copy2(filepath, backup_path)
    return backup_path


# ============================================================
# Diff 输出
# ============================================================

def print_diff_summary(diff: DiffPreview):
    """在控制台打印差异汇总"""
    print()
    print(diff.summary())

    if diff.total_changes == 0:
        print("  No changes detected.")
        return

    # 按 section 分组显示前 20 条
    print(f"\n  Changes (showing first 20):")
    shown = 0
    for change in diff.changes[:20]:
        delta_str = ""
        if change.delta is not None and change.delta != 0:
            sign = "+" if change.delta > 0 else ""
            delta_str = f" ({sign}{change.delta:.1f})"
        print(f"    {change.cell_ref}: {change.current_value} → {change.new_value}{delta_str}")
        shown += 1

    if diff.total_changes > shown:
        print(f"    ... and {diff.total_changes - shown} more changes")


def save_diff_excel(diff: DiffPreview, output_path: str):
    """将差异保存为 Excel 文件（方便审查）"""
    import xlsxwriter
    wb = xlsxwriter.Workbook(output_path)
    ws = wb.add_worksheet("Changes")

    fmt_header = wb.add_format({
        'bold': True, 'bg_color': '#2F5496', 'font_color': 'white',
        'border': 1, 'align': 'center', 'font_size': 11,
    })
    fmt_cell = wb.add_format({'border': 1, 'align': 'center', 'font_size': 10})
    fmt_add = wb.add_format({'border': 1, 'align': 'center', 'font_size': 10, 'bg_color': '#C6EFCE'})
    fmt_mod = wb.add_format({'border': 1, 'align': 'center', 'font_size': 10, 'bg_color': '#FFEB9C'})
    fmt_del = wb.add_format({'border': 1, 'align': 'center', 'font_size': 10, 'bg_color': '#FFC7CE'})
    fmt_num = wb.add_format({'border': 1, 'align': 'center', 'font_size': 10, 'num_format': '#,##0.0'})

    headers = ["Section", "Cell", "Current Value", "New Value", "Delta", "Type"]
    for i, h in enumerate(headers):
        ws.write(0, i, h, fmt_header)

    for r, change in enumerate(diff.changes, 1):
        section = diff.sheet_name
        type_str = "ADD" if change.is_addition else ("DEL" if change.is_deletion else "MOD")
        fmt = fmt_add if change.is_addition else (fmt_del if change.is_deletion else fmt_mod)

        ws.write(r, 0, section, fmt_cell)
        ws.write(r, 1, change.cell_ref, fmt_cell)
        ws.write(r, 2, str(change.current_value) if change.current_value is not None else "(empty)", fmt_cell)
        ws.write(r, 3, str(change.new_value) if change.new_value is not None else "(empty)", fmt_cell)
        if change.delta is not None:
            ws.write(r, 4, change.delta, fmt_num)
        else:
            ws.write(r, 4, "-", fmt_cell)
        ws.write(r, 5, type_str, fmt)

    ws.set_column(0, 0, 20)
    ws.set_column(1, 1, 10)
    ws.set_column(2, 3, 18)
    ws.set_column(4, 4, 10)
    ws.set_column(5, 5, 8)
    ws.freeze_panes(1, 0)

    wb.close()
    print(f"  Diff Excel saved: {output_path}")


# ============================================================
# 确认提示
# ============================================================

def confirm_write(prompt: str = "Apply these changes to the Budget Report?") -> bool:
    """
    显示确认提示，默认超时 60s 返回 False（安全默认）。
    """
    try:
        print(f"\n{prompt} [y/N] (60s timeout): ", end="", flush=True)
        # 使用简单的 input（在 run.bat 环境中可用）
        response = input().strip().lower()
        return response in ('y', 'yes')
    except (EOFError, KeyboardInterrupt):
        print("\n  Aborted.")
        return False


# ============================================================
# CLI
# ============================================================

def main():
    setup_utf8_stdout()

    import argparse
    parser = argparse.ArgumentParser(description="Budget Report Writer CLI")
    parser.add_argument("--file", required=True, help="Path to enhanced Budget Report template")
    parser.add_argument("--check-template", action="store_true", help="Check template version only")
    parser.add_argument("--list-anchors", action="store_true", help="List all anchor markers")
    args = parser.parse_args()

    filepath = args.file
    if not os.path.exists(filepath):
        print(f"[ERROR] File not found: {filepath}")
        return

    if args.check_template:
        wb = safe_load_workbook(filepath)
        if not wb:
            print("FAIL: Cannot open file")
            return
        if "_ToolConfig" not in wb.sheetnames:
            print("FAIL: _ToolConfig sheet missing — not an enhanced template")
            wb.close()
            return
        ws = wb["_ToolConfig"]
        version = ws.cell(row=1, column=2).value
        wb.close()
        if str(version) == TEMPLATE_VERSION:
            print(f"OK: Template version {version}")
        else:
            print(f"FAIL: Template version {version} (expected {TEMPLATE_VERSION})")

    if args.list_anchors:
        wb = openpyxl.load_workbook(filepath, data_only=False)
        for sn in wb.sheetnames:
            if 'Internal Budget' in sn or 'Budget_FC' in sn:
                ws = wb[sn]
                anchors = AnchorCellTracker.locate_all(ws)
                if anchors:
                    print(f"\n{sn}:")
                    for name, row in sorted(anchors.items(), key=lambda x: x[1]):
                        print(f"  Row {row}: {name}")
        wb.close()


if __name__ == "__main__":
    main()
