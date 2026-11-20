# -*- coding: utf-8 -*-
"""
Template Enhancer — 将现有 Budget Report 升级为工具可读写的 V2 模板

添加:
1. Column A 锚点标记（隐藏文字，标记各数据区域边界）
2. _ToolConfig 隐藏 sheet（模板版本号）
3. PO_Config sheet（结构化 PO 配置，从 po_config_v2.xlsx 同步）

用法:
    python template_enhancer.py --project CNTS-25165
    python template_enhancer.py --project CN03-22052
"""

import os
import sys
import shutil
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config
from utils import setup_utf8_stdout, safe_load_workbook, find_latest_file


# ============================================================
# 锚点配置 — 每个项目的 section 边界定义
# ============================================================

# CNTS-25165: Internal Budget_FC 2026 Q4 sheet 的行布局
CNTS25165_ANCHORS = {
    "Internal Budget_FC 2026 Q4": {
        "PO_SUMMARY_START": 3,
        "PO_SUMMARY_END": 7,
        "FORECAST_START": 9,
        "FORECAST_END": 30,
        "ACTUAL_START": 33,
        "ACTUAL_END": 54,
    },
    "Internal Budget_FC 2026 Q2": {
        "PO_SUMMARY_START": 3,
        "PO_SUMMARY_END": 7,
        "FORECAST_START": 9,
        "FORECAST_END": 30,
        "ACTUAL_START": 33,
        "ACTUAL_END": 54,
    },
}

# CN03-22052: Budget_FC 2026 sheet
CN0322052_ANCHORS = {
    "Budget_FC 2026": {
        # Driving section (-10000)
        "PO_SUMMARY_10000_START": 4,
        "PO_SUMMARY_10000_END": 10,
        "FORECAST_10000_START": 12,
        "FORECAST_10000_END": 20,
        "ACTUAL_10000_START": 22,
        "ACTUAL_10000_END": 28,
        # Guard section (-11000)
        "PO_SUMMARY_11000_START": 32,
        "PO_SUMMARY_11000_END": 39,
        "FORECAST_11000_START": 41,
        "FORECAST_11000_END": 49,
        "ACTUAL_11000_START": 50,
        "ACTUAL_11000_END": 54,
    },
}

# 锚点文字标记
ANCHOR_PREFIX = "=== "
ANCHOR_SUFFIX = " ==="

# 锚点文字颜色（白色，与背景融为一体）
ANCHOR_FONT_COLOR = "FFFFFF"


def get_project_anchors(project: str, sheet_name: str) -> dict:
    """获取指定项目和 sheet 的锚点配置"""
    if project == "CNTS-25165":
        return CNTS25165_ANCHORS.get(sheet_name, {})
    elif project == "CN03-22052":
        return CN0322052_ANCHORS.get(sheet_name, {})
    return {}


def add_anchors_to_sheet(ws, anchors: dict):
    """在 sheet 的 Column A 添加锚点标记"""
    from openpyxl.styles import Font

    white_font = Font(color=ANCHOR_FONT_COLOR, sz=8)

    for section_name, row_num in anchors.items():
        cell = ws.cell(row=row_num, column=1)
        cell.value = f"{ANCHOR_PREFIX}{section_name}{ANCHOR_SUFFIX}"
        cell.font = white_font

    print(f"  Added {len(anchors)} anchor markers to column A")


def add_tool_config_sheet(wb):
    """添加隐藏的 _ToolConfig sheet"""
    if "_ToolConfig" in wb.sheetnames:
        del wb["_ToolConfig"]

    ws = wb.create_sheet("_ToolConfig")
    ws.sheet_state = "hidden"

    ws.cell(row=1, column=1, value="TEMPLATE_VERSION")
    ws.cell(row=1, column=2, value="2.0")

    ws.cell(row=2, column=1, value="CREATED_DATE")
    ws.cell(row=2, column=2, value=datetime.now().strftime("%Y-%m-%d %H:%M"))

    ws.cell(row=3, column=1, value="TOOL")
    ws.cell(row=3, column=2, value="po_allocator_v2.py")

    print("  Added _ToolConfig hidden sheet (TEMPLATE_VERSION=2.0)")


def add_po_config_sheet(wb, project: str):
    """从 input/po_config_v2.xlsx 读取 PO 配置，写入 PO_Config sheet"""
    import openpyxl

    if "PO_Config" in wb.sheetnames:
        del wb["PO_Config"]

    ws = wb.create_sheet("PO_Config")

    # 表头
    config_headers = [
        "Project", "PO_ID", "Level", "Days_Quota", "Unit_Price",
        "Value_RMB", "Start_Month", "End_Month", "Priority", "Allocation_Mode"
    ]

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    header_font = Font(bold=True, color="FFFFFF", sz=11)
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    for i, h in enumerate(config_headers, 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    # 读取 po_config_v2.xlsx 中的数据
    config_path = os.path.join(config.INPUT_DIR, "po_config_v2.xlsx")
    row_num = 2

    if os.path.exists(config_path):
        try:
            cfg_wb = openpyxl.load_workbook(config_path, data_only=True)
            if "PO_Level_Config" in cfg_wb.sheetnames:
                cfg_ws = cfg_wb["PO_Level_Config"]
                headers, rows_data = [], []
                for r_idx, row in enumerate(
                    cfg_ws.iter_rows(min_row=1, max_row=cfg_ws.max_row, values_only=True)
                ):
                    if r_idx == 0:
                        headers = [str(h) if h else "" for h in row]
                    else:
                        rows_data.append(list(row))

                for row_data in rows_data:
                    if len(row_data) < 6:
                        continue
                    # 检查 project 筛选
                    row_project = str(row_data[0]).strip() if row_data[0] else ""
                    if project and row_project and row_project != project:
                        continue

                    # 写入数据
                    for c_idx, val in enumerate(row_data):
                        cell = ws.cell(row=row_num, column=c_idx + 1, value=val)
                        cell.border = thin_border
                    row_num += 1

            cfg_wb.close()
        except Exception as e:
            print(f"  [WARN] Failed to read po_config_v2.xlsx: {e}")

    # 如果没有从 Excel 读到数据，写入硬编码默认值
    if row_num == 2 and project == "CNTS-25165":
        default_configs = [
            ["CNTS-25165", "2601000336", "Testing Engineer-Expert", 132, 2460, 324720, "2026-01", "2026-04", 1, "usage"],
            ["CNTS-25165", "2601000337", "Testing Engineer-Expert", 132, 2460, 324720, "2026-01", "2026-04", 1, "usage"],
            ["CNTS-25165", "2601000337", "Testing Engineer-Senior", 132, 2260, 298320, "2026-01", "2026-04", 1, "usage"],
            ["CNTS-25165", "2601000338", "Testing Engineer-Expert", 66, 2460, 162360, "2026-01", "2026-04", 1, "usage"],
            ["CNTS-25165", "2601000338", "Testing Engineer-Experienced", 330, 2100, 693000, "2026-01", "2026-04", 1, "usage"],
            ["CNTS-25165", "2604000799", "Testing Engineer-Experienced", 350, 2100, 735000, "2026-04", "2026-09", 2, "usage"],
            ["CNTS-25165", "2604000799", "Testing Engineer-Senior", 280, 2260, 632800, "2026-04", "2026-09", 2, "usage"],
            ["CNTS-25165", "2604000799", "Testing Engineer-Expert", 210, 2460, 516600, "2026-04", "2026-09", 2, "usage"],
            ["CNTS-25165", "2607000312", "Testing Engineer-Experienced", 180, 2100, 378000, "2026-07", "2026-12", 3, "usage"],
            ["CNTS-25165", "2607000312", "Testing Engineer-Senior", 85, 2260, 192100, "2026-07", "2026-12", 3, "usage"],
            ["CNTS-25165", "2607000312", "Testing Engineer-Expert", 464, 2460, 1141440, "2026-07", "2026-12", 3, "usage"],
        ]
        for cfg_row in default_configs:
            for c_idx, val in enumerate(cfg_row):
                cell = ws.cell(row=row_num, column=c_idx + 1, value=val)
                cell.border = thin_border
            row_num += 1

    # 列宽
    col_widths = [14, 18, 30, 12, 12, 14, 12, 12, 10, 16]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(config_headers))}{row_num - 1}"

    print(f"  Added PO_Config sheet with {row_num - 2} entries for {project}")


def enhance_template(src_path: str, project: str) -> str:
    """
    从源文件创建增强版模板。

    返回增强模板的路径（保存在同目录，加 _V2 后缀）。
    """
    import openpyxl

    if not os.path.exists(src_path):
        raise FileNotFoundError(f"Source Budget Report not found: {src_path}")

    # 生成输出路径
    dir_name = os.path.dirname(src_path)
    base_name = os.path.basename(src_path)
    name, ext = os.path.splitext(base_name)
    # 去掉已有的日期后缀，替换为 TEMPLATE
    if "_V2" not in name:
        out_name = name.rsplit("_", 1)[0] + "_TEMPLATE_V2" + ext
    else:
        out_name = base_name
    out_path = os.path.join(dir_name, out_name)

    print(f"\nEnhancing template for {project}:")
    print(f"  Source: {src_path}")
    print(f"  Output: {out_path}")

    # 打开源文件（保留公式和格式）
    wb = openpyxl.load_workbook(src_path)

    # Step 1: 在每个 Budget_FC sheet 中添加 Column A 锚点
    for sheet_name in wb.sheetnames:
        anchors = get_project_anchors(project, sheet_name)
        if anchors:
            ws = wb[sheet_name]
            add_anchors_to_sheet(ws, anchors)
            print(f"  Anchored sheet: {sheet_name}")

    # Step 2: 添加 _ToolConfig 隐藏 sheet
    add_tool_config_sheet(wb)

    # Step 3: 添加 PO_Config sheet
    add_po_config_sheet(wb, project)

    # 保存
    wb.save(out_path)
    wb.close()

    print(f"\n  Template saved: {out_path}")
    return out_path


def main():
    setup_utf8_stdout()

    import argparse
    parser = argparse.ArgumentParser(description="Enhance Budget Report template for tool write-back")
    parser.add_argument("--project", default="CNTS-25165", help="Project code")
    parser.add_argument("--source", default=None, help="Source Budget Report path (auto-detected if omitted)")
    args = parser.parse_args()

    project = args.project

    if args.source:
        src_path = args.source
    else:
        # 自动查找最新的 Budget Report
        proj_config = config.PROJECTS.get(project, {})
        budget_dir = proj_config.get("budget_report_dir", "")
        if not budget_dir or not os.path.isdir(budget_dir):
            print(f"[ERROR] Budget Report directory not found for project: {project}")
            return

        # 查找 IAV_Budget_Report_*.xlsx
        import glob
        pattern = os.path.join(budget_dir, "IAV_Budget_Report_*.xlsx")
        files = glob.glob(pattern)
        if not files:
            print(f"[ERROR] No Budget Report found matching: {pattern}")
            return

        src_path = max(files, key=os.path.getmtime)

    try:
        out_path = enhance_template(src_path, project)
        print(f"\nDone! Enhanced template: {out_path}")
        print(f"Next step: Update config.py BUDGET_REPORT_TEMPLATE_PATHS to point to this file.")
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
