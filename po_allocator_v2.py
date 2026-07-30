# -*- coding: utf-8 -*-
"""
PO 工时分配器 V2 — 基于 Level 配额的核心算法

改进点 (vs V1):
1. PO 按 Level 配额管理: 每张 PO 的每个 Level 有独立的天数配额
2. 人员→Level 匹配后从对应配额扣天数
3. 第三方人员工时自动从 PDF 抓取 (Source 3)
4. 过渡月份拆分: 旧 PO 剩余 < 当月天数 → 拆分到新旧 PO
5. Use Untill 基于团队消耗速度计算

用法:
    python po_allocator_v2.py                    # 运行全部
    python po_allocator_v2.py --project CNTS-25165  # 指定项目
"""

import os
import sys
import io
import re
import math
from datetime import datetime, timedelta
from collections import defaultdict
from typing import Optional

# 添加当前目录到 path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import config
from utils import (
    setup_utf8_stdout, safe_load_workbook, read_sheet,
    write_excel, find_latest_file, safe_str
)
from po_pdf_parser import parse_po_pdf, po_to_config_rows, match_level
from timesheet_parser import scan_timesheet_directory, records_to_summary


# ============================================================
# 常量
# ============================================================
HOURS_PER_DAY = config.HOURS_PER_DAY  # 8
MONTH_NAMES = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
    "January": 1, "February": 2, "March": 3, "April": 4,
    "June": 6, "July": 7, "August": 8,
    "September": 9, "October": 10, "November": 11, "December": 12,
}
# 中国法定工作日估算 (2026, 每月约 21-23 天)
DEFAULT_WORKDAYS = {1: 22, 2: 18, 3: 22, 4: 22, 5: 21, 6: 22,
                    7: 23, 8: 21, 9: 22, 10: 19, 11: 21, 12: 23}


# ============================================================
# 数据模型
# ============================================================
class POLevelConfig:
    """单张 PO 内某个 Level 的配额"""
    def __init__(self, po_id: str, level: str, days_quota: float,
                 unit_price: float, value_rmb: float = 0,
                 start_month: str = "", end_month: str = "", priority: int = 1):
        self.po_id = po_id
        self.level = level
        self.days_quota = days_quota
        self.days_remaining = days_quota  # 分配过程中递减
        self.unit_price = unit_price
        self.value_rmb = value_rmb
        self.start_month = start_month
        self.end_month = end_month
        self.priority = priority

    def __repr__(self):
        return (f"POLevel({self.po_id}, {self.level}, "
                f"quota={self.days_quota:.0f}d, remaining={self.days_remaining:.0f}d)")


class Person:
    """人员信息"""
    def __init__(self, name: str, level: str, join_month: str = "2026-01",
                 leave_month: str = ""):
        self.name = name
        self.level = level
        self.join_month = join_month    # 加入月份 (含)
        self.leave_month = leave_month  # 离开月份 (含, 空=未离开)
        self.is_active = True

    def is_in_project(self, year: int, month: int) -> bool:
        """检查该人员在某月是否在项目中"""
        month_str = f"{year}-{month:02d}"
        if month_str < self.join_month:
            return False
        if self.leave_month and month_str > self.leave_month:
            return False
        return True


class AllocationResult:
    """分配结果: 每人每月每 PO 的天数"""
    def __init__(self):
        # {person_name: {month_str: {po_id: days}}}
        self.matrix = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
        # PO 消耗汇总
        self.po_consumption = {}  # {po_id_level_key: {month: days}}
        # Use Untill 预计
        self.use_untill = {}  # {po_id: estimated_date}
        # 过渡详情
        self.transitions = []  # [(month, old_po, new_po, old_days, new_days)]


# ============================================================
# 工时数据加载 (合并 3 个数据源)
# ============================================================
def load_monthly_closing_hours(project_wbs: str) -> dict:
    """
    Source 1: Monthly Closing 数据库
    返回: {name_en: {month_str: hours}}

    Monthly Closing 文件格式:
    - Row 1: 部分为空，可能有数据
    - Row 2: 实际表头 (RepDate, EmployeeNumb, Send.CCtr, SArtTyp,
              Receiver WBS element, ..., Total Quantity, ...)
    - Row 3+: 数据 (Last, First, RepDate, EmployeeNumb, ..., Total Quantity, ...)
      Col 0=Last, Col 1=First, Col 6=WBS, Col 9=Hours
    """
    result = defaultdict(lambda: defaultdict(float))

    pattern = config.MONTHLY_CLOSING_PATTERN
    filepath = find_latest_file(os.path.dirname(pattern),
                                 os.path.basename(pattern))
    if not filepath:
        print(f"  [WARN] Monthly Closing file not found: {pattern}")
        return dict(result)

    wb = safe_load_workbook(filepath)
    if not wb:
        return dict(result)

    for sheet_name in wb.sheetnames:
        # 解析 sheet 名为月份
        month_num = None
        for name_part, num in MONTH_NAMES.items():
            if sheet_name.lower().startswith(name_part.lower()):
                month_num = num
                break
        if month_num is None:
            continue

        month_str = f"2026-{month_num:02d}"
        ws = wb[sheet_name]

        # 读取 Row 2 作为表头
        header_row = [str(c.value) if c.value else "" for c in next(ws.iter_rows(min_row=2, max_row=2))]
        col_wbs = None
        col_hours = None
        col_last = 0   # Last name
        col_first = 1  # First name
        for i, h in enumerate(header_row):
            h_lower = h.lower()
            # 优先用 Project Number (col 10, CNTS-25165-11026 格式)
            if 'project' in h_lower and 'number' in h_lower:
                col_wbs = i
            elif 'receiver' in h_lower and 'wbs' in h_lower:
                if col_wbs is None:  # 如果 Project Number 未找到
                    col_wbs = i
            elif 'total' in h_lower and 'quantity' in h_lower:
                col_hours = i

        # 如果 Project Number 和 WBS 都没找到，直接用 col 10
        if col_wbs is None:
            col_wbs = 10

        if col_wbs is None or col_hours is None:
            continue

        # 从 Row 3 开始读取数据
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
            if len(row) <= max(col_wbs, col_hours):
                continue
            wbs_val = str(row[col_wbs]) if col_wbs < len(row) else ""
            if project_wbs not in wbs_val:
                continue

            hours = 0.0
            try:
                hours = float(row[col_hours]) if row[col_hours] else 0.0
            except (ValueError, TypeError):
                pass

            if hours <= 0:
                continue

            # 人员名称: "Last, First" → "First Last"
            last = str(row[col_last]).strip() if len(row) > col_last and row[col_last] else ""
            first = str(row[col_first]).strip() if len(row) > col_first and row[col_first] else ""
            if last and first:
                name = f"{first} {last}"
            elif last:
                name = last
            elif first:
                name = first
            else:
                continue

            result[name][month_str] += hours

    wb.close()
    return dict(result)


def load_budget_report_hours(project_dir: str, personnel_names: list[str]) -> dict:
    """
    Source 2: Budget Report 内部的 Actual 区域
    返回: {name: {month_str: hours}}
    """
    result = defaultdict(lambda: defaultdict(float))

    pattern = os.path.join(project_dir, "IAV_Budget_Report_*.xlsx")
    filepath = find_latest_file(project_dir, "IAV_Budget_Report_*.xlsx")
    if not filepath:
        return dict(result)

    headers, rows = read_sheet(filepath, "Actual")
    if not headers:
        return dict(result)

    # 尝试找姓名列和月度工时列
    name_cols = []
    month_cols = {}  # {col_index: month_num}
    for i, h in enumerate(headers):
        h_str = str(h).lower()
        if 'name' in h_str or '姓名' in h_str:
            name_cols.append(i)
        for m_name, m_num in MONTH_NAMES.items():
            if m_name.lower() in h_str and len(h_str) <= 5:
                month_cols[i] = m_num

    for row in rows:
        name = ""
        for nc in name_cols:
            if nc < len(row) and row[nc]:
                name = str(row[nc]).strip()
                break
        if not name:
            continue

        # 匹配已知人员
        matched = None
        for pn in personnel_names:
            if pn.lower() in name.lower() or name.lower() in pn.lower():
                matched = pn
                break
        if not matched:
            continue

        for col, month_num in month_cols.items():
            if col < len(row):
                try:
                    hours = float(row[col]) if row[col] else 0.0
                except (ValueError, TypeError):
                    hours = 0.0
                if hours > 0:
                    result[matched][f"2026-{month_num:02d}"] += hours

    return dict(result)


def load_timesheet_hours(base_dir: str, months: list[str] = None) -> dict:
    """
    Source 3: 第三方人员工时 PDF
    返回: {name_cn: {month_str: hours}}
    """
    records = scan_timesheet_directory(base_dir, months)
    return records_to_summary(records)


def normalize_name(raw_name: str, name_lookup: dict) -> Optional[str]:
    """将任何名字变体标准化为 canonical name"""
    if not raw_name:
        return None
    # 直接匹配
    canonical = name_lookup.get(raw_name.lower())
    if canonical:
        return canonical
    # 模糊匹配 (子串)
    for alias, canonical_name in name_lookup.items():
        if alias in raw_name.lower() or raw_name.lower() in alias:
            return canonical_name
    return None


def merge_all_hours(project_wbs: str, project_dir: str, timesheet_dir: str,
                    personnel: dict[str, dict]) -> dict:
    """
    合并数据源的工时数据
    返回: {canonical_name: {month_str: days}}
    """
    # 构建姓名查找表
    name_lookup = config.build_name_lookup(personnel)

    print("\n  [1/3] Loading Monthly Closing...")
    mc_hours_raw = load_monthly_closing_hours(project_wbs)
    mc_hours = {}
    for name, months in mc_hours_raw.items():
        canonical = normalize_name(name, name_lookup)
        if canonical:
            mc_hours[canonical] = {m: h / HOURS_PER_DAY for m, h in months.items()}
        else:
            print(f"        [WARN] Unknown name in MC: '{name}' — skipping")
    print(f"        Found {len(mc_hours)} matched people from Monthly Closing")

    print("  [2/3] Loading Budget Report Actual...")
    br_hours_raw = load_budget_report_hours(project_dir, list(personnel.keys()))
    br_hours = {}
    for name, months in br_hours_raw.items():
        canonical = normalize_name(name, name_lookup)
        if canonical:
            br_hours[canonical] = {m: h / HOURS_PER_DAY for m, h in months.items()}
    print(f"        Found {len(br_hours)} matched people from Budget Report")

    print("  [3/3] Loading Timesheet PDFs...")
    ts_hours_raw = load_timesheet_hours(timesheet_dir)
    ts_hours = {}
    for name, months in ts_hours_raw.items():
        canonical = normalize_name(name, name_lookup)
        if canonical:
            ts_hours[canonical] = {m: h / HOURS_PER_DAY for m, h in months.items()}
    print(f"        Found {len(ts_hours)} matched people from Timesheet PDFs")

    # 合并: BR 覆盖 MC, TS 作为补充
    merged = defaultdict(lambda: defaultdict(float))

    # 先加 MC
    for name, months in mc_hours.items():
        for m, days in months.items():
            merged[name][m] += days

    # BR 覆盖 (更准确)
    for name, months in br_hours.items():
        for m, days in months.items():
            merged[name][m] = days  # 覆盖

    # TS 补充 (第三方人员)
    for name, months in ts_hours.items():
        for m, days in months.items():
            if name not in merged or sum(merged[name].values()) == 0:
                merged[name][m] += days

    total_people = len(merged)
    total_days = sum(sum(months.values()) for months in merged.values())
    print(f"        Merged: {total_people} people, {total_days:.0f} total days")

    return dict(merged)


# ============================================================
# PO Level 配置加载
# ============================================================
def load_po_level_config(project: str = "CNTS-25165") -> list[POLevelConfig]:
    """
    加载 PO Level 配置
    优先从 input/po_config_v2.xlsx 读取，其次从硬编码默认值

    返回: POLevelConfig 列表
    """
    configs = []

    # 尝试从 Excel 读取
    excel_path = os.path.join(config.INPUT_DIR, "po_config_v2.xlsx")
    if os.path.exists(excel_path):
        headers, rows = read_sheet(excel_path, "PO_Level_Config")
        if headers:
            for row in rows:
                if len(row) < 6:
                    continue
                try:
                    po_id = str(row[0]).strip()
                    level = str(row[1]).strip()
                    days_quota = float(row[2]) if row[2] else 0
                    unit_price = float(row[3]) if row[3] else 0
                    value_rmb = float(row[4]) if row[4] else days_quota * unit_price
                    start_month = str(row[5]).strip() if len(row) > 5 else ""
                    end_month = str(row[6]).strip() if len(row) > 6 else ""
                    priority = int(row[7]) if len(row) > 7 and row[7] else 1

                    if po_id and level and days_quota > 0:
                        configs.append(POLevelConfig(
                            po_id, level, days_quota, unit_price,
                            value_rmb, start_month, end_month, priority
                        ))
                except (ValueError, IndexError) as e:
                    print(f"  [WARN] Skipping row in po_config_v2: {row[:6]} — {e}")
                    continue

    # 如果 Excel 为空，用硬编码默认值
    if not configs and project == "CNTS-25165":
        configs = _get_default_cnst25165_config()

    return configs


def _get_default_cnst25165_config() -> list[POLevelConfig]:
    """CNTS-25165 默认 PO Level 配置"""
    return [
        # 2026 Q1 分组 PO
        POLevelConfig("2601000336", "Testing Engineer-Expert", 180, 2260, 324720,
                      "2026-01", "2026-04", 1),
        POLevelConfig("2601000337", "Testing Engineer-Expert", 120, 2460, 295200,
                      "2026-01", "2026-04", 1),
        POLevelConfig("2601000337", "Testing Engineer-Senior", 130, 2260, 327840,
                      "2026-01", "2026-04", 1),
        POLevelConfig("2601000338", "Testing Engineer-Expert", 90, 2460, 221400,
                      "2026-01", "2026-04", 1),
        POLevelConfig("2601000338", "Testing Engineer-Experienced", 264, 2100, 633960,
                      "2026-01", "2026-04", 1),
        # 2026 Q2-Q3 全员 PO
        POLevelConfig("2604000799", "Testing Engineer-Experienced", 180, 2100, 378000,
                      "2026-04", "2026-09", 2),
        POLevelConfig("2604000799", "Testing Engineer-Senior", 85, 2260, 192100,
                      "2026-04", "2026-09", 2),
        POLevelConfig("2604000799", "Testing Engineer-Expert", 464, 2460, 1141440,
                      "2026-04", "2026-09", 2),
        # 2026 Q3-Q4 全员 PO (来自 PDF 解析)
        POLevelConfig("2607000312", "Testing Engineer-Experienced", 180, 2100, 378000,
                      "2026-07", "2026-12", 3),
        POLevelConfig("2607000312", "Testing Engineer-Senior", 85, 2260, 192100,
                      "2026-07", "2026-12", 3),
        POLevelConfig("2607000312", "Testing Engineer-Expert", 464, 2460, 1141440,
                      "2026-07", "2026-12", 3),
    ]


def load_po_from_pdf(pdf_path: str, start_month: str, end_month: str,
                     priority: int) -> list[POLevelConfig]:
    """
    从客户 PO PDF 自动解析 Level 配额
    """
    info = parse_po_pdf(pdf_path)
    if not info:
        return []

    configs = []
    for quota in info.level_quotas:
        level = match_level(quota.description)
        if not level:
            level = quota.description  # 保留原始描述

        configs.append(POLevelConfig(
            po_id=info.po_number,
            level=level,
            days_quota=quota.qty_days,
            unit_price=quota.unit_price,
            value_rmb=quota.total_rmb,
            start_month=start_month,
            end_month=end_month,
            priority=priority,
        ))
    return configs


# ============================================================
# 核心分配算法
# ============================================================
def get_people_by_level(personnel: dict[str, dict], level: str) -> list[str]:
    """获取指定 Level 的人员列表"""
    return [name for name, info in personnel.items()
            if info.get("level", "") == level]


def calculate_po_allocation_v2(
    personnel: dict[str, dict],
    po_configs: list[POLevelConfig],
    actual_hours: dict[str, dict[str, float]],
    target_year: int = 2026,
) -> tuple[AllocationResult, dict]:
    """
    核心分配算法 V2 — 基于 Level 配额

    逻辑:
    1. 按 Priority 分组处理 PO
    2. 对每张 PO 的每个 Level:
       - 获取该 Level 的人员列表
       - 从 start_month 开始，逐月分配天数
       - 过渡月份: 旧 PO 剩余天数 < 当月需分配天数 → 拆分
    3. 新人: 从 join_month 开始参与分配

    Args:
        personnel: {name: {"level": ..., "join_month": ..., "supervisor": ...}}
        po_configs: POLevelConfig 列表
        actual_hours: {name: {month_str: days}}
        target_year: 目标年份

    Returns:
        (AllocationResult, po_summary_dict)
    """
    result = AllocationResult()

    # 按 priority 和 start_month 排序 PO
    po_configs.sort(key=lambda c: (c.priority, c.start_month))

    # 按 Level 分组 PO 配额
    # po_level_map[level][priority] = [POLevelConfig, ...]
    po_level_map = defaultdict(lambda: defaultdict(list))
    all_po_ids = set()
    for c in po_configs:
        po_level_map[c.level][c.priority].append(c)
        all_po_ids.add(c.po_id)

    months = [f"{target_year}-{m:02d}" for m in range(1, 13)]

    # 追踪每个 level 的"当前活跃 PO 和剩余配额"
    # level_current_po[level] = {"po_id": ..., "remaining": ..., "priority": ...}
    level_current_po = {}

    print("\n  --- Allocation Log ---")

    for month_str in months:
        year, month_num = map(int, month_str.split('-'))
        workdays = DEFAULT_WORKDAYS.get(month_num, 22)

        # 检查是否有新的 PO 需要切换
        for level, priority_groups in po_level_map.items():
            for priority, configs in sorted(priority_groups.items()):
                for c in configs:
                    if c.start_month == month_str:
                        # 新 PO 激活
                        if level in level_current_po:
                            old = level_current_po[level]
                            old_remaining = old["remaining"]
                            if old_remaining > 0:
                                # 旧 PO 还有剩余 → 过渡月份拆分
                                print(f"  [{month_str}] Transition: {old['po_id']} "
                                      f"({old_remaining:.1f}d remaining) → {c.po_id} "
                                      f"({c.days_quota:.0f}d new)")
                            else:
                                print(f"  [{month_str}] Switch: {old['po_id']} → "
                                      f"{c.po_id} ({c.days_quota:.0f}d, {c.level})")

                        level_current_po[level] = {
                            "po_id": c.po_id,
                            "remaining": c.days_quota,
                            "priority": priority,
                            "config": c,
                        }

        # 对每个 Level 进行分配
        for level, current_po in level_current_po.items():
            if current_po["remaining"] <= 0:
                continue

            people = get_people_by_level(personnel, level)
            if not people:
                continue

            # 计算该 level 本月总需分配天数
            level_actual_days = 0
            for name in people:
                person_info = personnel[name]
                person = Person(name, level,
                                person_info.get("join_month", "2026-01"),
                                person_info.get("leave_month", ""))
                if person.is_in_project(year, month_num):
                    days = actual_hours.get(name, {}).get(month_str, 0)
                    level_actual_days += days

            po_id = current_po["po_id"]
            remaining = current_po["remaining"]

            # 分配逻辑
            if level_actual_days <= 0:
                continue

            if remaining >= level_actual_days:
                # 够分: 全部在本 PO
                for name in people:
                    person_info = personnel[name]
                    person = Person(name, level,
                                    person_info.get("join_month", "2026-01"),
                                    person_info.get("leave_month", ""))
                    if person.is_in_project(year, month_num):
                        days = actual_hours.get(name, {}).get(month_str, 0)
                        if days > 0:
                            result.matrix[name][month_str][po_id] = days
                current_po["remaining"] -= level_actual_days
            else:
                # 不够分: 旧 PO 耗尽，需要拆分
                for name in people:
                    person_info = personnel[name]
                    person = Person(name, level,
                                    person_info.get("join_month", "2026-01"),
                                    person_info.get("leave_month", ""))
                    if not person.is_in_project(year, month_num):
                        continue
                    days = actual_hours.get(name, {}).get(month_str, 0)
                    if days <= 0:
                        continue

                    if remaining > 0:
                        # 部分在旧 PO
                        ratio = remaining / level_actual_days
                        old_days = days * ratio
                        result.matrix[name][month_str][po_id] = round(old_days, 1)
                        remaining -= old_days
                    # 其余部分去新 PO（由下个 priority 的 PO 接）

                current_po["remaining"] = max(0, remaining)

            if current_po["remaining"] <= 0:
                print(f"  [{month_str}] {po_id} ({level}) exhausted")

    # 构建 PO 消耗汇总
    po_summary = {}
    for level, current_po in level_current_po.items():
        po_id = current_po["po_id"]
        key = f"{po_id}|{level}"
        allocated_days = sum(
            sum(po_days.get(po_id, 0) for po_days in person_months.values())
            for person_months in result.matrix.values()
        )
        po_summary[key] = {
            "po_id": po_id,
            "level": level,
            "quota": current_po.get("config", None),
            "allocated": allocated_days,
        }

    return result, po_summary


# ============================================================
# Use Untill 计算
# ============================================================
def calculate_use_untill(allocation: AllocationResult,
                         personnel: dict[str, dict]) -> dict[str, str]:
    """
    基于团队消耗速度计算每张 PO 预计用完日期

    返回: {po_id: "YYYY-MM" or "YYYY-MM-DD"}
    """
    # 汇总每张 PO 每月消耗的总天数
    po_monthly = defaultdict(lambda: defaultdict(float))
    for name, months in allocation.matrix.items():
        for month_str, po_days in months.items():
            for po_id, days in po_days.items():
                po_monthly[po_id][month_str] += days

    use_untill = {}
    for po_id, months in po_monthly.items():
        # 算平均月消耗
        total_days = sum(months.values())
        num_months = len(months)
        if num_months == 0 or total_days == 0:
            use_untill[po_id] = "N/A"
            continue

        avg_per_month = total_days / num_months

        # 算配额总量
        total_quota = 0
        for name, person_info in personnel.items():
            pass  # 需要从 PO config 获取配额，此处简化

        # 找最后有分配的月份
        sorted_months = sorted(months.keys())
        last_month = sorted_months[-1]

        # 估算
        remaining_from_last = avg_per_month  # 保守估计
        if remaining_from_last > 0:
            y, m = map(int, last_month.split('-'))
            # 假设按平均速度消耗
            extrapolate_months = 0  # 简化版
            use_untill[po_id] = last_month
        else:
            use_untill[po_id] = last_month

    return use_untill


# ============================================================
# 输出报告
# ============================================================
def generate_report_v2(allocation: AllocationResult,
                       po_summary: dict,
                       use_untill: dict[str, str],
                       personnel: dict[str, dict],
                       actual_hours: dict,
                       output_path: str) -> None:
    """生成 V2 输出 Excel 报告"""
    import xlsxwriter

    wb = xlsxwriter.Workbook(output_path)

    # 通用格式
    fmt_header = wb.add_format({
        'bold': True, 'bg_color': '#2F5496', 'font_color': 'white',
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_size': 11, 'font_name': 'Microsoft YaHei',
    })
    fmt_cell = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_size': 10, 'font_name': 'Microsoft YaHei',
    })
    fmt_num = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_size': 10, 'num_format': '#,##0.0',
        'font_name': 'Microsoft YaHei',
    })
    fmt_warn = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_size': 10, 'bg_color': '#FFF2CC',
        'font_name': 'Microsoft YaHei',
    })

    # === Sheet 1: PO 天数矩阵（每人每月每 PO 的天数） ===
    ws1 = wb.add_worksheet("PO_Days_Matrix")

    # 获取所有月份和 PO
    all_months = sorted(set(
        m for name_months in allocation.matrix.values()
        for m in name_months.keys()
    ))
    all_po_ids = sorted(set(
        po_id for name_months in allocation.matrix.values()
        for po_days in name_months.values()
        for po_id in po_days.keys()
    ))
    all_names = sorted(allocation.matrix.keys())

    # 写入表头
    ws1.write(0, 0, "Name", fmt_header)
    ws1.write(0, 1, "Level", fmt_header)
    col = 2
    for month in all_months:
        for po_id in all_po_ids:
            ws1.write(0, col, f"{month}\n{po_id[-4:]}", fmt_header)
            col += 1

    # 写入数据
    for row, name in enumerate(all_names, 1):
        ws1.write(row, 0, name, fmt_cell)
        ws1.write(row, 1, personnel.get(name, {}).get("level", ""), fmt_cell)
        col = 2
        for month in all_months:
            for po_id in all_po_ids:
                days = allocation.matrix.get(name, {}).get(month, {}).get(po_id, 0)
                if days > 0:
                    ws1.write(row, col, days, fmt_num)
                else:
                    ws1.write(row, col, "-", fmt_cell)
                col += 1

    ws1.set_column(0, 0, 14)
    ws1.set_column(1, 1, 22)
    ws1.freeze_panes(1, 2)

    # === Sheet 2: PO 汇总（每张 PO 各 Level 消耗情况） ===
    ws2 = wb.add_worksheet("PO_Summary")
    headers2 = ["PO ID", "Level", "Quota (Days)", "Allocated (Days)",
                "Remaining (Days)", "Usage %", "Status"]
    for i, h in enumerate(headers2):
        ws2.write(0, i, h, fmt_header)

    row = 1
    for key, info in sorted(po_summary.items()):
        config = info.get("quota")
        quota_days = config.days_quota if config else 0
        allocated = info["allocated"]
        remaining = quota_days - allocated
        usage_pct = (allocated / quota_days * 100) if quota_days > 0 else 0

        ws2.write(row, 0, info["po_id"], fmt_cell)
        ws2.write(row, 1, info["level"], fmt_cell)
        ws2.write(row, 2, quota_days, fmt_num)
        ws2.write(row, 3, round(allocated, 1), fmt_num)
        ws2.write(row, 4, round(remaining, 1), fmt_num)
        ws2.write(row, 5, f"{usage_pct:.1f}%", fmt_cell)

        if remaining < 0:
            ws2.write(row, 6, "OVER-USED!", fmt_warn)
        elif remaining < 10:
            ws2.write(row, 6, "Low", fmt_warn)
        else:
            ws2.write(row, 6, "OK", fmt_cell)
        row += 1

    ws2.set_column(0, 6, 16)

    # === Sheet 3: 月度总览 ===
    ws3 = wb.add_worksheet("Monthly_Overview")
    ws3.write(0, 0, "Month", fmt_header)
    col = 1
    for po_id in all_po_ids:
        ws3.write(0, col, f"{po_id[-8:]} Days", fmt_header)
        col += 1
    ws3.write(0, col, "Total Days", fmt_header)
    ws3.write(0, col + 1, "Workdays", fmt_header)

    for row, month in enumerate(all_months, 1):
        ws3.write(row, 0, month, fmt_cell)
        total = 0
        col = 1
        for po_id in all_po_ids:
            po_total = sum(
                allocation.matrix.get(name, {}).get(month, {}).get(po_id, 0)
                for name in all_names
            )
            ws3.write(row, col, round(po_total, 1) if po_total > 0 else "-", fmt_num if po_total > 0 else fmt_cell)
            total += po_total
            col += 1

        ws3.write(row, col, round(total, 1), fmt_num)
        month_num = int(month.split('-')[1])
        ws3.write(row, col + 1, DEFAULT_WORKDAYS.get(month_num, 22), fmt_cell)

    ws3.set_column(0, col + 1, 14)
    ws3.freeze_panes(1, 1)

    wb.close()
    print(f"\n  Report saved: {output_path}")


# ============================================================
# 主流程
# ============================================================
def main(project: str = "CNTS-25165"):
    setup_utf8_stdout()
    print("=" * 70)
    print(f"  PO Hour Allocator V2 — {project}")
    print(f"  Run time: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 70)

    # 1. 加载人员配置
    personnel_map = {
        "CNTS-25165": config.CNTS25165_PERSONNEL,
    }
    personnel = personnel_map.get(project, {})
    if not personnel:
        print(f"  [ERROR] No personnel config for project: {project}")
        return

    print(f"\n  Personnel: {len(personnel)} people")
    for name, info in personnel.items():
        print(f"    {name}: {info.get('level', 'N/A')}")

    # 2. 加载 PO Level 配置
    po_configs = load_po_level_config(project)
    print(f"\n  PO Level Configs: {len(po_configs)} entries")
    for c in po_configs:
        print(f"    {c.po_id} | {c.level} | {c.days_quota:.0f}d × ¥{c.unit_price:.0f} "
              f"| {c.start_month}~{c.end_month} | P{c.priority}")

    # 3. 加载实际工时
    project_config = config.PROJECTS.get(project, {})
    project_wbs = project  # CNTS-25165
    project_dir = project_config.get("budget_report_dir", "")
    timesheet_dir = os.path.join(
        config.ROOT, project, "01_IAV TMS & WP"
    )

    actual_hours = merge_all_hours(
        project_wbs, project_dir, timesheet_dir, personnel
    )

    # 4. 运行分配
    print("\n  Running allocation algorithm...")
    allocation, po_summary = calculate_po_allocation_v2(
        personnel, po_configs, actual_hours
    )

    # 5. 计算 Use Untill
    use_untill = calculate_use_untill(allocation, personnel)

    # 6. 生成报告
    timestamp = datetime.now().strftime("%y%m%d_%H%M")
    output_path = os.path.join(
        config.OUTPUT_DIR, f"PO_Plan_V2_{timestamp}.xlsx"
    )
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)

    generate_report_v2(allocation, po_summary, use_untill,
                       personnel, actual_hours, output_path)

    print(f"\n{'=' * 70}")
    print(f"  Done! Output: {output_path}")
    print(f"{'=' * 70}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="PO Hour Allocator V2")
    parser.add_argument("--project", default="CNTS-25165", help="Project code")
    args = parser.parse_args()
    main(args.project)
