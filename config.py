# -*- coding: utf-8 -*-
"""
IAV Budget Automation Tools — 配置文件
修改此文件即可适配不同项目和路径，无需改代码
"""

import os

# ============================================================
# 根路径
# ============================================================
ROOT = r"C:\Users\LinJing(CN-TV-F)\Desktop\Jing\00_IAV_TV\02_Budget Report"
AUTOTOOLS_DIR = os.path.join(ROOT, "00_AutoTools")
INPUT_DIR = os.path.join(AUTOTOOLS_DIR, "input")
OUTPUT_DIR = os.path.join(AUTOTOOLS_DIR, "output")
TRIAL_DIR = os.path.join(ROOT, "00_Project Auto Management Trial - Documents")

# ============================================================
# 主数据文件路径
# ============================================================
MASTER_DATA = os.path.join(TRIAL_DIR, "01_Master_Data")
HOURLY_RATE_FILE = os.path.join(MASTER_DATA, "Hourly Rate 2026.xlsx")
EMPLOYEE_FILE = os.path.join(MASTER_DATA, "Employee.xlsx")
MAPPING_FILE = os.path.join(MASTER_DATA, "Mapping.xlsx")
CN43N_FILE = os.path.join(MASTER_DATA, "CN43N.xlsx")

# Monthly Closing 数据（需每月更新文件名）
TMS_DIR = os.path.join(TRIAL_DIR, "02_TMS")
# 使用 latest 模式：自动找最新日期的文件
MONTHLY_CLOSING_PATTERN = os.path.join(TMS_DIR, "Monthly Closing_Data_Base_*.xlsx")

# PO 记录
PO_RECORD_FILE = os.path.join(TRIAL_DIR, "PO_Record.xlsx")

# ============================================================
# 差旅数据
# ============================================================
TRAVEL_DIR = os.path.join(ROOT, "03_Hose & Tengqing")

# ============================================================
# 核心项目路径
# ============================================================
PROJECTS = {
    "CN03-22052": {
        "name": "PPE",
        "budget_report_dir": os.path.join(ROOT, "CN03-22052",
                                          "02_Budget Report & Earned Value Analyse",
                                          "Budget Report", "2026"),
        "travel_cost_dir": os.path.join(ROOT, "CN03-22052",
                                        "06_Travel Cost & Materials"),
        "cost_centers": ["10000", "11000", "12000", "13000", "14000"],
    },
    "CN03-22039": {
        "name": "CADA/CAPA",
        "budget_report_dir": os.path.join(ROOT, "CN03-22039",
                                          "02_Budget Report & Earned Value Analyse",
                                          "Budget Report", "2026"),
        "cost_centers": ["11000", "12000"],
    },
    "CNTS-25165": {
        "name": "VCTC AI",
        "budget_report_dir": os.path.join(ROOT, "CNTS-25165", "02_Cost Overview"),
        "po_usage_dir": os.path.join(ROOT, "CNTS-25165", "02_Cost Overview"),
        "cost_centers": ["11000"],
    },
    "CN04-22028": {
        "name": "TTA",
        "budget_report_dir": os.path.join(ROOT, "CN04-22028 & CN71-21043",
                                          "Budget Report & Earned Value Analyse"),
        "cost_centers": ["11000", "12000"],
    },
    "CN71-21043": {
        "name": "TTA2",
        "budget_report_dir": os.path.join(ROOT, "CN04-22028 & CN71-21043",
                                          "Budget Report & Earned Value Analyse"),
        "cost_centers": ["10000"],
    },
}

# ============================================================
# Hourly Rate 配置（2026年，基于 Transferprice）
# Level -> (Cost, Transferprice, SArtTyp)
# ============================================================
HOURLY_RATES_2026 = {
    # SArtTyp: (Level名称, Cost, Transferprice)
    "EM001": ("Executive Engineer", 638.11, 670.02),
    "EM002": ("Senior Engineer", 487.54, 511.92),
    "EM005": ("Engineer", 376.34, 395.16),
    "EM006": ("Junior Engineer", 301.68, 316.76),
    "EM009": ("Expat", 960.59, 1008.62),
    "EM007": ("Student", 27.88, 29.27),
}

# VCTC 项目特殊费率（客户定价不同于内部 Transferprice）
VCTC_RATES = {
    "Testing Engineer-Senior": 2260,
    "Testing Engineer-Expert": 2460,
    "Testing Engineer-Experienced": 2100,
    "Testing Engineer-Junior": 1600,
}

# 每人每天工作小时数
HOURS_PER_DAY = 8

# ============================================================
# 项目号映射（Monthly Closing 的 WBS element → 项目名）
# 用于工时筛选，支持前缀匹配
# ============================================================
WBS_TO_PROJECT = {
    "CN03-22052": "PPE",
    "CN03-22039": "CADA/CAPA",
    "CNTS-25165": "VCTC AI",
    "CN04-22028": "TTA",
    "CN71-21043": "TTA2",
    "CN03-22060": "BMW AD&ITS",
    "CN03-22045": "CARIAD TF Other",
    "CN03-22050": "CARIAD TF Other",
    "CNTS-25048": "VCTC AI (old)",
    "CN04-24012": "CTN-2",
    "CN04-24025": "CTI-3",
    "CN05-24019": "VCTC AAF",
    "CNTS-25052": "VCTC AI (deliverable)",
    "CNTS-25071": "VCTC AAE",
    "CNVA-25021": "CARIAD ADAS",
}

# ============================================================
# CNTS-25165 PO 配置（可从 Excel 覆盖）
# ============================================================
CNTS25165_PO_CONFIG = [
    # (PO_ID, Value_RMB, people_list, start_month, end_month, priority)
    # 2025 个人 PO（已结束）
    # 2026 Q1 分组 PO
    ("2601000336", 324720, ["Cui Han", "Chen Yakun"], "2026-01", "2026-04", 1),
    ("2601000337", 623040, ["Tang Qiwei", "Zhao Shihui", "Yin Zicheng", "Lv Ansheng"],
     "2026-01", "2026-04", 1),
    ("2601000338", 855360,
     ["Li Xiaoxuan", "Hou Jing", "Huang Dejia", "Xuan Liangliang", "Liu Xu", "Sun Shumin"],
     "2026-01", "2026-04", 1),
    # 2026 Q2-Q3 全员 PO
    ("2604000799", 1884400, "all", "2026-04", "2026-09", 2),
    # 2026 Q3-Q4 全员 PO
    ("2607000312", 1711540, "all", "2026-07", "2026-12", 3),
]

# ============================================================
# CNTS-25165 人员-Level 映射
# 每个人员包含: level, supervisor, 以及所有可能出现的名字变体
# ============================================================
CNTS25165_PERSONNEL = {
    "Tang Qiwei": {
        "level": "Testing Engineer-Senior", "supervisor": "Wu Ning",
        "aliases": ["Tang Qiwei", "Qiwei Tang", "汤其伟"],
        "join_month": "2026-01",
    },
    "Li Xiaoxuan": {
        "level": "Testing Engineer-Expert", "supervisor": "Wu Ning",
        "aliases": ["Li Xiaoxuan", "Xiaoxuan Li", "李晓璇"],
        "join_month": "2026-01",
    },
    "Yin Zicheng": {
        "level": "Testing Engineer-Expert", "supervisor": "Wu Ning",
        "aliases": ["Yin Zicheng", "Zicheng Yin", "尹自成"],
        "join_month": "2026-01",
    },
    "Zhao Shihui": {
        "level": "Testing Engineer-Senior", "supervisor": "Yu Minghui",
        "aliases": ["Zhao Shihui", "Shihui Zhao", "赵士辉"],
        "join_month": "2026-01",
    },
    "Hou Jing": {
        "level": "Testing Engineer-Experienced", "supervisor": "Yu Minghui",
        "aliases": ["Hou Jing", "Jing Hou", "侯景"],
        "join_month": "2026-01",
    },
    "Huang Dejia": {
        "level": "Testing Engineer-Experienced", "supervisor": "Yu Minghui",
        "aliases": ["Huang Dejia", "Dejia Huang", "黄德佳"],
        "join_month": "2026-01",
    },
    "Liu Xu": {
        "level": "Testing Engineer-Experienced", "supervisor": "Yu Minghui",
        "aliases": ["Liu Xu", "Xu Liu", "刘旭"],
        "join_month": "2026-01",
    },
    "Xuan Liangliang": {
        "level": "Testing Engineer-Experienced", "supervisor": "Yu Minghui",
        "aliases": ["Xuan Liangliang", "Liangliang Xuan", "轩亮亮"],
        "join_month": "2026-01",
    },
    "Sun Shumin": {
        "level": "Testing Engineer-Experienced", "supervisor": "Yu Minghui",
        "aliases": ["Sun Shumin", "Shumin Sun", "孙舒敏"],
        "join_month": "2026-01",
    },
    "Qian Haiqing": {
        "level": "Testing Engineer-Expert", "supervisor": "Wu Ning",
        "aliases": ["Qian Haiqing", "Haiqing Qian", "钱海青"],
        "join_month": "2026-01",
    },
    "Xu Chang": {
        "level": "Testing Engineer-Expert", "supervisor": "Wu Ning",
        "aliases": ["Xu Chang", "Chang Xu", "许昌"],
        "join_month": "2026-05",
    },
    "Lv Ansheng": {
        "level": "Testing Engineer-Expert", "supervisor": "Wu Ning",
        "aliases": ["Lv Ansheng", "Ansheng Lv", "吕安生"],
        "join_month": "2026-05",
    },
    "Cui Han": {
        "level": "Testing Engineer-Expert", "supervisor": "Wu Ning",
        "aliases": ["Cui Han", "Han Cui", "崔涵"],
        "join_month": "2026-01",
    },
    "Chen Yakun": {
        "level": "Testing Engineer-Expert", "supervisor": "Wu Ning",
        "aliases": ["Chen Yakun", "Yakun Chen", "陈亚昆"],
        "join_month": "2026-01",
    },
    "Xu Kai": {
        "level": "Testing Engineer-Senior", "supervisor": "Wu Ning",
        "aliases": ["Xu Kai", "Kai Xu", "许凯"],
        "join_month": "2026-01",
    },
    "Liu Xing": {
        "level": "Testing Engineer-Expert", "supervisor": "Wu Ning",
        "aliases": ["Liu Xing", "Xing Liu", "刘星"],
        "join_month": "2026-01",
    },
}

# 构建反向查找表: alias → canonical name
def build_name_lookup(personnel_dict: dict) -> dict:
    """构建所有名字变体到标准名的映射"""
    lookup = {}
    for canonical, info in personnel_dict.items():
        for alias in info.get("aliases", [canonical]):
            lookup[alias.lower()] = canonical
    return lookup

# ============================================================
# 工具自身配置
# ============================================================
OUTPUT_FORMAT = "xlsx"  # 输出文件格式
DATE_FORMAT = "%Y-%m-%d"
MONTH_FORMAT = "%Y-%m"


def verify_paths():
    """验证所有配置的路径是否存在，打印检查结果"""
    paths_to_check = {
        "Hourly Rate": HOURLY_RATE_FILE,
        "Employee": EMPLOYEE_FILE,
        "Mapping": MAPPING_FILE,
        "PO Record": PO_RECORD_FILE,
        "TMS Dir": TMS_DIR,
        "Travel Dir": TRAVEL_DIR,
        "Input Dir": INPUT_DIR,
        "Output Dir": OUTPUT_DIR,
    }
    for name, path in paths_to_check.items():
        exists = os.path.exists(path)
        status = "✓" if exists else "✗ MISSING"
        print(f"  [{status}] {name}: {path}")


if __name__ == "__main__":
    import sys
    import io
    # Fix Windows console encoding for Chinese characters
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    print("=== IAV Budget Automation Tools - Config Verification ===\n")
    verify_paths()
    print("\nVerification complete.")
