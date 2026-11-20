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
# GR Documents (外援情况 / 工时清单表)
# ============================================================
GR_DOCUMENTS_DIR = os.path.join(ROOT, "CNTS-25165", "03_GR Documents")

# GR WP 合并输出根目录（按季度/PO/人员合并审核后工时 PDF）
VCTC_GR_YEAR_DIR = os.path.join(GR_DOCUMENTS_DIR, "2026", "VCTC")

# 季度 → 财务月份（用于 WP 合并的季度范围）
QUARTER_TO_MONTHS = {
    "Q1": ["2026-01", "2026-02", "2026-03"],
    "Q2": ["2026-04", "2026-05", "2026-06"],
    "Q3": ["2026-07", "2026-08", "2026-09"],
    "Q4": ["2026-10", "2026-11", "2026-12"],
}

# Rate → Level 反向映射
RATE_TO_LEVEL = {
    2460: "Testing Engineer-Expert",
    2260: "Testing Engineer-Senior",
    2100: "Testing Engineer-Experienced",
    1600: "Testing Engineer-Junior",
}

# 2026 实际工作日 (含中国法定假日调整)
WORKDAYS_2026 = {
    1: 19, 2: 17, 3: 22, 4: 22, 5: 19, 6: 22,
    7: 23, 8: 21, 9: 22, 10: 17, 11: 21, 12: 23,
}

# 月消耗率计算回溯窗口
CONSUMPTION_LOOKBACK_MONTHS = 3

# Monthly Closing 核查阈值 (差异 > 20% 告警)
MC_VERIFY_THRESHOLD = 0.20

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
    "ChenYakun": {
        "level": "Testing Engineer-Expert", "supervisor": "Wu Ning",
        "aliases": ["Chen Yakun", "Chen, Yakun", "Yakun Chen", "陈亚昆"],
        "join_month": "2026-01",
    },
    "Xu Kai": {
        "level": "Testing Engineer-Senior", "supervisor": "Wu Ning",
        "aliases": ["Xu Kai", "Kai Xu", "许凯", "徐凯"],
        "join_month": "2026-01",
    },
    "Liu Xing": {
        "level": "Testing Engineer-Expert", "supervisor": "Wu Ning",
        "aliases": ["Liu Xing", "Xing Liu", "刘星"],
        "join_month": "2026-01",
    },
    "Ge Andong": {
        "level": "Testing Engineer-Expert", "supervisor": "Wu Ning",
        "aliases": ["Ge Andong", "Andong Ge"],
        "join_month": "2026-07",
    },
    "He Bo": {
        "level": "Testing Engineer-Expert", "supervisor": "Wu Ning",
        "aliases": ["He Bo", "Bo He", "何博"],
        "join_month": "2026-07",
    },
    "Yang Yumeng": {
        "level": "Testing Engineer-Senior", "supervisor": "Wu Ning",
        "aliases": ["Yang Yumeng", "Yumeng Yang"],
        "join_month": "2026-07",
    },
    "PanChengchao": {
        "level": "Testing Engineer-Senior",
        "aliases": ["PanChengchao", "Pan Chengchao", "Panchengchao", "Chengchao Pan", "Pan, Chengchao"],
        "join_month": "2026-08",
    },
}

# 合并输出文件名过长时使用的姓名缩写（姓 + 名各音节首字母），与 Q2 WP 目录命名一致。
# 仅影响生成的文件/文件夹名，不改变系统内的人员标准名。
CNTS25165_NAME_ABBR = {
    "Tang Qiwei": "TangQW",
    "Li Xiaoxuan": "LiXX",
    "Yin Zicheng": "YinZC",
    "Zhao Shihui": "ZhaoSH",
    "Hou Jing": "HouJ",
    "Huang Dejia": "HuangDJ",
    "Liu Xu": "LiuXu",
    "Xuan Liangliang": "XuanLL",
    "Sun Shumin": "SunSM",
    "Qian Haiqing": "QianHQ",
    "Xu Chang": "XuC",
    "Lv Ansheng": "LvAS",
    "Cui Han": "CuiH",
    "ChenYakun": "ChenYK",
    "Xu Kai": "XuK",
    "Liu Xing": "LiuXing",
    "Ge Andong": "GeAD",
    "He Bo": "HeB",
    "Yang Yumeng": "YangYM",
    "PanChengchao": "PanCC",
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

# Budget Report 模板路径（增强版 V2 模板，含锚点标记）
BUDGET_REPORT_TEMPLATE_PATHS = {
    "CNTS-25165": os.path.join(
        ROOT, "CNTS-25165", "02_Cost Overview",
        "IAV_Budget_Report_VCTC AI_TEMPLATE_V2.xlsx"
    ),
    "CN03-22052": os.path.join(
        ROOT, "CN03-22052",
        "02_Budget Report & Earned Value Analyse", "Budget Report", "2026",
        "IAV_Budget Report_PPE_CN03-22052_TEMPLATE_V2.xlsx"
    ),
}

# 备份目录（write 模式自动备份）
BACKUP_DIR = os.path.join(AUTOTOOLS_DIR, "backups")

# 差异预览输出目录
DIFF_OUTPUT_DIR = os.path.join(AUTOTOOLS_DIR, "diffs")

# ============================================================
# 跨项目事实存储与项目 Profile
# ============================================================
# SQLite 只保存本地事实、来源指纹、Run、调整、异常和审批记录。
# Excel 仍是审核/公式/PivotTable/交付快照，不作为唯一事实数据库。
FACT_STORE_PATH = os.path.join(AUTOTOOLS_DIR, "data", "iav_facts.sqlite3")
FACT_SCHEMA_VERSION = 3
TESSERACT_EXE = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
    "Tesseract-OCR", "tesseract.exe",
)
OCR_LANGUAGE = "chi_sim+eng"
OCR_DPI = 300

# PaddleOCR benchmark/runtime environment.  Keep this isolated from the
# production interpreter so installing PaddlePaddle cannot change the current
# Tesseract route.  The setup/runner scripts also redirect model caches to the
# workspace because the managed Windows profile may be read-only.
PADDLE_ENV_DIR = os.path.join(AUTOTOOLS_DIR, ".paddle-venv")
PADDLE_PYTHON = os.path.join(PADDLE_ENV_DIR, "Scripts", "python.exe")
PADDLE_RUNTIME_HOME = os.path.join(AUTOTOOLS_DIR, ".paddle-home")
PADDLE_CACHE_DIR = os.path.join(AUTOTOOLS_DIR, ".paddle-cache")
PADDLE_PDX_CACHE_HOME = os.path.join(PADDLE_CACHE_DIR, "paddlex")
PADDLE_MODELSCOPE_CACHE = os.path.join(PADDLE_CACHE_DIR, "modelscope")
PADDLE_HF_HOME = os.path.join(PADDLE_CACHE_DIR, "huggingface")
PADDLE_TEMP_DIR = os.path.join(AUTOTOOLS_DIR, ".paddle-tmp")
PADDLE_LANG = "ch"
PADDLE_DEVICE = "cpu"
PADDLE_DISABLE_MODEL_SOURCE_CHECK = True

# Explicitly documented reconciliation exceptions. These remain visible as
# deferred review items and are not silently treated as coverage risks.
STEP2_DEFERRED_RECONCILIATION = {
    "CNTS-25165": {
        ("Cui Han", "2026-08"): "Exited project; no external Timesheet is expected",
    },
}

# 新流程的项目差异配置。旧 PROJECTS/CNTS25165_* 常量暂时保留，供旧脚本兼容；
# project_registry.py 只读取这里的项目级差异，不在核心逻辑中写死项目号。
PROJECT_PROFILE_OVERRIDES = {
    "CNTS-25165": {
        "root_dir": os.path.join(ROOT, "CNTS-25165"),
        "po_usage_template_path": os.path.join(
            ROOT, "CNTS-25165", "02_Cost Overview", "PO Usage Status_26917.xlsx"
        ),
        "budget_report_patterns": [
            "IAV_Budget_Report*.xlsx",
            "IAV_Budget Report*.xlsx",
        ],
        "source_dirs": {
            "timesheet_pdf": os.path.join(ROOT, "CNTS-25165", "01_IAV TMS & WP"),
            "po_pdf": os.path.join(ROOT, "CNTS-25165", "00_PO & CBD"),
            "gr_documents": os.path.join(ROOT, "CNTS-25165", "03_GR Documents"),
            "monthly_closing": os.path.join(ROOT, "01_PO & GR Status", "Raw_Data", "Monthly Closing"),
            "po_file": os.path.join(ROOT, "01_PO & GR Status", "Raw_Data", "PO file"),
            "cost_tracking": os.path.join(ROOT, "02_Cost Tracking Tool"),
            "finance_pending": TRAVEL_DIR,
        },
        "fiscal_year_start_month": 1,
        "currency": "CNY",
        "hours_per_day": 8,
        "display_unit": "day",
        "allocation_mode": "usage",
        "required_sources": [
            "timesheet_pdf", "monthly_closing", "po_file", "cost_tracking",
        ],
        "deliverables": {
            "po_usage_status": True,
            "customer_po_usage": True,
            "budget_report": True,
        },
        "macro_policy": "xlsx_only",
    },
    "CN03-22052": {
        "root_dir": os.path.join(ROOT, "CN03-22052"),
        "budget_report_patterns": [
            "IAV_Budget Report*.xlsx",
            "IAV_Budget_Report*.xlsx",
        ],
        "source_dirs": {
            "timesheet_pdf": os.path.join(ROOT, "CN03-22052", "01_TSE"),
            "po_pdf": os.path.join(ROOT, "CN03-22052", "03_PO"),
            "gr_documents": os.path.join(ROOT, "CN03-22052", "04_GR"),
            "monthly_closing": os.path.join(ROOT, "01_PO & GR Status", "Raw_Data", "Monthly Closing"),
            "po_file": os.path.join(ROOT, "01_PO & GR Status", "Raw_Data", "PO file"),
            "cost_tracking": os.path.join(ROOT, "CN03-22052", "02_Budget Report & Earned Value Analyse"),
            "finance_pending": os.path.join(ROOT, "CN03-22052", "06_Travel Cost & Materials"),
        },
        "fiscal_year_start_month": 1,
        "currency": "CNY",
        "hours_per_day": 8,
        "display_unit": "hour",
        "allocation_mode": "quarterly",
        "required_sources": [
            "timesheet_pdf", "monthly_closing", "po_file", "cost_tracking",
        ],
        "deliverables": {
            "po_usage_status": False,
            "customer_po_usage": False,
            "budget_report": True,
        },
        "acceptance_mapping": {
            "adapter": "overview_breakdown",
            "value_column_offset": 1,
            "metadata": {
                "po_number": "PO Number",
                "service_period": "Service Period",
                "project_name": "Project Name",
                "service_company": "Service Company",
                "service_department": "Service Department",
            },
        },
        "macro_policy": "xlsx_only",
    },
}

# 确认超时（秒）
CONFIRM_TIMEOUT_SECONDS = 60

# ============================================================
# 部门层级映射（按 Project Leader 分组）
# 从 Mapping.xlsx 读取 Project Definition → Project Leader → Department
# ============================================================
PROJECT_LEADER_TO_DEPT = {
    "Chai, Jiaqi": "TV",
    "Yu, Yongzhen": "TVA",
    "Lian, Tianyi": "TVM",
    "Qian, Haiqing": "TVX",
    "Li, Shengfei": "TVC",
}

# 部门显式顺序
DEPT_ORDER = ["TV", "TVA", "TVM", "TVX", "TVC"]


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
