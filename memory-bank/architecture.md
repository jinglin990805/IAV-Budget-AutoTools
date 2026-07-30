# Architecture Decisions

## 决策记录

### AD-001: 两个独立脚本 vs 一个统一工具
**决策**: 多模块独立脚本 + 统一工作台入口
**理由**: 使用场景不同、用户可按需单独运行、工作台 HTML 提供统一视图
**更新 (2026-07-28)**: 新增 travel_summary.py (差旅) 和 po_gr_tracker.py (状态追踪)

### AD-002: 配置文件格式 — Python config.py vs YAML
**决策**: Python config.py
**理由**: 用户已安装 Python，不需要额外 YAML 解析库；Python 文件可以有注释和类型提示

### AD-003: 输出策略 — 修改原文件 vs 生成新文件
**决策**: 仅生成新文件，不修改原始 Budget Report
**理由**: 防止数据损坏、可对比新旧结果、用户手动确认后再更新

### AD-004: 数据读取方式
**决策**: openpyxl（data_only=True）用于读取，xlsxwriter 用于写入
**理由**: openpyxl 支持 .xlsm 文件、可忽略公式仅读取数据值；xlsxwriter 格式控制更好

### AD-005: 编码处理
**决策**: 所有字符串输出使用 UTF-8，文件名中的全角字符做规范化处理
**理由**: Windows cp1252 控制台无法打印全角英文字符

### AD-006: 项目号映射
**决策**: 在 config.py 中维护 WBS_TO_PROJECT dict，支持前缀匹配 + 姓名别名映射
**更新 (2026-07-28)**: 新增 `build_name_lookup()` 函数，支持中英文姓名的双向查找

### AD-007: PO 分配算法 — Level 配额管理 (NEW)
**决策**: PO 按 Level 独立配额管理，不再使用统一资金池
**理由**: 客户 PO 内每个 Level 有固定天数额（如 180d×2100 + 85d×2260 + 464d×2460）
**实现**: `po_allocator_v2.py` 中 `calculate_po_allocation_v2()`，每个 Level 的配额独立消耗

### AD-008: 多数据源合并策略 (NEW)
**决策**: 工时数据三源合并，优先级：Budget Report > Monthly Closing > Timesheet PDF
**实现**: `merge_all_hours()` 函数，通过姓名别名映射统一人员标识

### AD-009: PO/GR 状态追踪 — 记录与呈现分离 (NEW)
**决策**: 数据层用 JSON (`data/po_gr_status.json`)，呈现层由 workbench.html 负责
**理由**: 避免当前 Excel 中"巨型矩阵既是记录又是呈现"的问题

### AD-010: 工作台架构 — 单文件 HTML + 本地 Python API (NEW)
**决策**: 工作台为单个 HTML 文件（零依赖），通过轻量 HTTP API 调用 Python 脚本
**理由**: 无需 npm/web 框架，浏览器直接打开，与现有 Python 生态无缝集成

---

## 文件职责地图

### 核心工具（Phase 1 重构）
| 文件 | 职责 | 数据源 |
|---|---|---|
| `po_allocator_v2.py` | PO 工时分配器 V2 — Level 配额算法 | Monthly Closing + Budget Report + Timesheet PDF |
| `po_pdf_parser.py` | 客户 PO PDF 解析 — 提取 Level 配额 | VCTC Call-off Purchase Order PDF |
| `timesheet_parser.py` | 第三方工时 PDF 解析 — 提取月度工时 | `01_IAV TMS & WP/YYYYMM/` PDF |
| `po_gr_tracker.py` | PO/GR 状态追踪 — JSON 读写 + CLI | `data/po_gr_status.json` |

### 核心工具（V1 保留）
| 文件 | 职责 | 数据源 |
|---|---|---|
| `po_allocator.py` | PO 工时分配器 V1（保留参考） | Monthly Closing + Budget Report |
| `budget_updater.py` | Budget 月度更新器 | Monthly Closing + PO_Record + 差旅文件夹 |
| `travel_summary.py` | 差旅费用汇总 | `03_Hose & Tengqing/YYYYMM/` |

### 基础设施
| 文件 | 职责 |
|---|---|
| `config.py` | 集中配置：路径、费率、人员（含别名）、WBS 映射 |
| `utils.py` | 共用工具：Excel I/O、文件查找、编码处理 |
| `run.bat` | 菜单式运行入口 |
| `CLAUDE.md` | Agent 指令 & 项目概览 |

### 数据
| 路径 | 内容 |
|---|---|
| `input/` | PO 配置 Excel 模板 |
| `output/` | 生成的报告（带时间戳） |
| `data/po_gr_status.json` | PO/GR 状态数据（结构化） |

### 项目记忆
| 文件 | 内容 |
|---|---|
| `memory-bank/design-document.md` | 需求澄清 & 功能设计 |
| `memory-bank/tech-stack.md` | 技术栈选择 |
| `memory-bank/implementation-plan.md` | 实施步骤 |
| `memory-bank/architecture.md` | 本文件 — 架构决策 & 文件职责 |
| `memory-bank/progress.md` | 进度日志 |

---

## 数据流

```
客户 PO PDF ──→ po_pdf_parser.py ──→ PO Level 配额
工时 PDF   ──→ timesheet_parser.py ──→ 第三方工时
                     │
Monthly Closing ─────┼──→ merge_all_hours() ──→ po_allocator_v2.py ──→ PO_Plan_V2_*.xlsx
Budget Report ───────┘
                     
PO Record ──→ budget_updater.py ──→ Budget_Summary_*.xlsx
差旅文件夹 ──→ travel_summary.py ──→ Travel_Summary_*.xlsx

po_gr_status.json ──→ po_gr_tracker.py ──→ CLI 查询 / workbench.html 呈现
```
