# IAV Budget Automation Tools — Agent Instructions

## 项目概述

IAV TV 部门的预算管理自动化工具集，包含 PO 工时分配、Budget 月度更新、差旅汇总、PO/GR 状态追踪等模块。最终目标是一个统一的项目管理工作台（本地 HTML 页面）。

## 重要提示

- 写任何代码前必须完整阅读 memory-bank/architecture.md
- 写任何代码前必须完整阅读 memory-bank/design-document.md
- 每完成一个重大功能或里程碑后，必须更新 memory-bank/architecture.md
- 每完成一个实施步骤后，必须更新 memory-bank/progress.md

## 项目结构

```
00_AutoTools/
├── CLAUDE.md                # 本文件
├── config.py                # 集中配置（路径、费率、人员、WBS 映射）
├── utils.py                 # 共用工具函数（Excel I/O、文件查找、编码处理）
├── po_allocator_v2.py       # PO 工时分配器 V2（Level 配额算法）
├── po_allocator.py          # PO 工时分配器 V1（保留参考）
├── po_pdf_parser.py         # 客户 PO PDF 解析器
├── timesheet_parser.py      # 第三方工时 PDF 解析器
├── budget_updater.py        # Budget 月度更新器
├── travel_summary.py        # 差旅费用汇总
├── po_gr_tracker.py         # PO/GR 状态追踪
├── input/                   # 输入配置文件
├── output/                  # 输出结果
├── data/                    # 结构化状态数据（JSON）
├── memory-bank/             # 项目记忆（长期真相源）
│   ├── design-document.md   # 需求澄清 & 功能设计
│   ├── tech-stack.md        # 技术栈选择
│   ├── implementation-plan.md # 实施步骤
│   ├── architecture.md      # 架构决策 & 文件职责
│   └── progress.md          # 进度日志
├── knowledge-base/          # 知识库（FAQ + 流程文档 + 模板）
└── .claude/                 # Claude Code 配置
```

## 技术栈

详见 memory-bank/tech-stack.md

核心：Python 3.12 + openpyxl + pandas + xlsxwriter + pdfplumber
前端：单文件 HTML + vanilla CSS/JS
