# Tech Stack

## 运行时
- **Python 3.12.10**（用户已安装，Windows Store 版本）

## 核心依赖（已安装）
| 库 | 版本 | 用途 |
|----|------|------|
| openpyxl | latest | 读写 .xlsx/.xlsm 文件（支持公式、样式保留） |
| pandas | latest | 数据处理、筛选、聚合 |
| xlsxwriter | 3.2.9 | 创建带有格式的 Excel 输出（不支持读取） |

## 可选依赖
| 库 | 用途 |
|----|------|
| python-docx | 1.2.0 | 后续填充 Word 验收文档 |
| tkinter | 内置 | 后续可添加简单 GUI |

## 不需要
- ❌ 数据库（所有数据在 Excel 中，规模小）
- ❌ Web 框架（本地工具，不需要网页界面）
- ❌ API 调用（数据全部本地）
- ❌ Node.js / npm

## 项目结构
```
00_AutoTools/
├── config.py              # 单一配置文件（路径、费率映射、项目号）
├── utils.py               # 共用工具函数
├── po_allocator.py        # 工具 A: PO 工时分配器
├── budget_updater.py      # 工具 B: Budget 月度更新器
├── input/                 # 输入配置表（Excel）
├── output/                # 输出结果
├── memory-bank/           # 项目记忆
├── run.bat                # 双击运行入口
├── PRD.md
├── design-document.md
├── tech-stack.md
└── implementation-plan.md
```

## 设计原则
1. **最小依赖** — 只用已安装的库，避免额外安装步骤
2. **配置与代码分离** — 路径、费率等变动项全在 config.py
3. **单文件脚本** — 每个工具最多 300-400 行，便于 AI 辅助维护
4. **防御性编程** — 所有文件操作检查存在性、编码处理
