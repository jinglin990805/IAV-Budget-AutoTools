# Progress Log

## 2026-07-27 — Phase 1: MVP 完成 + 4 项修复

### ✅ 已完成的所有功能

**工具 A: PO 工时分配器 (po_allocator.py)**
- [x] 双数据源自动合并（Monthly Closing + Budget Report）
- [x] 每月独立天数显示（不累计到月末）
- [x] 输入配置表 (input/po_config.xlsx) — 添加新 PO 只需编辑 Excel
- [x] Use Untill 基于团队消耗速度计算（非单人）
- [x] 过渡月份拆分详情（Old PO vs New PO 天数对比）
- [x] 5 个输出 Sheet：PO天数矩阵、PO汇总、Use Untill、过渡详情、月度总览

**工具 B: Budget 月度更新器 (budget_updater.py)**
- [x] 从 Monthly Closing 汇总 124 个项目工时
- [x] 按项目号+成本中心+月度汇总
- [x] 26 个 PO 状态读取
- [x] 4 个输出 Sheet：工时、差旅、PO状态、项目概览

**工具 C: 差旅费用汇总 (travel_summary.py)** — 新增
- [x] 扫描 03_Hose & Tengqing 全部 23 个月数据
- [x] 解析 合思对账单（酒店/火车/机票/用车）
- [x] 解析 机票对账单（含项目号，可用 xlrd 读 .xls）
- [x] 按项目+人员+月度自动归类
- [x] 1,605 条差旅记录成功解析
- [x] 4 个输出 Sheet：明细、按项目、按月、按人员

**运行入口 (run.bat)**
- [x] 菜单循环（执行后回到菜单，不再退出）
- [x] 4 个选项：PO分配 / Budget更新 / 差旅 / 全部

---

## 2026-07-28 — Phase 1 Refactor: Level 配额算法重写

### 新建文件

- [x] **po_pdf_parser.py** — 客户 PO PDF 解析器
  - 用 pdfplumber 从 VCTC Call-off Purchase Order PDF 提取 Level 配额
  - 成功解析 2607000312: 180d×¥2,100 + 85d×¥2,260 + 464d×¥2,460 = ¥1,711,540
  - 自动匹配 Level 描述到标准名称（Experience → Experienced, Expert → Expert, Senior → Senior）

- [x] **timesheet_parser.py** — 第三方工时 PDF 解析器
  - 扫描 `CNTS-25165/01_IAV TMS & WP/` 下的 YYYYMM 目录
  - 81 个文本型 PDF（可提取工时），96 个扫描件（仅文件名解析）
  - 从文本提取：姓名、PO编号、月份、总工时
  - 从文件名解析：YYYYMM姓名IAV.pdf → 年月+中文名

- [x] **po_allocator_v2.py** — PO 工时分配器 V2
  - 核心改进：PO 按 Level 配额独立管理，不再统一资金池
  - 三源合并（Monthly Closing Col 10 Project Number + Budget Report + Timesheet PDF）
  - 姓名别名映射（中文名/英文名双向查找）
  - 加载 445 天工时（15 人），覆盖 2026-01 至 2026-06
  - 过渡月份拆分正确触发

### 修改文件

- [x] **config.py** — 新增中英文姓名别名映射 + `build_name_lookup()` 函数

---

## 2026-07-28 — Phase 2: PO/GR 状态追踪

### 新建文件

- [x] **data/po_gr_status.json** — 结构化状态数据
  - 5 个项目（CNTS-25165, CN03-22052, CN03-22039, CN04-22028, CN71-21043）
  - 每个项目包含 PO 列表，记录了 GR 状态、发票状态、截止日期
  - 记录与呈现分离：JSON 存数据，workbench.html 负责展示

- [x] **po_gr_tracker.py** — PO/GR 状态追踪 CLI
  - `list` — 列出所有项目概览
  - `show <project>` — 查看单个项目详情
  - `alerts` — 显示逾期和即将到期的提醒（🔴 overdue / 🟡 upcoming / ℹ️ info）
  - `update <project> <po_id> <field> <value>` — 更新 PO 状态
  - `add <project> <po_id> ...` — 添加新 PO

---

## 2026-07-28 — 项目文件规范化

### 执行 vibe-coding-standard-workflow

- [x] 设计文档移入 `memory-bank/`（design-document.md, tech-stack.md, implementation-plan.md）
- [x] 删除根目录 `PRD.md`（已被 grill-me session + 批准方案取代）
- [x] 创建 `CLAUDE.md`，包含 memory-bank 规则块
- [x] 更新 `memory-bank/architecture.md`（新增 AD-007~AD-010，文件职责地图）
- [x] 更新 `memory-bank/progress.md`（本文件）

---

## 2026-07-29 — Phase 3: 工作台 HTML 页面

### 新建文件

- [x] **api_server.py** — Flask 轻量 HTTP API 服务器
  - GET /api/status — 返回 PO/GR 状态 JSON
  - GET /api/status/alerts — 逾期/即将到期提醒（🔴danger/🟡warning/ℹ️info）
  - POST /api/status/update — 更新 PO 状态字段（JSON 持久化）
  - POST /api/run/<tool> — 运行 Python 工具（subprocess + 超时控制）
  - GET /api/outputs — 输出文件列表
  - GET /api/output/<filename> — 下载输出文件
  - GET /api/knowledge/search?q= — 全文搜索知识库
  - POST /api/translate — 翻译（占位，待接入 AI API）
  - GET / — 服务 workbench.html
  - 自动打开浏览器，端口 8765

- [x] **workbench.html** — 单文件工作台前端（42KB, 零依赖）
  - 左侧导航栏（CSS Grid 布局，7 个模块 + 分隔线）
  - 顶部搜索框（全局搜索项目号/PO号/文档，带下拉结果）
  - 首页看板：警报卡片 + 项目概览 + 关键时间线 + 快捷操作 + 最近输出
  - PO 工时规划页：一键运行按钮 + 实时输出终端
  - Budget 月度更新页：同上模式
  - 差旅汇总页：同上模式
  - PO/GR 状态页：筛选表格 + 内联编辑 + 模态框保存
  - 知识库页：搜索 + 卡片式浏览
  - 小工具页：翻译输入框 + 常用链接 + 输出文件列表
  - Toast 通知 + 编辑模态框
  - 响应式设计（支持桌面端）

### 验证结果
- ✅ /api/status → 5 个项目正确返回
- ✅ /api/status/alerts → 3 个提醒（info 级别）
- ✅ /api/outputs → 17 个输出文件列出
- ✅ /api/run/po_gr_tracker → 工具执行正常，stdout 捕获正确
- ✅ / → workbench.html 正确服务（42KB）
- ✅ 全部端点无 500 错误

---

## 2026-07-29 — Phase 4: 知识库搭建

### 新建文件

- [x] **knowledge-base/index.md** — 知识库目录索引（FAQ + 流程 + 模板）
- [x] **knowledge-base/faq/po-wbs-mapping.md** — PO/WBS/项目号映射表（15 个项目 + 格式说明）
- [x] **knowledge-base/faq/travel-reimbursement.md** — 差旅报销完整流程（3 个数据源 + 验收步骤）
- [x] **knowledge-base/faq/b2b-account.md** — B2B 账号申请流程（材料 + 步骤 + 注意事项）
- [x] **knowledge-base/faq/invoice-submit.md** — 发票提交流程（用友平台操作步骤）
- [x] **knowledge-base/processes/po-receiving.md** — PO 接收→录入→通知完整流程
- [x] **knowledge-base/processes/quarterly-acceptance.md** — 季度验收全流程（时间节点 + 材料清单）
- [x] **knowledge-base/processes/subcontractor.md** — 分包商采购 & 验收流程
- [x] **knowledge-base/templates/index.md** — 模板索引（B2B / 简历 / 验收 / 邮件）

### 验证结果
- ✅ 9 个 Markdown 文件，全中文内容
- ✅ API 搜索测试通过: "B2B"→3, "差旅"→4, "发票提交"→2, "验收"→6, "分包商"→3

---

## 2026-07-29 — Phase 5: 小工具集成 + run.bat 更新

### 修改文件

- [x] **run.bat** — 全面更新菜单
  - [1] PO 工时规划 (V1)
  - [2] PO 工时规划 (V2 - Level 配额) ← 新增
  - [3] Budget 月度更新
  - [4] 差旅汇总
  - [5] PO/GR 状态追踪 ← 新增
  - [6] 运行全部工具（更新为 5 步骤）
  - [W] 启动工作台（浏览器）← 新增
  - 中文菜单 + UTF-8 编码支持

### 工作台小工具（已集成在 workbench.html）
- ✅ 翻译输入框（中→英/英→中/中→德/德→中，API 占位待接入 AI）
- ✅ 常用链接（B2B 申请、发票提交、验收材料、邮件模板 → 跳转知识库）
- ✅ 外部链接（MS Forms PO 录入、PowerBI、AI 对话）
- ✅ 输出文件列表（API 实时加载）

---

## 🔜 待优化（非核心路径）

- [ ] Budget Report 分包商工时读取（恢复 Source 2 数据源）
- [ ] 扫描件 PDF 工时提取（OCR，96 个文件）
- [ ] PDF 发票自动提取（pdfplumber 实验）
- [ ] Power Automate PO 录入流程优化
- [ ] PowerBI 看板数据源对接
- [ ] 多项目 PO 分配泛化
- [ ] 翻译 API 接入（LLM）

### 待优化（非核心路径）
- [ ] Budget Report 分包商工时读取（恢复 Source 2 数据源）
- [ ] 扫描件 PDF 工时提取（OCR，96 个文件）
- [ ] PDF 发票自动提取（pdfplumber 实验）
- [ ] Power Automate PO 录入流程优化
- [ ] PowerBI 看板数据源对接
- [ ] 多项目 PO 分配泛化

---

## 输出文件清单

```
output/
├── PO_Plan_<timestamp>.xlsx       # PO分配结果 (V1)
├── PO_Plan_V2_<timestamp>.xlsx    # PO分配结果 (V2, Level 配额)
├── Budget_Summary_<timestamp>.xlsx # 预算月度汇总
└── Travel_Summary_<timestamp>.xlsx # 差旅费用汇总

data/
└── po_gr_status.json              # PO/GR 状态数据

input/
└── po_config.xlsx                 # PO配置表（V1，编辑后自动生效）
```
