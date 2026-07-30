# Design Document: IAV Budget Automation

## 需求澄清结果

### Q1: 数据源是什么？
- **工时数据**: `Monthly Closing_Data_Base_YYYYMM.xlsx`（controlling 部门提供，每月更新）
  - 位置: `00_Project Auto Management Trial - Documents/02_TMS/`
  - 结构: 每月一个 sheet（Jan-Dec），列包括 EmployeeNumb、Receiver WBS element（项目号）、Total Quantity（工时h）、Project Number
- **辅助工时验证**: `00_TSE_All/`（各员工 .xlsm 工时文件）
- **差旅数据**: `03_Hose & Tengqing/YYYYMM/`（机票对账单、合思报销、IAV China 月度报表）
- **费率数据**: `Hourly Rate 2026.xlsx`（各 level 的 Cost 和 Transferprice）
- **员工数据**: `Employee.xlsx`（Personal ID、姓名、BU、Cost Center、Functional Role）
- **项目映射**: `Mapping.xlsx`（WBS → Project Leader → Cluster → Project Name）
- **PO 数据**: `PO_Record.xlsx`（TOP_WBS、PO号、金额、日期、备注）

### Q2: 最痛的点是什么？
**CNTS-25165 的 PO 工时计划。** 该项目特点：
- 3 个 level 的工程师共用 PO
- 客户要求：同一 level 的同事基本同时用完一张 PO 再出新 PO
- 不断上新人，工时安排持续调整
- 实际工时和计划不一致，切换 PO 时需要重新按实际工时统计
- 逻辑不复杂但各项目小有区别，每个项目都要手动执行一遍

### Q3: 使用场景和习惯？
- 用户是技术小白，不会 Coding，靠 AI vibe coding
- 接受 Python + Excel 方案，也接受小工具
- 希望能推荐给其他同事使用
- 已有 controlling 部门提供的 Cost Tracking Tools.exe（但只能抓已记入系统的数据）

### Q4: 具体分配规则（从实际数据反推）?
从 CNTS-25165 Internal Budget_FC 2026 Q2 的实际数据分析：

**过渡月份分配逻辑**（5月从 Q1 PO 切换到 Q2 PO 为例）:
```
规则1：旧 PO 剩余天数 > 当月实际天数 → 全部在旧 PO
规则2：旧 PO 剩余天数 = 0 → 全部在新 PO
规则3：0 < 旧 PO 剩余 < 当月实际天数 → 拆分（部分旧、部分新）
规则4：旧 PO 有负值 → 前期实际超用，在新 PO 中补偿调整
规则5：新人加入 → 从加入月开始分配，旧 PO 存量不受影响
```

**费率到天数的换算**:
- 每人日费率 = 对应 level 的 Transferprice × 8h
- PO 可支撑天数 = PO Value ÷ 该 PO 覆盖人员的日费率合计
- Senior: 2260 × 8 = 18,080/天
- Expert: 2460 × 8 = 19,680/天
- Experienced: 2100 × 8 = 16,800/天

### Q5: 输出格式要求?
- **两种都要**：
  1. 独立分析报告（新表格，含矩阵、甘特图数据、总结）
  2. 可基于现有模板填充（PO Usage Status 或 Budget Report 对应单元格）
- 非破坏性：不修改原始文件，生成新文件

### Q6: 技术水平?
- Python 3.12.10 已安装
- 已有库: openpyxl, pandas, xlsxwriter, python-docx, python-pptx, tkinter
- 缺少: pdfplumber（需安装）

## 功能需求

### 工具 A: PO 工时分配器
- F1: 读取 Hourly Rate 计算每人日费率
- F2: 读取实际工时（来自 Monthly Closing 或手动输入）
- F3: 按 PO 优先级分配每月工时到对应 PO
- F4: 计算过渡月份的拆分比例（旧 PO vs 新 PO）
- F5: 处理新人加入（从指定月份开始分配）
- F6: 处理偏差调整（实际超用在新 PO 中补偿）
- F7: 计算"Use Untill"预计用完日期
- F8: 输出独立报告 + 模板填充

### 工具 B: Budget Report 月度更新器
- F9: 解析 Monthly Closing 数据，按项目号筛选工时
- F10: 解析 Hose & Tengqing 机票/酒店/报销数据
- F11: 按项目和成本中心汇总工时和差旅
- F12: 读取 PO_Record 标记 PO 状态
- F13: 生成综合汇总 Excel

## 约束条件

- 代码必须是纯 Python，依赖库最少
- 所有路径配置集中在单一配置文件
- 输出文件不覆盖原文件
- 文件路径使用绝对路径（避免工作目录问题）
- 支持中文字符（UTF-8 编码）
- 不依赖外部 API 或网络
