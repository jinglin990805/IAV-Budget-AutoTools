# 季度验收全流程

## 概述

每个季度结束后与客户进行验收（Acceptance），确认上一季度的工作量和费用。如有分包商，还需进行供应商验收。

## 时间节点

| 节点 | 时间 | 负责 |
|---|---|---|
| 季度结束 | Q 末最后一天 | — |
| 工时汇总 | Q 末 + 1 周内 | Budget Controller |
| 差旅发票核对 | Q 末 + 2 周内 | Budget Controller |
| 验收材料准备 | Q 末 + 3 周内 | Budget Controller |
| 客户验收 | Q 末 + 1 月内 | 项目负责人 + Budget Controller |
| 供应商验收 | 客户验收完成后 | Budget Controller |
| 发票上传用友 | 供应商验收后 | Budget Controller |

## 验收材料清单

### 客户验收
- [ ] 工时清单（Timesheet 汇总）
- [ ] 差旅费用清单 + 发票复印件
- [ ] PO Usage Status（每张 PO 的消耗情况）
- [ ] GR（Goods Receipt）准备
- [ ] 客户要求的其他材料

### 供应商验收（分包商）
- [ ] 供应商工时确认
- [ ] 供应商发票
- [ ] 供应商验收报告
- [ ] 上传用友平台

## 差旅验收（最费时）

参考路径: `CN03-22052/05_Acceptance/2026/TF3_BT`

关键步骤:
1. 从 `travel_summary.py` 获取季度差旅汇总
2. 逐条核对发票（发票号、日期、金额、项目号）
3. 对于德国发票（Rechnungsnummer, Betrag），需额外核对字段
4. 将核对结果整理到验收材料中

## 相关工具

- PO Hours Allocator V2: `po_allocator_v2.py` → PO Usage Status
- Travel Summary: `travel_summary.py` → 差旅明细
- Budget Updater: `budget_updater.py` → 月度汇总数据
