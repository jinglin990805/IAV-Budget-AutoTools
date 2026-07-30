# PO / WBS / 项目号对应关系

## 项目号 (WBS Element) 映射表

| 项目号 | 客户 | 项目名称 | 成本中心 | 负责人 |
|---|---|---|---|---|
| CNTS-25165 | VCTC | VCTC AI | 11000 | Lin Jing |
| CNTS-25048 | VCTC | VCTC AI (old) | — | — |
| CN03-22052 | CARIAD | PPE | 10000/11000/12000/13000/14000 | Lin Jing |
| CN03-22039 | CARIAD | CADA/CAPA | 11000/12000 | Liu Jiaqi |
| CN04-22028 | CARIAD | TTA | 11000/12000 | Li Yang |
| CN71-21043 | CARIAD | TTA2 | 10000 | — |
| CN03-22060 | BMW | BMW AD&ITS | — | — |
| CN03-22045 | CARIAD | CARIAD TF Other | — | — |
| CN03-22050 | CARIAD | CARIAD TF Other | — | — |
| CN04-24012 | — | CTN-2 | — | — |
| CN04-24025 | — | CTI-3 | — | — |
| CN05-24019 | VCTC | VCTC AAF | — | — |
| CNVA-25021 | CARIAD | CARIAD ADAS | — | — |

## Monthly Closing 中的项目号格式

Monthly Closing 数据（Monthly Closing_Data_Base_YYYYMM.xlsx）中的项目号在 **Col 10 (Project Number)**，格式为 `CNTS-25165-11026`（项目号 + 成本中心后缀）。

## WBS 前缀匹配规则

在 `config.py` 中维护了 `WBS_TO_PROJECT` 字典，支持前缀匹配。例如：
- `CNTS-25165` 匹配 `CNTS-25165-10026`、`CNTS-25165-11026`

## PO 编号格式

- **VCTC PO**: `PO-VCTC00-XXXXXXXX` (如 `PO-VCTC00-2607000312`)
- **CARIAD PO**: `73XXXXXXXX` (10 位数字，如 `7300058773`)
- **SAP PO**: `85XXXXXXXX` (SAP 内部编号)

## 人员姓名匹配

在 `config.py` 中维护了 `CNTS25165_PERSONNEL` 字典，每个人员包含中英文别名（aliases），支持双向查找。所有工时数据加载时通过 `build_name_lookup()` 统一标准化。
