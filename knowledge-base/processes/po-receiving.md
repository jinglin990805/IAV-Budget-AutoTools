# PO 接收 → 录入 → 通知流程

## 完整流程

```
1. Outlook 收到客户 PO 邮件（含 PDF 附件）
         ↓
2. 打开 MS Forms 填写 PO 信息
   https://forms.office.com/...
   - PO 编号
   - 金额
   - 项目号
   - 有效期
   - 备注
         ↓
3. Power Automate 自动执行：
   a. 写入 PO_Record.xlsx
   b. 生成草稿邮件发送给相关人员
         ↓
4. 【手动】存档 PO 文件到目标文件夹
   路径: {项目}/00_PO & CBD/{季度}/
   例: CNTS-25165/00_PO & CBD/26Q3/PO-VCTC00-2607000312.pdf
         ↓
5. 【手动】在草稿邮件中添加 PO PDF 附件
         ↓
6. 发送邮件通知相关人员
```

## PO_Record.xlsx 结构

路径: `OneDrive/.../00_Project Auto Management Trial - Documents/PO_Record.xlsx`

主要字段: PO ID、金额、项目号、收到日期、有效期、备注

## 待优化

- Power Automate 自动附加 PO 文件（根据 PO 号匹配邮件中的附件）
- 自动存档 PO 到项目文件夹（根据项目号匹配路径）

## 相关工具

- MS Forms PO 录入: `https://forms.office.com/Pages/ResponsePage.aspx?id=_qZESQ-LekSONc7F6Apc4JwfPlG6KdVOnHZIzLUNYeZUNVdNTVBBSUlSNTNBSUI4V0RMQThYMDZQRC4u`
- 工作台 PO/GR 状态: 可查看已录入的 PO 状态
