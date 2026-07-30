# -*- coding: utf-8 -*-
"""
PO/GR 状态追踪器 — 记录与呈现分离

数据层: data/po_gr_status.json (结构化状态数据)
本工具: CLI 接口用于增删改查
呈现层: workbench.html (Phase 3)

用法:
    python po_gr_tracker.py list              # 列出所有项目状态
    python po_gr_tracker.py show CNTS-25165   # 查看单个项目详情
    python po_gr_tracker.py alerts            # 显示逾期和即将到期的提醒
    python po_gr_tracker.py update CNTS-25165 2607000312 gr_status done  # 更新状态
    python po_gr_tracker.py add CNTS-25165    # 交互式添加新 PO
"""

import json
import os
import sys
import io
from datetime import datetime, timedelta
from typing import Optional

# 配置
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "data", "po_gr_status.json")

STATUS_OPTIONS = ["pending", "done", "not_started", "n/a"]
FIELD_LABELS = {
    "po_id": "PO ID",
    "amount_rmb": "Amount (RMB)",
    "received_date": "Received",
    "valid_until": "Valid Until",
    "gr_status": "GR Status",
    "gr_date": "GR Date",
    "gr_deadline": "GR Deadline",
    "invoice_status": "Invoice",
    "invoice_date": "Invoice Date",
    "subcontract_ac_status": "Subcontract AC",
    "notes": "Notes",
}


def load_data() -> dict:
    """加载状态数据"""
    if not os.path.exists(DATA_FILE):
        return {"_schema_version": "1.0", "projects": []}
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_data(data: dict):
    """保存状态数据"""
    data["_last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def find_project(data: dict, project_code: str) -> Optional[dict]:
    """查找项目"""
    for p in data.get("projects", []):
        if p["project_code"].upper() == project_code.upper():
            return p
    return None


def find_po(project: dict, po_id: str) -> Optional[dict]:
    """查找 PO"""
    for po in project.get("pos", []):
        if po["po_id"] == po_id:
            return po
    return None


def get_alerts(data: dict, days_ahead: int = 30) -> list[dict]:
    """
    获取提醒：
    - 🔴 Overdue: GR deadline 已过且 gr_status != done
    - 🟡 Upcoming: GR deadline 在未来 days_ahead 天内
    """
    today = datetime.now().date()
    alerts = []

    for proj in data.get("projects", []):
        for po in proj.get("pos", []):
            gr_status = po.get("gr_status", "")
            deadline_str = po.get("gr_deadline", "")

            if gr_status == "done":
                continue

            if not deadline_str:
                # 没有 deadline 的 PO 如果 GR 未做也提醒
                if gr_status == "pending":
                    alerts.append({
                        "level": "info",
                        "project": proj["project_code"],
                        "project_name": proj.get("project_name", ""),
                        "po_id": po["po_id"],
                        "message": f"GR pending (no deadline set)",
                        "deadline": None,
                        "days_remaining": None,
                    })
                continue

            try:
                deadline = datetime.strptime(deadline_str, "%Y-%m-%d").date()
            except ValueError:
                continue

            days_remaining = (deadline - today).days

            if days_remaining < 0:
                alerts.append({
                    "level": "danger",
                    "project": proj["project_code"],
                    "project_name": proj.get("project_name", ""),
                    "po_id": po["po_id"],
                    "message": f"GR OVERDUE by {-days_remaining} days!",
                    "deadline": deadline_str,
                    "days_remaining": days_remaining,
                })
            elif days_remaining <= days_ahead:
                alerts.append({
                    "level": "warning",
                    "project": proj["project_code"],
                    "project_name": proj.get("project_name", ""),
                    "po_id": po["po_id"],
                    "message": f"GR due in {days_remaining} days",
                    "deadline": deadline_str,
                    "days_remaining": days_remaining,
                })

    # 按紧急程度排序
    level_order = {"danger": 0, "warning": 1, "info": 2}
    alerts.sort(key=lambda a: (level_order.get(a["level"], 99), a.get("days_remaining") or 0))

    return alerts


# ============================================================
# CLI Commands
# ============================================================
def cmd_list(data: dict):
    """列出所有项目概览"""
    print(f"\n{'='*80}")
    print(f"  PO/GR Status Overview — {datetime.now().strftime('%Y-%m-%d')}")
    print(f"{'='*80}")

    for proj in data.get("projects", []):
        name = proj.get("project_name", "")
        code = proj["project_code"]
        pos = proj.get("pos", [])
        total_pos = len(pos)
        pending_gr = sum(1 for p in pos if p.get("gr_status") == "pending")
        pending_inv = sum(1 for p in pos if p.get("invoice_status") == "pending")

        status_bar = ""
        if pending_gr > 0:
            status_bar += f"🔴{pending_gr} GR pending  "
        if pending_inv > 0:
            status_bar += f"📝{pending_inv} Invoice pending  "
        if not status_bar:
            status_bar = "✅ All clear"

        print(f"\n  [{code}] {name}")
        print(f"    Responsible: {proj.get('responsible', 'N/A')}")
        print(f"    POs: {total_pos} | {status_bar}")
        for po in pos:
            po_short = po["po_id"][-8:] if len(po["po_id"]) > 8 else po["po_id"]
            gr_icon = "✅" if po.get("gr_status") == "done" else "⏳"
            inv_icon = "✅" if po.get("invoice_status") == "done" else "📝"
            deadline = po.get("gr_deadline", "") or ""
            amount = po.get("amount_rmb", 0)
            amt_str = f"¥{amount:,.0f}" if amount else "¥?"
            print(f"      {gr_icon}{inv_icon} {po_short}  {amt_str}  {deadline}  "
                  f"{po.get('notes', '')[:50]}")


def cmd_show(data: dict, project_code: str):
    """显示单个项目详情"""
    proj = find_project(data, project_code)
    if not proj:
        print(f"  [ERROR] Project not found: {project_code}")
        return

    print(f"\n{'='*80}")
    print(f"  {proj['project_code']} — {proj.get('project_name', '')}")
    print(f"  Customer: {proj.get('customer', '')}  |  "
          f"Departments: {', '.join(proj.get('cost_centers', []))}")
    print(f"  Responsible: {proj.get('responsible', 'N/A')}  |  "
          f"Coordinator: {proj.get('iavs_coordinator', 'N/A')}")
    print(f"{'='*80}")

    for po in proj.get("pos", []):
        print(f"\n  ┌─ PO: {po['po_id']}")
        print(f"  ├─ Amount: ¥{po.get('amount_rmb', 0):,.0f}")
        print(f"  ├─ Received: {po.get('received_date', 'N/A')}")
        print(f"  ├─ Valid Until: {po.get('valid_until', 'N/A')}")
        gr = po.get('gr_status', '')
        gr_icon = "✅" if gr == "done" else ("⏳" if gr == "pending" else "—")
        print(f"  ├─ GR: {gr_icon} {gr}  |  Date: {po.get('gr_date', 'N/A')}  "
              f"|  Deadline: {po.get('gr_deadline', 'N/A')}")
        inv = po.get('invoice_status', '')
        inv_icon = "✅" if inv == "done" else ("📝" if inv == "pending" else "—")
        print(f"  ├─ Invoice: {inv_icon} {inv}  |  Date: {po.get('invoice_date', 'N/A')}")
        print(f"  ├─ Subcontract AC: {po.get('subcontract_ac_status', 'N/A')}")
        print(f"  └─ Notes: {po.get('notes', '')}")


def cmd_alerts(data: dict):
    """显示提醒"""
    alerts = get_alerts(data)

    print(f"\n{'='*80}")
    print(f"  Alerts & Reminders — {datetime.now().strftime('%Y-%m-%d')}")
    print(f"{'='*80}")

    if not alerts:
        print("\n  ✅ All POs are on track. No overdue or upcoming deadlines.")
        return

    level_labels = {"danger": "🔴 OVERDUE", "warning": "🟡 UPCOMING", "info": "ℹ️  INFO"}

    for a in alerts:
        label = level_labels.get(a["level"], a["level"])
        proj_name = a.get("project_name", "")
        print(f"\n  {label}")
        print(f"    [{a['project']}] {proj_name} — {a['po_id']}")
        print(f"    {a['message']}")
        if a.get("deadline"):
            print(f"    Deadline: {a['deadline']}")


def cmd_update(data: dict, project_code: str, po_id: str, field: str, value: str):
    """更新 PO 状态字段"""
    proj = find_project(data, project_code)
    if not proj:
        print(f"  [ERROR] Project not found: {project_code}")
        return

    po = find_po(proj, po_id)
    if not po:
        print(f"  [ERROR] PO not found: {po_id} in {project_code}")
        return

    if field not in po:
        print(f"  [ERROR] Unknown field: {field}. Available: {list(po.keys())}")
        return

    old_value = po[field]
    po[field] = value
    save_data(data)

    print(f"  ✅ Updated {project_code}/{po_id}: {field} = {old_value} → {value}")


def cmd_add_po(data: dict, project_code: str, po_id: str,
               amount: float = 0, received_date: str = "",
               gr_deadline: str = "", notes: str = ""):
    """添加新 PO"""
    proj = find_project(data, project_code)
    if not proj:
        print(f"  [ERROR] Project not found: {project_code}")
        return

    if find_po(proj, po_id):
        print(f"  [WARN] PO {po_id} already exists in {project_code}. Use update instead.")
        return

    new_po = {
        "po_id": po_id,
        "amount_rmb": amount,
        "received_date": received_date or datetime.now().strftime("%Y-%m-%d"),
        "valid_until": "",
        "gr_status": "pending" if gr_deadline else "not_started",
        "gr_date": None,
        "gr_deadline": gr_deadline,
        "invoice_status": "pending",
        "invoice_date": None,
        "subcontract_ac_status": "n/a",
        "notes": notes,
    }
    proj.setdefault("pos", []).append(new_po)
    save_data(data)

    print(f"  ✅ Added PO {po_id} to {project_code}")
    print(f"     Amount: ¥{amount:,.0f}  |  "
          f"Received: {received_date}  |  "
          f"GR Deadline: {gr_deadline or 'N/A'}")


# ============================================================
# Main
# ============================================================
def main():
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    if len(sys.argv) < 2:
        print("Usage:")
        print("  python po_gr_tracker.py list")
        print("  python po_gr_tracker.py show <project>")
        print("  python po_gr_tracker.py alerts")
        print("  python po_gr_tracker.py update <project> <po_id> <field> <value>")
        print("  python po_gr_tracker.py add <project> <po_id> [amount] [received_date] [gr_deadline] [notes]")
        return

    cmd = sys.argv[1].lower()
    data = load_data()

    if cmd == "list":
        cmd_list(data)
    elif cmd == "show" and len(sys.argv) >= 3:
        cmd_show(data, sys.argv[2])
    elif cmd == "alerts":
        cmd_alerts(data)
    elif cmd == "update" and len(sys.argv) >= 6:
        cmd_update(data, sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
    elif cmd == "add" and len(sys.argv) >= 4:
        amount = float(sys.argv[4]) if len(sys.argv) > 4 else 0
        received = sys.argv[5] if len(sys.argv) > 5 else ""
        deadline = sys.argv[6] if len(sys.argv) > 6 else ""
        notes = sys.argv[7] if len(sys.argv) > 7 else ""
        cmd_add_po(data, sys.argv[2], sys.argv[3], amount, received, deadline, notes)
    else:
        print(f"  [ERROR] Unknown command or missing arguments: {cmd}")


if __name__ == "__main__":
    main()
