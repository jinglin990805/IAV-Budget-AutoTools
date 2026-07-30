# -*- coding: utf-8 -*-
"""
IAV 工作台 API 服务器 — 轻量 Flask HTTP 服务

提供:
- GET  /api/status          — 返回 PO/GR 状态 JSON
- GET  /api/status/alerts   — 返回逾期/即将到期提醒
- POST /api/run/po_allocator — 运行 PO 分配器 V2
- POST /api/run/budget       — 运行 Budget 月度更新器
- POST /api/run/travel       — 运行差旅汇总
- GET  /api/outputs          — 列出 output/ 目录文件
- GET  /api/knowledge        — 搜索知识库
- POST /api/translate        — 翻译（占位）
- GET  /                    — 服务 workbench.html

用法:
    python api_server.py
    python api_server.py --port 8765
"""

import os
import sys
import io
import json
import glob
import subprocess
import threading
import time
import webbrowser
from datetime import datetime

try:
    from flask import Flask, jsonify, request, send_from_directory, send_file
except ImportError:
    print("[ERROR] Flask is required. Install: pip install flask")
    sys.exit(1)

# 路径配置
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
KNOWLEDGE_DIR = os.path.join(BASE_DIR, "knowledge-base")
STATUS_FILE = os.path.join(DATA_DIR, "po_gr_status.json")

app = Flask(__name__)

# ============================================================
# 状态数据
# ============================================================
def load_status():
    """加载 PO/GR 状态 JSON"""
    if not os.path.exists(STATUS_FILE):
        return {"error": "Status file not found", "projects": []}
    with open(STATUS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def get_alerts():
    """获取提醒（从 po_gr_tracker 逻辑复制）"""
    data = load_status()
    today = datetime.now().date()
    alerts = []

    for proj in data.get("projects", []):
        for po in proj.get("pos", []):
            gr_status = po.get("gr_status", "")
            deadline_str = po.get("gr_deadline", "")

            if gr_status == "done":
                continue

            if not deadline_str:
                if gr_status == "pending":
                    alerts.append({
                        "level": "info",
                        "project": proj["project_code"],
                        "project_name": proj.get("project_name", ""),
                        "po_id": po["po_id"],
                        "message": "GR pending (no deadline set)",
                        "deadline": None,
                        "days_remaining": None,
                        "amount_rmb": po.get("amount_rmb", 0),
                    })
                continue

            try:
                deadline = datetime.strptime(deadline_str, "%Y-%m-%d").date()
            except ValueError:
                continue

            days = (deadline - today).days
            if days < 0:
                alerts.append({
                    "level": "danger",
                    "project": proj["project_code"],
                    "project_name": proj.get("project_name", ""),
                    "po_id": po["po_id"],
                    "message": f"GR OVERDUE by {-days} days!",
                    "deadline": deadline_str,
                    "days_remaining": days,
                    "amount_rmb": po.get("amount_rmb", 0),
                })
            elif days <= 30:
                alerts.append({
                    "level": "warning",
                    "project": proj["project_code"],
                    "project_name": proj.get("project_name", ""),
                    "po_id": po["po_id"],
                    "message": f"GR due in {days} days",
                    "deadline": deadline_str,
                    "days_remaining": days,
                    "amount_rmb": po.get("amount_rmb", 0),
                })

    level_order = {"danger": 0, "warning": 1, "info": 2}
    alerts.sort(key=lambda a: (level_order.get(a["level"], 99), a.get("days_remaining") or 0))
    return alerts


# ============================================================
# API 路由
# ============================================================
@app.route("/")
def index():
    """服务 workbench.html"""
    html_path = os.path.join(BASE_DIR, "workbench.html")
    if os.path.exists(html_path):
        return send_file(html_path)
    return jsonify({"error": "workbench.html not found"}), 404


@app.route("/api/status")
def api_status():
    """返回完整 PO/GR 状态"""
    data = load_status()
    return jsonify(data)


@app.route("/api/status/alerts")
def api_alerts():
    """返回提醒列表"""
    return jsonify({"alerts": get_alerts()})


@app.route("/api/status/update", methods=["POST"])
def api_update_status():
    """更新 PO 状态字段"""
    body = request.get_json()
    project_code = body.get("project")
    po_id = body.get("po_id")
    field = body.get("field")
    value = body.get("value")

    if not all([project_code, po_id, field]):
        return jsonify({"error": "Missing required fields"}), 400

    data = load_status()
    for proj in data.get("projects", []):
        if proj["project_code"].upper() == project_code.upper():
            for po in proj.get("pos", []):
                if po["po_id"] == po_id:
                    old = po.get(field)
                    po[field] = value
                    data["_last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                    # 保存
                    os.makedirs(os.path.dirname(STATUS_FILE), exist_ok=True)
                    with open(STATUS_FILE, "w", encoding="utf-8") as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                    return jsonify({
                        "ok": True,
                        "project": project_code,
                        "po_id": po_id,
                        "field": field,
                        "old": old,
                        "new": value,
                    })
            return jsonify({"error": f"PO not found: {po_id}"}), 404
    return jsonify({"error": f"Project not found: {project_code}"}), 404


@app.route("/api/run/<tool>", methods=["POST"])
def api_run_tool(tool):
    """运行 Python 工具脚本"""
    scripts = {
        "po_allocator": "po_allocator_v2.py",
        "budget": "budget_updater.py",
        "travel": "travel_summary.py",
        "po_gr_tracker": "po_gr_tracker.py",
    }

    if tool not in scripts:
        return jsonify({"error": f"Unknown tool: {tool}. Available: {list(scripts.keys())}"}), 404

    script = scripts[tool]
    script_path = os.path.join(BASE_DIR, script)
    if not os.path.exists(script_path):
        return jsonify({"error": f"Script not found: {script}"}), 404

    try:
        result = subprocess.run(
            [sys.executable, script_path],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            timeout=120,
        )

        # 找最新生成的输出文件
        latest_outputs = []
        for f in sorted(glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx")), key=os.path.getmtime, reverse=True)[:3]:
            latest_outputs.append({
                "name": os.path.basename(f),
                "size": os.path.getsize(f),
                "mtime": datetime.fromtimestamp(os.path.getmtime(f)).strftime("%Y-%m-%d %H:%M"),
            })

        return jsonify({
            "ok": result.returncode == 0,
            "tool": tool,
            "stdout": result.stdout[-2000:] if result.stdout else "",
            "stderr": result.stderr[-500:] if result.stderr else "",
            "returncode": result.returncode,
            "latest_outputs": latest_outputs,
        })
    except subprocess.TimeoutExpired:
        return jsonify({"error": f"Tool {tool} timed out (120s)"}), 504
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/outputs")
def api_outputs():
    """列出输出文件"""
    files = []
    for f in sorted(glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx")), key=os.path.getmtime, reverse=True):
        files.append({
            "name": os.path.basename(f),
            "size": os.path.getsize(f),
            "mtime": datetime.fromtimestamp(os.path.getmtime(f)).strftime("%Y-%m-%d %H:%M"),
        })
    return jsonify({"outputs": files[:20]})


@app.route("/api/output/<filename>")
def api_download_output(filename):
    """下载输出文件"""
    filepath = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(filepath):
        return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)
    return jsonify({"error": "File not found"}), 404


@app.route("/api/knowledge/search")
def api_knowledge_search():
    """搜索知识库"""
    query = request.args.get("q", "").lower()
    if not query:
        return jsonify({"results": []})

    results = []
    if os.path.exists(KNOWLEDGE_DIR):
        for root, dirs, files in os.walk(KNOWLEDGE_DIR):
            for fname in files:
                if fname.endswith(".md"):
                    fpath = os.path.join(root, fname)
                    try:
                        with open(fpath, "r", encoding="utf-8") as f:
                            content = f.read()
                    except Exception:
                        content = ""

                    if query in content.lower() or query in fname.lower():
                        # 提取匹配上下文
                        idx = content.lower().find(query)
                        snippet = ""
                        if idx >= 0:
                            start = max(0, idx - 60)
                            end = min(len(content), idx + len(query) + 60)
                            snippet = content[start:end]

                        rel_path = os.path.relpath(fpath, KNOWLEDGE_DIR)
                        results.append({
                            "file": rel_path,
                            "snippet": snippet.strip(),
                            "title": fname.replace(".md", "").replace("-", " ").title(),
                        })

    return jsonify({"results": results[:10]})


@app.route("/api/translate", methods=["POST"])
def api_translate():
    """翻译（占位 — 后续接入 AI API）"""
    body = request.get_json()
    text = body.get("text", "")
    direction = body.get("direction", "zh-en")

    # 占位: 简单反转提示
    # 实际实现可调用 LLM API 或集成其他翻译工具
    return jsonify({
        "ok": True,
        "input": text,
        "direction": direction,
        "output": f"[Translation placeholder] {text}",
        "note": "Translation API not yet configured. This is a placeholder."
    })


# ============================================================
# 启动
# ============================================================
def open_browser(port: int):
    """延迟打开浏览器"""
    time.sleep(1.5)
    webbrowser.open(f"http://localhost:{port}")


def main():
    import argparse
    parser = argparse.ArgumentParser(description="IAV Workbench API Server")
    parser.add_argument("--port", type=int, default=8765, help="Server port (default: 8765)")
    parser.add_argument("--no-browser", action="store_true", help="Don't open browser")
    args = parser.parse_args()

    # 修复 Windows 控制台编码
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    print("=" * 60)
    print(f"  IAV Workbench API Server")
    print(f"  Port: {args.port}")
    print(f"  URL:  http://localhost:{args.port}")
    print(f"  Ctrl+C to stop")
    print("=" * 60)

    if not args.no_browser:
        threading.Thread(target=open_browser, args=(args.port,), daemon=True).start()

    app.run(host="127.0.0.1", port=args.port, debug=False)


if __name__ == "__main__":
    main()
