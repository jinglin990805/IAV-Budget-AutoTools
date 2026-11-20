# -*- coding: utf-8 -*-
"""
IAV 工作台 API 服务器 — 轻量 Flask HTTP 服务

提供:
- GET  /api/status          — 返回 PO/GR 状态 JSON
- GET  /api/cost/status     — 返回 Cost Tracking/PO 分配/付款闸门只读状态
- GET  /api/status/alerts   — 返回逾期/即将到期提醒
- POST /api/run/po_allocator — 运行 PO 分配器 V2
- POST /api/run/budget       — 运行 Budget 月度更新器
- POST /api/run/travel       — 运行差旅汇总
- GET  /api/outputs          — 列出 output/ 目录文件
- GET  /api/knowledge        — 搜索知识库
- POST /api/translate        — 翻译（占位）
- POST /api/po/parse-pdf     — 上传 PO PDF 自动解析 Level 配额
- POST /api/po/save-config   — 新增 PO 配置写入 Excel
- POST /api/po/preview       — 预览新 PO 规划结果
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
import copy
import shutil
import tempfile
import subprocess
import threading
import time
import webbrowser
import re
import uuid
import hashlib
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from collections import defaultdict

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
from fact_store import FactStore
from project_registry import build_profile
from review_scope import path_in_scope, review_scope
from timesheet_review_batches import classify_candidate, summarise_batches
try:
    from flask import Flask, jsonify, request, send_from_directory, send_file
except ImportError:
    print("[ERROR] Flask is required. Install: pip install flask")
    sys.exit(1)

# 路径配置
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
KNOWLEDGE_DIR = os.path.join(BASE_DIR, "knowledge-base")
STATUS_FILE = os.path.join(DATA_DIR, "po_gr_status.json")

app = Flask(__name__)

# 最近一次差旅审核生成的文件下载句柄。仅用于本地工作台会话，不把任意路径
# 暴露为公开下载接口。
TRAVEL_REVIEW_FILES = {}
TRAVEL_REVIEW_SESSIONS = {}
TRAVEL_PADDLE_AUTO_LOCK = threading.Lock()
TRAVEL_REVIEW_LOCK = threading.Lock()


def _travel_invoice_output_directory(template_path, default_output_dir, default_template_name):
    """Route the configured IAV Q3 travel-template family to its acceptance folder."""
    def family(path_or_name):
        stem = os.path.splitext(os.path.basename(str(path_or_name)))[0]
        return re.sub(r"_\d{6}(?:_v\d+)?$", "", stem, flags=re.IGNORECASE).casefold()

    if family(template_path) == family(default_template_name):
        return os.path.abspath(default_output_dir)
    return os.path.dirname(os.path.abspath(template_path)) or OUTPUT_DIR


@app.route("/api/timesheets/scan", methods=["POST"])
def api_timesheet_scan():
    """快速递归扫描 PDF 并登记候选；默认不在此步骤执行 OCR。"""
    body = request.get_json(silent=True) or {}
    project = body.get("project", "CNTS-25165")
    limit = body.get("limit")
    cmd = [sys.executable, os.path.join(BASE_DIR, "timesheet_ingest.py"), "scan",
           "--project", project, "--no-ocr", "--json"]
    if limit:
        cmd.extend(["--limit", str(int(limit))])
    try:
        result = subprocess.run(
            cmd, cwd=BASE_DIR, capture_output=True, text=True,
            encoding="utf-8", timeout=180,
            env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        )
        try:
            payload = json.loads(result.stdout)
        except json.JSONDecodeError:
            payload = {"stdout": result.stdout, "stderr": result.stderr}
        payload["ok"] = result.returncode == 0
        if result.stderr:
            payload["stderr"] = result.stderr[-2000:]
        return jsonify(payload), 200 if result.returncode == 0 else 500
    except subprocess.TimeoutExpired:
        return jsonify({"error": "Timesheet scan timed out"}), 504


@app.route("/api/timesheets/ocr", methods=["POST"])
def api_timesheet_ocr():
    """对当前审核月份 PDF 补做/重做本地 OCR。"""
    body = request.get_json(silent=True) or {}
    project = body.get("project", "CNTS-25165")
    limit = body.get("limit")
    cmd = [sys.executable, os.path.join(BASE_DIR, "timesheet_ingest.py"), "ocr-pending",
           "--project", project, "--current-month", "--json"]
    if limit:
        cmd.extend(["--limit", str(int(limit))])
    try:
        result = subprocess.run(
            cmd, cwd=BASE_DIR, capture_output=True, text=True,
            encoding="utf-8", timeout=900,
            env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        )
        try:
            payload = json.loads(result.stdout)
        except json.JSONDecodeError:
            payload = {"stdout": result.stdout, "stderr": result.stderr}
        payload["ok"] = result.returncode == 0
        if result.stderr:
            payload["stderr"] = result.stderr[-2000:]
        if not payload["ok"] and not payload.get("error"):
            errors = payload.get("errors") or []
            if errors:
                first = errors[0]
                payload["error"] = (
                    f"OCR 已处理 {payload.get('processed', 0)} 个文件，"
                    f"其中 {len(errors)} 个失败：{first.get('error', '未知错误')}"
                )
            else:
                payload["error"] = (result.stderr or result.stdout or "OCR 子进程失败").strip()[-1000:]
        return jsonify(payload), 200 if result.returncode == 0 else 500
    except subprocess.TimeoutExpired:
        return jsonify({"ok": False, "error": "OCR 超过 15 分钟仍未完成，请检查 PDF 或重试"}), 504
    except OSError as exc:
        return jsonify({"ok": False, "error": f"无法启动 OCR 子进程：{exc}"}), 500


@app.route("/api/reconcile/step2", methods=["POST"])
def api_step2_reconcile():
    """运行 PO Master 与 Monthly Closing 审核，不批准工时也不写回 Excel。"""
    body = request.get_json(silent=True) or {}
    project = body.get("project", "CNTS-25165")
    cmd = [sys.executable, os.path.join(BASE_DIR, "step2_reconcile.py"),
           "--project", project, "--json"]
    try:
        result = subprocess.run(
            cmd, cwd=BASE_DIR, capture_output=True, text=True,
            encoding="utf-8", timeout=300,
            env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        )
        try:
            payload = json.loads(result.stdout)
        except json.JSONDecodeError:
            payload = {"ok": False, "stdout": result.stdout, "stderr": result.stderr}
        payload["ok"] = result.returncode == 0 and payload.get("ok", True)
        if result.stderr:
            payload["stderr"] = result.stderr[-3000:]
        return jsonify(payload), 200 if result.returncode == 0 else 500
    except subprocess.TimeoutExpired:
        return jsonify({"ok": False, "error": "Step 2 reconciliation timed out"}), 504


@app.route("/api/timesheets/review")
def api_timesheet_review():
    """返回工时候选和待审核项。"""
    project = request.args.get("project", "CNTS-25165")
    include_resolved = request.args.get("include_resolved", "0") == "1"
    scope = review_scope(project)
    with FactStore() as store:
        store.initialize()
        where = "WHERE c.project_id=?" if include_resolved else "WHERE c.project_id=? AND c.review_status NOT IN ('APPROVED', 'SUPERSEDED', 'DISMISSED')"
        candidates = [dict(row) for row in store.connection.execute(
            f"""SELECT c.*, s.path AS source_path, s.sha256
                FROM timesheet_candidates c
                LEFT JOIN source_files s ON s.source_id=c.source_id
                {where} ORDER BY c.year DESC, c.month DESC, c.created_at""",
            (project,),
        )]
        # Historical months are reconciled from Budget Report + Monthly
        # Closing.  Only the latest closed month and its matching PDF folder
        # enter this review screen.
        candidates = [
            candidate for candidate in candidates
            if candidate.get("year") == scope["year"]
            and candidate.get("month") == scope["month_number"]
            and path_in_scope(candidate.get("source_path"), scope)
        ]
        for candidate in candidates:
            candidate["source_available"] = bool(
                candidate.get("source_path") and os.path.isfile(candidate["source_path"])
            )
            candidate["parse_json"] = json.loads(candidate.get("parse_json") or "{}")
            candidate["review_items"] = [dict(row) for row in store.connection.execute(
                """SELECT review_id, queue_type, severity, status, payload_json, created_at
                   FROM review_queue WHERE candidate_id=? ORDER BY created_at""",
                (candidate["candidate_id"],),
            )]
            for item in candidate["review_items"]:
                item["payload_json"] = json.loads(item.get("payload_json") or "{}")
        return jsonify({
            "project": project,
            "candidates": candidates,
            "pending_count": sum(1 for item in candidates if item["review_status"] not in {"APPROVED", "SUPERSEDED", "DISMISSED"}),
            "scope": scope,
        })


@app.route("/api/timesheets/review-options")
def api_timesheet_review_options():
    """Return person, Level and customer-day-rate choices for the review page."""
    project = request.args.get("project", "CNTS-25165")
    try:
        from person_rate_options import build_review_options

        with FactStore() as store:
            store.initialize()
            result = build_review_options(project, store.connection)
            result["scope"] = review_scope(project)
            return jsonify(result)
    except Exception as exc:
        return jsonify({"project": project, "people": [], "levels": [], "rates": [], "error": str(exc)}), 500


@app.route("/api/timesheets/review/batches")
def api_timesheet_review_batches():
    """Return read-only historical Timesheet/PO review batches.

    The existing ``/review`` endpoint intentionally serves only the current
    review month because that is the only scope allowed to create approved
    facts.  This endpoint exposes the historical backlog separately so it can
    be processed in controlled batches without widening the approval boundary.
    """
    project = request.args.get("project", "CNTS-25165")
    batch_filter = str(request.args.get("batch", "")).strip().upper()
    period_filter = str(request.args.get("period", "")).strip()
    include_resolved = request.args.get("include_resolved", "0") == "1"
    try:
        limit = max(1, min(int(request.args.get("limit", "100")), 500))
        offset = max(0, int(request.args.get("offset", "0")))
    except ValueError:
        return jsonify({"error": "limit/offset must be integers"}), 400

    if batch_filter and batch_filter not in {
        "VALIDATION_BLOCKED", "PO_MISSING", "PO_POTENTIAL", "PO_UNKNOWN",
        "VALIDATION_WARNING", "READY",
    }:
        return jsonify({"error": f"Unknown review batch: {batch_filter}"}), 400

    scope = review_scope(project)
    with FactStore() as store:
        store.initialize()
        status_clause = "" if include_resolved else " AND c.review_status NOT IN ('APPROVED', 'DISMISSED')"
        rows = [dict(row) for row in store.connection.execute(
            f"""SELECT c.*, s.path AS source_path, s.sha256
                FROM timesheet_candidates c
                LEFT JOIN source_files s ON s.source_id=c.source_id
                WHERE c.project_id=? AND c.review_status <> 'SUPERSEDED'{status_clause}
                ORDER BY c.year ASC, c.month ASC, c.created_at ASC, c.candidate_id ASC""",
            (project,),
        )]

        candidates = []
        excluded_current = 0
        for candidate in rows:
            is_current = (
                candidate.get("year") == scope["year"]
                and candidate.get("month") == scope["month_number"]
                and path_in_scope(candidate.get("source_path"), scope)
            )
            if is_current:
                excluded_current += 1
                continue

            review_items = [dict(item) for item in store.connection.execute(
                """SELECT review_id, queue_type, severity, status, payload_json, created_at,
                          resolved_at, resolution_note
                   FROM review_queue WHERE candidate_id=? ORDER BY created_at""",
                (candidate["candidate_id"],),
            )]
            for item in review_items:
                try:
                    item["payload_json"] = json.loads(item.get("payload_json") or "{}")
                except (TypeError, ValueError):
                    item["payload_json"] = {}
            classification = classify_candidate(candidate, review_items)
            if period_filter and classification["period"] != period_filter:
                continue
            if batch_filter and classification["batch"] != batch_filter:
                continue

            candidate["classification"] = classification
            candidate["source_available"] = bool(
                candidate.get("source_path") and os.path.isfile(candidate["source_path"])
            )
            candidate["review_items"] = review_items
            try:
                candidate["parse_json"] = json.loads(candidate.get("parse_json") or "{}")
            except (TypeError, ValueError):
                candidate["parse_json"] = {}
            candidates.append(candidate)

        # Build the selector from the complete historical set, not the active
        # filter, so changing batch/period does not make other counts vanish.
        all_items = []
        for candidate in rows:
            is_current = (
                candidate.get("year") == scope["year"]
                and candidate.get("month") == scope["month_number"]
                and path_in_scope(candidate.get("source_path"), scope)
            )
            if is_current:
                continue
            review_items = [dict(item) for item in store.connection.execute(
                "SELECT queue_type, severity, status, payload_json FROM review_queue WHERE candidate_id=?",
                (candidate["candidate_id"],),
            )]
            all_items.append({**candidate, **classify_candidate(candidate, review_items)})

        return jsonify({
            "project": project,
            "scope": scope,
            "read_only": True,
            "approval_note": "历史批次仅供复核；正式工时事实仍只能从当前审核月份显式确认产生。",
            "batches": summarise_batches(all_items),
            "total_historical": len(all_items),
            "current_scope_excluded": excluded_current,
            "pending_count": sum(1 for item in all_items if item.get("review_status") not in {"APPROVED", "DISMISSED"}),
            "filters": {"batch": batch_filter or None, "period": period_filter or None, "include_resolved": include_resolved},
            "offset": offset,
            "limit": limit,
            "items": candidates[offset:offset + limit],
            "has_more": offset + limit < len(candidates),
        })


def _safe_candidate_source(candidate_id: str):
    """Resolve a candidate PDF only inside the configured timesheet folder."""
    with FactStore() as store:
        store.initialize()
        row = store.connection.execute(
            """SELECT c.project_id, s.path, s.file_name FROM timesheet_candidates c
               JOIN source_files s ON s.source_id=c.source_id
               WHERE c.candidate_id=?""",
            (candidate_id,),
        ).fetchone()
    if not row:
        return None, "Candidate source not found"
    profile = build_profile(row["project_id"])
    root = os.path.abspath(profile.source_dirs.get("timesheet_pdf", ""))
    path = os.path.abspath(row["path"])
    try:
        inside = os.path.commonpath([root, path]) == root
    except ValueError:
        inside = False
    if not inside or not path.lower().endswith(".pdf"):
        return None, "Source file is outside the configured PDF directory"
    if not os.path.isfile(path):
        return None, f"Source PDF is missing: {row['file_name']}"
    return {"project": row["project_id"], "path": path, "file_name": row["file_name"]}, None


@app.route("/api/timesheets/source/preview/<candidate_id>")
def api_timesheet_source_preview(candidate_id):
    """Open a candidate PDF inline in the browser's PDF viewer."""
    source, error = _safe_candidate_source(candidate_id)
    if error:
        return jsonify({"error": error}), 404
    return send_file(source["path"], mimetype="application/pdf", as_attachment=False,
                     download_name=source["file_name"])


@app.route("/api/timesheets/source/open/<candidate_id>", methods=["POST"])
def api_timesheet_source_open(candidate_id):
    """Open the exact parent folder of the candidate PDF."""
    source, error = _safe_candidate_source(candidate_id)
    if error:
        return jsonify({"error": error}), 404
    try:
        if os.name != "nt":
            return jsonify({"error": "Explorer定位仅支持Windows，请使用PDF预览"}), 400
        # Use the Windows shell to open the verified folder in the interactive
        # desktop session. This is more reliable than starting explorer.exe as
        # a child process from the Flask service, which may return successfully
        # without showing a new Explorer window.
        folder = os.path.dirname(source["path"])
        opened_by = "shell"
        try:
            os.startfile(folder)
        except (AttributeError, OSError):
            opened_by = "explorer_fallback"
            subprocess.Popen(["explorer.exe", folder], close_fds=True)
        return jsonify({
            "ok": True, "path": source["path"], "folder": folder,
            "file_name": source["file_name"], "mode": "folder", "opened_by": opened_by,
        })
    except Exception as exc:
        return jsonify({"error": f"无法打开文件定位：{exc}"}), 500


def _review_person_match(options: dict, value: str | None) -> dict | None:
    key = re.sub(r"[^a-z0-9\u4e00-\u9fff]", "", str(value or "").lower())
    if not key:
        return None
    for person in options.get("people", []):
        values = [person.get("canonical_name", "")] + list(person.get("aliases") or [])
        if any(re.sub(r"[^a-z0-9\u4e00-\u9fff]", "", str(item).lower()) == key for item in values):
            return person
    return None


@app.route("/api/timesheets/review/bulk-approve", methods=["POST"])
def api_timesheet_bulk_approve():
    """Approve only fully auto-matched candidates in the current month scope."""
    body = request.get_json(silent=True) or {}
    project = body.get("project", "CNTS-25165")
    scope = review_scope(project)
    if body.get("month") and body["month"] != scope["month"]:
        return jsonify({"error": f"只能确认当前审核月份 {scope['month']}，历史月份由 Budget Report/Monthly Closing 负责"}), 400
    try:
        from person_rate_options import build_review_options

        with FactStore() as store:
            store.initialize()
            options = build_review_options(project, store.connection)
            profile = build_profile(project)
            rows = [dict(row) for row in store.connection.execute(
                """SELECT c.*, s.path AS source_path FROM timesheet_candidates c
                   LEFT JOIN source_files s ON s.source_id=c.source_id
                   WHERE c.project_id=? AND c.review_status NOT IN ('APPROVED', 'SUPERSEDED', 'DISMISSED')
                   ORDER BY c.candidate_id""",
                (project,),
            )]
            requested = set(body.get("candidate_ids") or [])
            ready = []
            skipped = []
            for candidate in rows:
                if requested and candidate["candidate_id"] not in requested:
                    continue
                if candidate.get("year") != scope["year"] or candidate.get("month") != scope["month_number"] or not path_in_scope(candidate.get("source_path"), scope):
                    # The default one-click action is already limited to the
                    # current scope; only report this reason when the caller
                    # explicitly requested a historical candidate.
                    if requested:
                        skipped.append({"candidate_id": candidate["candidate_id"], "reason": "OUT_OF_CURRENT_REVIEW_SCOPE"})
                    continue
                open_items = store.connection.execute(
                    "SELECT queue_type, severity FROM review_queue WHERE candidate_id=? AND status='OPEN'",
                    (candidate["candidate_id"],),
                ).fetchall()
                try:
                    parse_payload = json.loads(candidate.get("parse_json") or "{}")
                except (TypeError, ValueError):
                    parse_payload = {}
                validation = parse_payload.get("validation") or {}
                validation_status = str(validation.get("status", "")).upper()
                validation_issues = validation.get("issues") or []
                has_blocking_validation = (
                    validation_status == "BLOCKED"
                    or any(
                        str(issue.get("severity", "")).upper() == "BLOCKING"
                        for issue in validation_issues
                        if isinstance(issue, dict)
                    )
                )
                person = _review_person_match(options, candidate.get("canonical_name") or candidate.get("raw_name"))
                if (
                    open_items or has_blocking_validation
                    or validation_status != "PASS"
                    or not person or not person.get("level")
                    or person.get("customer_day_rate") is None
                    or not candidate.get("po_id") or candidate.get("total_hours") is None
                ):
                    skipped.append({
                        "candidate_id": candidate["candidate_id"],
                        "reason": "NEEDS_REVIEW",
                        "open_items": [item["queue_type"] for item in open_items],
                        "validation_status": validation_status or "UNKNOWN",
                    })
                    continue
                ready.append((candidate, person))

            if not body.get("confirm", False):
                return jsonify({
                    "ok": True, "dry_run": True, "scope": scope,
                    "ready_count": len(ready), "skipped_count": len(skipped), "skipped": skipped,
                })

            approved = []
            for candidate, person in ready:
                store.upsert_person_rate(
                    project, person["canonical_name"], person["canonical_name"],
                    [person["canonical_name"], candidate.get("raw_name") or ""],
                    person["level"], float(person["customer_day_rate"]), scope["month"],
                    profile.hours_per_day,
                )
                fact_id = store.save_approved_timesheet(
                    candidate["candidate_id"], project, candidate["run_id"],
                    person["canonical_name"], scope["month"], str(candidate["po_id"]),
                    float(candidate["total_hours"]), candidate.get("source_id"),
                    note="一键确认：人员、Level、客户日费率、PO 和工时均已自动匹配",
                )
                year, month = scope["month"].split("-", 1)
                store.connection.execute(
                    """UPDATE timesheet_candidates SET canonical_name=?, year=?, month=?, review_status='APPROVED'
                       WHERE candidate_id=?""",
                    (person["canonical_name"], int(year), int(month), candidate["candidate_id"]),
                )
                store.resolve_candidate_reviews(candidate["candidate_id"], "一键确认：自动匹配通过")
                approved.append({"candidate_id": candidate["candidate_id"], "fact_id": fact_id})
            return jsonify({
                "ok": True, "dry_run": False, "scope": scope,
                "approved_count": len(approved), "approved": approved,
                "skipped_count": len(skipped), "skipped": skipped,
            })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/timesheets/review/<candidate_id>/dismiss", methods=["POST"])
def api_timesheet_dismiss_candidate(candidate_id):
    """Remove a candidate from the review queue without deleting audit data."""
    body = request.get_json(silent=True) or {}
    project = body.get("project", "CNTS-25165")
    reason = str(body.get("reason") or "用户删除审核条目").strip()[:500]
    try:
        with FactStore() as store:
            store.initialize()
            row = store.connection.execute(
                """SELECT c.*, s.path AS source_path
                   FROM timesheet_candidates c
                   LEFT JOIN source_files s ON s.source_id=c.source_id
                   WHERE c.candidate_id=? AND c.project_id=?""",
                (candidate_id, project),
            ).fetchone()
            if not row:
                return jsonify({"error": "Candidate not found"}), 404
            if row["review_status"] in {"APPROVED", "DISMISSED"}:
                return jsonify({"error": f"该条目当前状态为 {row['review_status']}，不能重复删除"}), 409
            scope = review_scope(project)
            if (
                row["year"] != scope["year"] or row["month"] != scope["month_number"]
                or not path_in_scope(row["source_path"], scope)
            ):
                return jsonify({"error": f"只能处理当前审核范围 {scope['month']} 的 PDF 条目"}), 400
            store.dismiss_candidate(candidate_id, reason)
            return jsonify({"ok": True, "candidate_id": candidate_id, "status": "DISMISSED", "reason": reason})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/timesheets/review/<candidate_id>", methods=["POST"])
def api_timesheet_review_candidate(candidate_id):
    """确认候选：保存人员/费率并写入正式 timesheet_facts。"""
    body = request.get_json(silent=True) or {}
    required = ("project", "person_key", "month", "po_id", "hours", "level", "customer_day_rate")
    missing = [key for key in required if body.get(key) in (None, "")]
    if missing:
        return jsonify({"error": f"Missing required fields: {', '.join(missing)}"}), 400
    try:
        with FactStore() as store:
            store.initialize()
            row = store.connection.execute(
                "SELECT * FROM timesheet_candidates WHERE candidate_id=? AND project_id=?",
                (candidate_id, body["project"]),
            ).fetchone()
            if not row:
                return jsonify({"error": "Candidate not found"}), 404
            if row["review_status"] in {"APPROVED", "DISMISSED", "SUPERSEDED"}:
                return jsonify({"error": f"该条目当前状态为 {row['review_status']}，不能再次确认"}), 409
            scope = review_scope(body["project"])
            source = store.connection.execute(
                "SELECT path FROM source_files WHERE source_id=?", (row["source_id"],)
            ).fetchone()
            if body["month"] != scope["month"] or row["year"] != scope["year"] or row["month"] != scope["month_number"] or not path_in_scope(source["path"] if source else None, scope):
                return jsonify({"error": f"只能确认当前审核月份 {scope['month']} 且来源为 {scope['source_dir']} 的 PDF"}), 400
            profile = build_profile(body["project"])
            aliases = body.get("aliases") or [body["person_key"], row["raw_name"] or ""]
            store.upsert_person_rate(
                body["project"], body["person_key"], body["person_key"], aliases,
                body["level"], float(body["customer_day_rate"]), body["month"],
                profile.hours_per_day, body.get("internal_cost_day_rate"),
            )
            fact_id = store.save_approved_timesheet(
                candidate_id, body["project"], row["run_id"], body["person_key"],
                body["month"], str(body["po_id"]), float(body["hours"]),
                row["source_id"], note=body.get("note", ""),
            )
            year, month = body["month"].split("-", 1)
            store.connection.execute(
                "UPDATE timesheet_candidates SET canonical_name=?, year=?, month=?, po_id=?, total_hours=?, review_status='APPROVED' WHERE candidate_id=?",
                (body["person_key"], int(year), int(month), str(body["po_id"]), float(body["hours"]), candidate_id),
            )
            store.resolve_candidate_reviews(candidate_id, body.get("note", "用户确认"))
            return jsonify({"ok": True, "candidate_id": candidate_id, "fact_id": fact_id})
    except (TypeError, ValueError) as exc:
        return jsonify({"error": f"Invalid review values: {exc}"}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


# ============================================================
# CORS 支持 — 允许 workbench.html 从任意本地来源访问 API
# ============================================================
@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, DELETE, PATCH, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    response.headers["Access-Control-Max-Age"] = "3600"
    return response


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
    """服务 workbench.html（禁止浏览器缓存）"""
    html_path = os.path.join(BASE_DIR, "workbench.html")
    if os.path.exists(html_path):
        resp = send_file(html_path)
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        resp.headers["Pragma"] = "no-cache"
        resp.headers["Expires"] = "0"
        return resp
    return jsonify({"error": "workbench.html not found"}), 404


@app.route("/api/status")
def api_status():
    """返回完整 PO/GR 状态"""
    data = load_status()
    return jsonify(data)


@app.route("/api/cost/status")
def api_cost_status():
    """Return read-only Cost Tracking and payment-gate status."""
    project = request.args.get("project", "CNTS-25165")
    with FactStore() as store:
        store.initialize()
        facts = store.connection.execute(
            """SELECT approval_status, booking_status, payment_status,
                      COUNT(*) AS count, ROUND(SUM(amount), 2) AS amount
               FROM cost_facts
               WHERE project_id=?
               GROUP BY approval_status, booking_status, payment_status
               ORDER BY approval_status, booking_status, payment_status""",
            (project,),
        ).fetchall()
        allocations = store.connection.execute(
            """SELECT approval_status, booking_status, payment_status,
                      COUNT(*) AS count, ROUND(SUM(amount), 2) AS amount
               FROM cost_fact_po_allocations
               WHERE project_id=?
               GROUP BY approval_status, booking_status, payment_status
               ORDER BY approval_status, booking_status, payment_status""",
            (project,),
        ).fetchall()
        open_exceptions = store.connection.execute(
            """SELECT exception_type, COUNT(*) AS count
               FROM exceptions
               WHERE project_id=? AND status='OPEN'
               GROUP BY exception_type
               ORDER BY exception_type""",
            (project,),
        ).fetchall()
        exception_details = store.connection.execute(
            """SELECT exception_id, exception_type, target_key, message,
                      details_json, created_at
               FROM exceptions
               WHERE project_id=? AND status='OPEN'
                 AND exception_type LIKE 'COST_PO_%'
               ORDER BY exception_type, target_key""",
            (project,),
        ).fetchall()
        invoice_gate = store.connection.execute(
            """SELECT COUNT(*) AS target_count,
                      COUNT(DISTINCT invoice_ref) AS invoice_count
               FROM cost_facts
               WHERE project_id=? AND booking_status='POSTED'
                 AND payment_status='PENDING' AND invoice_ref IS NOT NULL""",
            (project,),
        ).fetchone()
        allocation_gate = store.connection.execute(
            """SELECT COUNT(*) AS target_count
               FROM cost_fact_po_allocations
               WHERE project_id=? AND booking_status='POSTED'
                 AND payment_status='PENDING' AND invoice_ref IS NOT NULL""",
            (project,),
        ).fetchone()
        invoice_rows = store.connection.execute(
            """SELECT invoice_ref, gr_ref, po_id, COUNT(*) AS count,
                      ROUND(SUM(amount), 2) AS amount, 'cost_fact' AS target_type
               FROM cost_facts
               WHERE project_id=? AND invoice_ref IS NOT NULL
               GROUP BY invoice_ref, gr_ref, po_id
               UNION ALL
               SELECT invoice_ref, gr_ref, po_id, COUNT(*) AS count,
                      ROUND(SUM(amount), 2) AS amount, 'allocation' AS target_type
               FROM cost_fact_po_allocations
               WHERE project_id=? AND invoice_ref IS NOT NULL
               GROUP BY invoice_ref, gr_ref, po_id
               ORDER BY invoice_ref, target_type, po_id""",
            (project, project),
        ).fetchall()
    exception_detail_payload = []
    for row in exception_details:
        try:
            details = json.loads(row["details_json"] or "{}")
        except (TypeError, json.JSONDecodeError):
            details = {}
        exception_detail_payload.append({
            "exception_id": row["exception_id"],
            "exception_type": row["exception_type"],
            "target_key": row["target_key"],
            "message": row["message"],
            "created_at": row["created_at"],
            "details": details,
        })
    return jsonify({
        "ok": True,
        "project": project,
        "cost_facts": [dict(row) for row in facts],
        "po_allocations": [dict(row) for row in allocations],
        "open_exceptions": [dict(row) for row in open_exceptions],
        "open_exception_details": exception_detail_payload,
        "payment_gate": {
            "registered_invoice_count": invoice_gate["invoice_count"],
            "cost_fact_target_count": invoice_gate["target_count"],
            "allocation_target_count": allocation_gate["target_count"],
            "total_target_count": invoice_gate["target_count"] + allocation_gate["target_count"],
            "write_enabled": False,
        },
        "invoice_targets": [dict(row) for row in invoice_rows],
    })


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
    """运行 Python 工具脚本，支持 mode 参数"""
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

    # 读取 mode 参数
    data = request.get_json(silent=True) or {}
    mode = data.get("mode", "report")
    project = data.get("project", "CNTS-25165")
    as_of = data.get("as_of", None)
    month = data.get("month", None)

    # 构建命令行参数
    cmd = [sys.executable, script_path, "--project", project]
    if mode:
        cmd.extend(["--mode", mode])
    if tool == "budget" and month:
        cmd.extend(["--month", str(month).strip()])
    elif as_of:
        cmd.extend(["--as-of", as_of])
    if data.get("generate_usage_status"):
        cmd.append("--generate-usage-status")

    try:
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        # Travel summary scans every month of source workbooks and routinely
        # runs ~3 minutes, so it needs a larger budget than the other tools.
        # Keep this in sync with the client-side timeout in runTool().
        tool_timeouts = {
            "travel": 600,
            "budget": 300 if mode == "write" else 120,
        }
        run_options = {
            "cwd": BASE_DIR,
            "capture_output": True,
            "text": True,
            "encoding": "utf-8",
            "timeout": tool_timeouts.get(tool, 120),
            "env": env,
        }
        # The monthly write button is an explicit user action. Confirm the
        # script's previewed write instead of letting non-interactive stdin
        # cause confirm_write() to cancel with EOFError.
        if tool == "budget" and mode == "write":
            run_options["input"] = "y\n"
        result = subprocess.run(cmd, **run_options)

        # 找最新生成的输出文件
        latest_outputs = []
        for f in sorted(glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx")), key=os.path.getmtime, reverse=True)[:3]:
            latest_outputs.append({
                "name": os.path.basename(f),
                "size": os.path.getsize(f),
                "mtime": datetime.fromtimestamp(os.path.getmtime(f)).strftime("%Y-%m-%d %H:%M"),
            })

        stdout = result.stdout[-3000:] if result.stdout else ""
        cancelled = "Write cancelled." in stdout
        written = "Monthly Actuals updated in:" in stdout
        write_warning = any(
            marker in stdout
            for marker in (
                "Original Budget Report could not be replaced",
                "Cannot stage an update beside the report",
            )
        )
        no_changes = ("No updates to apply." in stdout
                      or "No changes to apply." in stdout)

        return jsonify({
            "ok": result.returncode == 0 and not cancelled,
            "tool": tool,
            "mode": mode,
            "stdout": stdout,
            "stderr": result.stderr[-500:] if result.stderr else "",
            "written": written,
            "no_changes": no_changes,
            "warning": write_warning,
            "returncode": result.returncode,
            "latest_outputs": latest_outputs,
        })
    except subprocess.TimeoutExpired as e:
        stdout = e.stdout.decode("utf-8", errors="replace") if isinstance(e.stdout, bytes) else (e.stdout or "")
        stderr = e.stderr.decode("utf-8", errors="replace") if isinstance(e.stderr, bytes) else (e.stderr or "")
        timeout_seconds = 300 if tool == "budget" and mode == "write" else 120
        return jsonify({
            "ok": False,
            "error": f"Tool {tool} timed out ({timeout_seconds}s)",
            "stdout": stdout[-3000:],
            "stderr": stderr[-2000:],
        }), 504
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/wp/merge", methods=["POST"])
def api_wp_merge():
    """按季度 + PO + 人员合并审核后工时 PDF 到 GR WP 目录。

    body: {"quarter": "Q3", "dry_run": false}
    返回结构化结果：merged（已生成）、manual（含电子签/缺失，需手动）、errors。
    """
    from timesheet_wp_merger import run as wp_merge_run
    data = request.get_json(silent=True) or {}
    quarter = data.get("quarter", "Q3")
    dry_run = bool(data.get("dry_run", False))
    try:
        result = wp_merge_run(quarter=quarter, dry_run=dry_run)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    return jsonify(result)


@app.route("/api/budget/preview", methods=["POST"])
def api_budget_preview():
    """预览 Budget Report 差异（不修改文件）"""
    data = request.get_json(silent=True) or {}
    project = data.get("project", "CNTS-25165")
    as_of = data.get("as_of", None)

    script_path = os.path.join(BASE_DIR, "po_allocator_v2.py")
    cmd = [sys.executable, script_path, "--project", project, "--mode", "preview"]
    if as_of:
        cmd.extend(["--as-of", as_of])

    try:
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        result = subprocess.run(
            cmd, cwd=BASE_DIR, capture_output=True, text=True,
            encoding="utf-8", timeout=120, env=env,
        )
        stdout = result.stdout[-4000:] if result.stdout else ""
        stderr = result.stderr[-2000:] if result.stderr else ""

        # 提取 DIFF SUMMARY 部分
        diff_summary = ""
        in_diff = False
        for line in stdout.split("\n"):
            if "=== DIFF SUMMARY ===" in line:
                in_diff = True
            if in_diff:
                diff_summary += line + "\n"

        return jsonify({
            "ok": result.returncode == 0,
            "stdout": stdout,
            "stderr": stderr,
            "diff_summary": diff_summary.strip(),
            "has_changes": "No changes detected" not in diff_summary and diff_summary != "",
        })
    except subprocess.TimeoutExpired:
        return jsonify({"error": "Preview timed out (120s)"}), 504
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/budget/write", methods=["POST"])
def api_budget_write():
    """回写 Budget Report（三层安全：备份→预览→确认写入）"""
    data = request.get_json(silent=True) or {}
    project = data.get("project", "CNTS-25165")
    as_of = data.get("as_of", None)

    script_path = os.path.join(BASE_DIR, "po_allocator_v2.py")
    cmd = [sys.executable, script_path, "--project", project, "--mode", "write",
           "--generate-usage-status"]
    if as_of:
        cmd.extend(["--as-of", as_of])

    try:
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        # 通过 stdin 传入 'y' 自动确认（API 模式下无交互）
        result = subprocess.run(
            cmd, cwd=BASE_DIR, capture_output=True, text=True,
            encoding="utf-8", timeout=360, env=env,
            input="y\n",  # auto-confirm
        )
        stdout = result.stdout[-4000:] if result.stdout else ""
        stderr = result.stderr[-2000:] if result.stderr else ""
        written = "Budget Report updated:" in stdout or "Budget Report updated in:" in stdout
        no_changes = "No changes to apply." in stdout
        cancelled = "Write cancelled." in stdout
        warning = any(
            marker in stdout
            for marker in (
                "Original Budget Report could not be replaced",
                "Cannot stage an update beside the report",
            )
        )

        return jsonify({
            "ok": result.returncode == 0 and (written or no_changes) and not cancelled,
            "stdout": stdout,
            "stderr": stderr,
            "written": written,
            "no_changes": no_changes,
            "warning": warning,
            "returncode": result.returncode,
        })
    except subprocess.TimeoutExpired:
        return jsonify({"error": "Write timed out (360s)"}), 504
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/po/data-sources")
def api_po_data_sources():
    """扫描项目数据源状态"""
    project = request.args.get("project", "CNTS-25165")

    sys.path.insert(0, BASE_DIR)
    import config

    result = {"project": project, "sources": []}

    # 工时清单表 (Q1/Q2/Q3)
    gr_dir = getattr(config, 'GR_DOCUMENTS_DIR',
                     os.path.join(config.ROOT, project, "03_GR Documents"))
    vctc_dir = os.path.join(gr_dir, "2026", "VCTC")
    for quarter, fname in [("Q1", "IAV Q1 工时清单表.xlsx"),
                            ("Q2", "IAV Q2 工时清单表.xlsx"),
                            ("Q3", "IAV Q3 工时清单表.xlsx")]:
        fpath = os.path.join(vctc_dir, quarter, fname)
        exists = os.path.exists(fpath)
        result["sources"].append({"name": f"工时清单表 {quarter}", "exists": exists, "path": fpath})

    # 外援情况
    for q, fn in [("Q1", "AI IAV 外援情况_26Q1.xlsx"),
                   ("Q2", "Cost Breakdown_26Q2_260608.xlsx")]:
        fpath = os.path.join(vctc_dir, q, fn)
        exists = os.path.exists(fpath)
        result["sources"].append({"name": f"外援情况 {q}", "exists": exists, "path": fpath})

    # Budget Report
    project_dir = config.PROJECTS.get(project, {}).get("budget_report_dir", "")
    if project_dir and os.path.isdir(project_dir):
        import glob as _g
        br_files = _g.glob(os.path.join(project_dir, "IAV_Budget_Report_*.xlsx"))
        result["sources"].append({
            "name": "Budget Report",
            "exists": len(br_files) > 0,
            "count": len(br_files),
            "path": max(br_files, key=os.path.getmtime) if br_files else "",
        })

    # Timesheet
    ts_dir = os.path.join(os.path.dirname(project_dir), "01_IAV TMS & WP") if project_dir else ""
    ts_exists = os.path.isdir(ts_dir)
    result["sources"].append({"name": "Timesheet PDFs", "exists": ts_exists, "path": ts_dir})

    return jsonify(result)


@app.route("/api/budget/usage-status/allocated", methods=["POST"])
@app.route("/api/budget/usage-status", methods=["POST"])
def api_budget_usage_status():
    """按完整 PO 分配结果单独更新 PO Usage Status，不回写 Budget Report。"""
    data = request.get_json(silent=True) or {}
    project = data.get("project", "CNTS-25165")
    if project != "CNTS-25165":
        return jsonify({
            "ok": False,
            "error": f"PO Usage Status update is currently configured for CNTS-25165, not {project}.",
        }), 400

    as_of = str(data.get("as_of") or datetime.now().strftime("%Y-%m")).strip()
    if re.fullmatch(r"\d{6}", as_of):
        as_of = f"{as_of[:4]}-{as_of[4:]}"
    if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", as_of):
        return jsonify({"ok": False, "error": "as_of must be YYYY-MM or YYYYMM."}), 400

    script_path = os.path.join(BASE_DIR, "po_allocator_v2.py")
    cmd = [
        sys.executable, script_path,
        "--project", project,
        "--as-of", as_of,
        "--mode", "usage-status",
    ]
    if data.get("output"):
        cmd.extend(["--usage-status-output", data["output"]])

    try:
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        result = subprocess.run(
            cmd, cwd=BASE_DIR, capture_output=True, text=True,
            encoding="utf-8", timeout=300, env=env,
        )
        stdout = result.stdout[-4000:] if result.stdout else ""
        stderr_lines = [line.strip() for line in (result.stderr or "").splitlines() if line.strip()]
        used_local_fallback = "PO Usage Status generated locally; copy it to the project folder:" in stdout
        success = result.returncode == 0 and (
            "PO Usage Status updated:" in stdout or used_local_fallback
        )
        if success:
            return jsonify({
                "ok": True,
                "stdout": stdout,
                "stderr": "",
                "warning": (
                    "The project folder rejected the write. A complete file was generated in AutoTools/output; copy it to 02_Cost Overview."
                    if used_local_fallback else ""
                ),
                "returncode": result.returncode,
            })

        stdout_lines = [line.strip() for line in stdout.splitlines() if line.strip()]
        detail = stderr_lines[-1] if stderr_lines else (stdout_lines[-1] if stdout_lines else "No output from allocator")
        return jsonify({
            "ok": False,
            "error": f"PO Usage Status update failed: {detail}",
            "stdout": "",
            "stderr": "",
            "returncode": result.returncode,
        })
    except subprocess.TimeoutExpired:
        return jsonify({"ok": False, "error": "PO Usage Status update timed out (300s)"}), 504


@app.route("/api/budget/usage-status-working", methods=["POST"])
def api_budget_usage_status_working():
    """从现有模板生成保留公式/格式的 PO Usage Status Working 副本。"""
    data = request.get_json(silent=True) or {}
    project = data.get("project", "CNTS-25165")
    try:
        profile = build_profile(project)
        output = data.get("output") or os.path.join(
            BASE_DIR, "output", f"PO_Usage_Status_{project}_Working_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        from po_usage_status_generator import _is_customer_template, select_customer_template
        requested_template = data.get("template")
        if requested_template:
            template = os.path.abspath(requested_template)
            if not os.path.isfile(template):
                return jsonify({"ok": False, "error": f"Template not found: {template}"}), 400
            if not _is_customer_template(template):
                return jsonify({"ok": False, "error": f"Template does not contain the required customer formulas and PivotTables: {template}"}), 400
        else:
            template = select_customer_template(
                profile.budget_report_dir,
                output,
                getattr(profile, "po_usage_template_path", None) or None,
            )
        script_path = os.path.join(BASE_DIR, "po_usage_working.py")
        cmd = [sys.executable, script_path, "--project", project, "--template", template, "--output", output]
        result = subprocess.run(
            cmd, cwd=BASE_DIR, capture_output=True, text=True,
            encoding="utf-8", timeout=180,
            env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        )
        try:
            payload = json.loads(result.stdout.strip().splitlines()[-1])
        except (json.JSONDecodeError, IndexError):
            payload = {"stdout": result.stdout}
        payload["ok"] = result.returncode == 0 and payload.get("ok", True)
        if result.stderr:
            payload["stderr"] = result.stderr[-3000:]
        return jsonify(payload), 200 if result.returncode == 0 else 500
    except subprocess.TimeoutExpired:
        return jsonify({"ok": False, "error": "Formula-preserving PO Usage Status timed out (180s)"}), 504
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/budget/dashboard")
def api_budget_dashboard():
    """返回部门看板数据"""
    dept = request.args.get("dept", "TV")
    month = request.args.get("month", None)

    script_path = os.path.join(BASE_DIR, "budget_updater.py")
    cmd = [sys.executable, script_path, "--dept", dept]
    if month:
        cmd.extend(["--month", month])

    try:
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        result = subprocess.run(
            cmd, cwd=BASE_DIR, capture_output=True, text=True,
            encoding="utf-8", timeout=60, env=env,
        )
        stdout = result.stdout[-4000:] if result.stdout else ""

        # 解析输出格式: "code name: hours, ¥travel, active/total POs"
        projects = []
        for line in stdout.split("\n"):
            line = line.strip()
            if line.startswith("  ") and ":" in line and "h" in line:
                try:
                    parts = line.split(":", 1)
                    code_name = parts[0].strip().split(" ", 1)
                    code = code_name[0]
                    name = code_name[1] if len(code_name) > 1 else ""
                    stats = parts[1].split(",")
                    hours_str = stats[0].strip().rstrip("h") if len(stats) > 0 else "0"
                    travel_str = stats[1].strip().lstrip("¥").replace(",", "") if len(stats) > 1 else "0"
                    po_str = stats[2].strip().split(" ")[0] if len(stats) > 2 else "0/0"
                    po_parts = po_str.split("/")
                    projects.append({
                        "code": code, "name": name,
                        "hours": float(hours_str) if hours_str else 0,
                        "travel": float(travel_str) if travel_str else 0,
                        "po_active": int(po_parts[0]) if len(po_parts) > 0 else 0,
                        "po_count": int(po_parts[1]) if len(po_parts) > 1 else 0,
                    })
                except (ValueError, IndexError):
                    pass

        dept_line = next((l for l in stdout.split("\n") if "Department:" in l), "")
        return jsonify({"ok": True, "dept": dept, "projects": projects, "stdout": stdout})
    except subprocess.TimeoutExpired:
        return jsonify({"error": "Dashboard query timed out"}), 504
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
# PO 配置管理 API
# ============================================================

@app.route("/api/po/levels")
def api_po_levels():
    """返回 engineer level 配置 (Level ↔ 单价 映射)"""
    import openpyxl
    levels_path = os.path.join(BASE_DIR, "input", "engineer level.xlsx")
    result = {"customers": {}, "levels": [], "rates": []}
    if not os.path.exists(levels_path):
        return jsonify(result)

    try:
        wb = openpyxl.load_workbook(levels_path, data_only=True)
        ws = wb.active
        seen_levels = set()
        seen_rates = set()
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            if not row or len(row) < 3:
                continue
            customer = str(row[0]).strip() if row[0] else ""
            level = str(row[1]).strip() if row[1] else ""
            rate_str = str(row[2]).strip() if row[2] else ""
            if not level or not rate_str:
                continue
            try:
                rate = float(rate_str)
            except ValueError:
                continue
            key = f"{level}|{rate:.0f}"
            if key in seen_levels:
                continue
            seen_levels.add(key)
            result["levels"].append(level)
            result["rates"].append(f"{rate:.0f}")
            if customer:
                if customer not in result["customers"]:
                    result["customers"][customer] = []
                result["customers"][customer].append({"level": level, "unit_price": rate})
        wb.close()
    except Exception as e:
        result["_error"] = str(e)

    return jsonify(result)


@app.route("/api/po/projects")
def api_po_projects():
    """返回 Chai Jiaqi 负责的项目列表 (来自 Mapping.xlsx, 去重 Project Definition)"""
    import openpyxl
    mapping_path = os.path.join(
        BASE_DIR, "..", "..", "00_Project Auto Management Trial - Documents",
        "01_Master_Data", "Mapping.xlsx"
    )
    # 规范化路径
    mapping_path = os.path.normpath(mapping_path)
    if not os.path.exists(mapping_path):
        # fallback to OneDrive path
        mapping_path = os.path.join(
            os.path.expanduser("~"), "OneDrive - IAV Automotive Engineering(Shanghai) Co., Ltd",
            "00_Project Auto Management Trial - Documents", "01_Master_Data", "Mapping.xlsx"
        )
    projects = []
    if os.path.exists(mapping_path):
        try:
            wb = openpyxl.load_workbook(mapping_path, data_only=True)
            ws = wb.active
            seen = set()
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
                if len(row) < 6:
                    continue
                leader = str(row[1]).strip() if row[1] else ""
                proj_def = str(row[2]).strip() if row[2] else ""
                proj_name = str(row[5]).strip() if row[5] else ""
                if "Chai" not in leader:
                    continue
                if not proj_def or proj_def in seen:
                    continue
                seen.add(proj_def)
                projects.append({"code": proj_def, "name": proj_name})
            wb.close()
        except Exception as e:
            projects.append({"code": "_error", "name": str(e)})
    return jsonify({"projects": projects})


@app.route("/api/po/parse-pdf", methods=["POST"])

def api_po_parse_pdf():
    """上传客户 PO PDF，自动解析 Level 配额"""
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded. Use field 'file'."}), 400

    f = request.files['file']
    if f.filename == '':
        return jsonify({"error": "Empty filename"}), 400

    # 保存临时文件
    suffix = os.path.splitext(f.filename)[1] or '.pdf'
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=suffix)
    os.close(tmp_fd)
    f.save(tmp_path)

    try:
        sys.path.insert(0, BASE_DIR)
        from po_pdf_parser import parse_po_pdf, match_level

        info = parse_po_pdf(tmp_path)
        if not info or (not info.po_number and not info.level_quotas):
            return jsonify({
                "error": "This PDF does not appear to be a VCTC Call-off Purchase Order.",
                "hint": "Expected format: PO-VCTC00-xxxxxxxx with WP1/WP2 line items. You can manually enter PO details instead."
            }), 422

        quotas = []
        warnings = []
        for q in info.level_quotas:
            level = match_level(q.description)
            if not level:
                warnings.append(f"Could not match Level for: '{q.description}' — using raw description")
                level = q.description
            quotas.append({
                "description": q.description,
                "qty_days": q.qty_days,
                "unit_price": q.unit_price,
                "total_rmb": q.total_rmb,
                "level": level,
            })

        return jsonify({
            "ok": True,
            "po_number": info.po_number,
            "po_date": info.po_date,
            "total_excl_vat": info.total_excl_vat,
            "delivery_date": info.delivery_date,
            "level_quotas": quotas,
            "warnings": warnings,
        })

    except Exception as e:
        return jsonify({"error": f"Parse error: {str(e)}"}), 500
    finally:
        # 清理临时文件
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


@app.route("/api/travel/invoice-review", methods=["POST"])
def api_travel_invoice_review():
    """识别 PDF/DOCX 差旅材料，并复制模板后追加差旅报销明细。"""
    files = request.files.getlist("files") or request.files.getlist("file")
    if not files or all(not item.filename for item in files):
        return jsonify({"error": "请至少上传一个 PDF 或 DOCX 文件。"}), 400

    correction_upload = request.files.get("correction_manifest")
    correction_manifest = None
    if correction_upload and correction_upload.filename:
        if os.path.splitext(correction_upload.filename)[1].lower() != ".json":
            return jsonify({"error": "人工动作清单必须是 JSON 文件。"}), 400
        if sum(bool(item.filename) for item in files) != 1:
            return jsonify({"error": "应用人工动作清单时，请只上传一个差旅 PDF/DOCX 文件。"}), 400
        manifest_bytes = correction_upload.read(1_000_001)
        if len(manifest_bytes) > 1_000_000:
            return jsonify({"error": "人工动作清单不能超过 1 MB。"}), 400
        try:
            correction_manifest = json.loads(manifest_bytes.decode("utf-8-sig"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            return jsonify({"error": f"人工动作清单 JSON 无法读取：{exc}"}), 400

    claim_upload = request.files.get("claim_workbook")
    claim_workbook_path = None
    claim_temp_path = None
    claim_retained_for_session = False

    template_path = (request.form.get("template_path") or "").strip().strip('"')
    if not template_path:
        return jsonify({"error": "请填写差旅报销 Excel 模板路径。"}), 400
    output_date = (request.form.get("output_date") or "").strip() or None
    po_no = (request.form.get("po_no") or "").strip() or None
    policy_name = (request.form.get("policy_name") or "").strip()
    preview_only = (request.form.get("mode") or "").strip().lower() == "preview"
    default_year = None
    if output_date and re.fullmatch(r"\d{6}", output_date):
        try:
            default_year = datetime.strptime(output_date, "%y%m%d").year
        except ValueError:
            pass

    from travel_invoice_reviewer import (
        DEFAULT_TRAVEL_OUTPUT_DIR,
        DEFAULT_TRAVEL_TEMPLATE_NAME,
        _derive_output_path,
        assess_trip_delivery,
        apply_reviewed_travel_actions,
        create_review_workbook,
        load_travel_policies,
        parse_trip_file,
        preview_review_workbook,
    )
    travel_output_dir = _travel_invoice_output_directory(
        template_path, DEFAULT_TRAVEL_OUTPUT_DIR, DEFAULT_TRAVEL_TEMPLATE_NAME,
    )
    uses_configured_acceptance_dir = (
        os.path.normcase(os.path.abspath(travel_output_dir))
        == os.path.normcase(os.path.abspath(DEFAULT_TRAVEL_OUTPUT_DIR))
    )
    policy_config = load_travel_policies()
    policy_name = policy_name or policy_config.get("active_policy", "VGC")
    if policy_name not in policy_config.get("policies", {}):
        return jsonify({"error": f"差标不存在: {policy_name}"}), 400

    if claim_upload and claim_upload.filename:
        suffix = os.path.splitext(claim_upload.filename)[1].lower()
        if suffix not in {".xlsx", ".xlsm"}:
            return jsonify({"error": "报销明细必须是 .xlsx 或 .xlsm 文件。"}), 400
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        claim_name = re.sub(r'[<>:"/\\|?*]', "_", os.path.basename(claim_upload.filename))
        claim_temp_path = os.path.join(
            OUTPUT_DIR, f".travel-claim-{uuid.uuid4().hex}-{claim_name or 'claim.xlsx'}"
        )
        claim_upload.save(claim_temp_path)
        claim_workbook_path = claim_temp_path

    # Windows sandboxed Python installations can expose a read-only system
    # temp directory. Keep short-lived uploads as uniquely named files in the
    # workbench output directory instead of creating a nested temp directory.
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    temp_paths = []
    trips = []
    parse_errors = []
    saved_uploads = []
    input_coverage = []
    if not TRAVEL_REVIEW_LOCK.acquire(blocking=False):
        if claim_temp_path and not claim_retained_for_session:
            try:
                os.unlink(claim_temp_path)
            except OSError:
                pass
        return jsonify({
            "ok": False,
            "error": "已有差旅发票审核任务正在处理，请等待当前批次完成后再提交。",
        }), 409
    try:
        for uploaded in files:
            if not uploaded.filename:
                continue
            coverage_item = {"file": uploaded.filename, "status": "PROCESSING"}
            input_coverage.append(coverage_item)
            original_name = os.path.basename(uploaded.filename)
            # Keep Unicode filename hints (e.g. Chinese employee/city names)
            # while removing characters that are invalid in a Windows path.
            safe_name = re.sub(r'[<>:"/\\|?*]', "_", original_name) or "uploaded_file"
            suffix = os.path.splitext(safe_name)[1].lower()
            if suffix not in {".pdf", ".docx", ".doc"}:
                error = "仅支持 PDF、DOCX 或 DOC。"
                parse_errors.append({"file": uploaded.filename, "error": error})
                coverage_item.update({"status": "FAILED", "error": error})
                continue
            saved_path = os.path.join(OUTPUT_DIR, f".travel-review-{uuid.uuid4().hex}-{safe_name}")
            uploaded.save(saved_path)
            temp_paths.append(saved_path)
            saved_uploads.append((uploaded.filename, saved_path, coverage_item))

        auto_paddle_page_budget = {"remaining": 3}
        auto_paddle_budget_lock = threading.Lock()

        def parse_saved_upload(item):
            original_filename, saved_path, coverage_item = item
            trip = parse_trip_file(saved_path, default_year=default_year)
            digest = hashlib.sha256()
            with open(saved_path, "rb") as source_handle:
                for chunk in iter(lambda: source_handle.read(1024 * 1024), b""):
                    digest.update(chunk)
            source_sha256 = digest.hexdigest()
            coverage_item["source_sha256"] = source_sha256
            for invoice in trip.structured_invoices:
                invoice.source["sha256"] = source_sha256
            if correction_manifest is not None:
                coverage_item["manual_actions"] = apply_reviewed_travel_actions(
                    trip,
                    correction_manifest,
                    source_sha256,
                )
            if os.path.splitext(saved_path)[1].lower() == ".pdf":
                page_priority = {}
                urgent_statuses = {
                    "OCR_FAILED", "NO_TEXT", "PROJECTION_MISMATCH", "DUPLICATE_REVIEW",
                    "DIDI_AMOUNT_MISMATCH", "UNMATCHED_RECEIPT_CUE",
                }
                for page in trip.page_coverage:
                    try:
                        page_number = int(page.get("source_page", 0))
                    except (TypeError, ValueError):
                        continue
                    if page_number < 1:
                        continue
                    status = page.get("status")
                    supporting_only = page.get("supporting_only") is True
                    mixed_didi_receipt = (
                        page.get("has_didi_table") is True
                        and (
                            page.get("has_receipt_cue") is True
                            or status in {"UNMATCHED_INVOICE_CUE", "UNMATCHED_RECEIPT_CUE"}
                        )
                    )
                    page_risks = page.get("risk_flags") or []
                    has_specific_risk = any(
                        not isinstance(risk, dict)
                        or risk.get("code") not in {"UNMATCHED_INVOICE", "UNMATCHED_TAXI_RECEIPT"}
                        for risk in page_risks
                    )
                    has_generic_unmatched_risk = bool(page_risks) and not supporting_only
                    if mixed_didi_receipt or status in urgent_statuses or has_specific_risk or has_generic_unmatched_risk:
                        page_priority[page_number] = 0
                    elif status == "UNMATCHED_INVOICE_CUE" and not supporting_only:
                        page_priority[page_number] = 2
                    elif status == "UNCLASSIFIED_PAGE" and not supporting_only:
                        page_priority[page_number] = 3
                for invoice in trip.structured_invoices:
                    validation = getattr(invoice, "validation", {}) or {}
                    if validation.get("status") in {"WARNING", "BLOCKED"}:
                        try:
                            page_number = int((getattr(invoice, "source", {}) or {}).get("page", 0))
                        except (TypeError, ValueError):
                            continue
                        if page_number > 0:
                            issues = validation.get("issues", [])
                            amount_issue = any(
                                str(issue.get("field", "")).startswith(("total", "net", "tax", "line_items"))
                                or "AMOUNT" in str(issue.get("code", ""))
                                or "TOTAL" in str(issue.get("code", ""))
                                for issue in issues if isinstance(issue, dict)
                            )
                            priority = 1 if amount_issue else 4
                            page_priority[page_number] = min(page_priority.get(page_number, priority), priority)
                for expense in trip.expenses:
                    if getattr(expense, "category", "") == "taxi":
                        try:
                            page_number = int(getattr(expense, "source_page", 0) or 0)
                        except (TypeError, ValueError):
                            continue
                        if page_number > 0:
                            page_priority[page_number] = min(page_priority.get(page_number, 1), 1)

                ordered_pages = sorted(page_priority, key=lambda page: (page_priority[page], page))
                max_auto_paddle_pages = 3
                with auto_paddle_budget_lock:
                    selected_count = min(len(ordered_pages), auto_paddle_page_budget["remaining"])
                    selected_pages = ordered_pages[:selected_count]
                    auto_paddle_page_budget["remaining"] -= selected_count
                skipped_pages = ordered_pages[len(selected_pages):]
                auto_paddle = {
                    "status": "NO_TARGET_PAGES" if not ordered_pages else ("SKIPPED_BATCH_LIMIT" if not selected_pages else "PENDING"),
                    "engine": "PaddleOCR",
                    "model": "PP-OCRv6 medium det+rec",
                    "selected_pages": selected_pages,
                    "skipped_pages": skipped_pages,
                    "cache": {},
                    "elapsed_ms": 0,
                    "review_only": True,
                    "auto_applied": False,
                }
                for page in trip.page_coverage:
                    page_number = int(page.get("source_page", 0) or 0)
                    if page_number in skipped_pages:
                        page["paddle_review"] = {
                            "status": "SKIPPED_LIMIT",
                            "message": f"本次上传批次的自动复读上限为 {max_auto_paddle_pages} 页；此页仍保留原有风险提示。",
                            "channels": [],
                        }
                if selected_pages:
                    try:
                        from travel_invoice_reviewer import reocr_pdf_pages_with_paddle_for_review

                        # Paddle's local CPU model is memory-heavy; serialize automatic
                        # runs when several PDFs are submitted in one batch.
                        paddle_results = []
                        # The manual review endpoint accepts at most three pages per
                        # call, so run the automatic selection in small serialized batches.
                        with TRAVEL_PADDLE_AUTO_LOCK:
                            for offset in range(0, len(selected_pages), 3):
                                paddle_results.append(
                                    reocr_pdf_pages_with_paddle_for_review(
                                        saved_path, selected_pages[offset:offset + 3]
                                    )
                                )
                        paddle_result = {
                            "pages": [
                                page
                                for result in paddle_results
                                for page in result.get("pages", [])
                            ],
                            "cache": {
                                "hits": sum(int((result.get("cache") or {}).get("hits", 0)) for result in paddle_results),
                                "misses": sum(int((result.get("cache") or {}).get("misses", 0)) for result in paddle_results),
                            },
                            "elapsed_ms": sum(int(result.get("elapsed_ms", 0) or 0) for result in paddle_results),
                            "cache_warnings": [
                                warning
                                for result in paddle_results
                                for warning in result.get("cache_warnings", [])
                            ],
                        }
                        candidate_by_page = {
                            int(candidate.get("page", 0)): candidate
                            for candidate in paddle_result.get("pages", [])
                            if isinstance(candidate, dict)
                        }
                        for page in trip.page_coverage:
                            page_number = int(page.get("source_page", 0) or 0)
                            if page_number not in selected_pages:
                                continue
                            candidate = candidate_by_page.get(page_number, {})
                            channels = []
                            for channel_candidate in [candidate, *(candidate.get("alternate_candidates") or [])]:
                                if not isinstance(channel_candidate, dict):
                                    continue
                                channels.append({
                                    "channel": channel_candidate.get("channel", "RGB"),
                                    "preprocessing": channel_candidate.get("preprocessing", ""),
                                    "text": str(channel_candidate.get("text", ""))[:3000],
                                    "confidence": channel_candidate.get("confidence"),
                                    "elapsed_ms": channel_candidate.get("elapsed_ms"),
                                    "cache_status": channel_candidate.get("cache_status", ""),
                                    "amount_evidence": (channel_candidate.get("amount_evidence") or [])[:30],
                                })
                            page["paddle_review"] = {
                                "status": "FAILED" if candidate.get("error") else "COMPLETED",
                                "message": candidate.get("error", ""),
                                "channels": channels,
                            }
                            if candidate.get("error"):
                                continue

                            def labelled_totals(channel_candidate, allow_amount_label=False):
                                values = set()
                                for evidence in channel_candidate.get("amount_evidence", []) or []:
                                    field_type = evidence.get("field_type")
                                    labels = [str(label).strip().casefold() for label in evidence.get("labels", [])]
                                    is_receipt_amount = (
                                        allow_amount_label
                                        and field_type == "amount_label"
                                        and any(label in {"金额", "amount"} for label in labels)
                                    )
                                    if field_type != "total_or_paid" and not is_receipt_amount:
                                        continue
                                    if evidence.get("association") not in {"same_block", "same_row"}:
                                        continue
                                    try:
                                        confidence = float(evidence.get("confidence"))
                                        value = round(float(evidence.get("value_text")), 2)
                                    except (TypeError, ValueError):
                                        continue
                                    if confidence >= 0.90 and value > 0:
                                        values.add(value)
                                return values

                            red_candidate = next((
                                value for value in candidate.get("alternate_candidates", []) or []
                                if isinstance(value, dict) and value.get("channel") == "R"
                            ), None)
                            page_expenses = [
                                expense for expense in trip.expenses
                                if int(getattr(expense, "source_page", 0) or 0) == page_number
                            ]
                            allow_amount_label = (
                                page.get("has_didi_table") is True
                                or any(expense.category == "taxi" for expense in page_expenses)
                            )
                            rgb_totals = labelled_totals(candidate, allow_amount_label)
                            red_totals = labelled_totals(red_candidate or {}, allow_amount_label)
                            risk = None
                            if len(rgb_totals) == 1 and len(red_totals) == 1:
                                rgb_total = next(iter(rgb_totals))
                                red_total = next(iter(red_totals))
                                if abs(rgb_total - red_total) <= 0.01:
                                    paddle_total = rgb_total
                                    page["paddle_review"]["agreed_labelled_amount"] = paddle_total
                                    existing_amounts = [round(float(expense.amount), 2) for expense in page_expenses]
                                    if page_expenses and not any(abs(value - paddle_total) <= 0.01 for value in existing_amounts):
                                        current_text = "、".join(f"¥{value:.2f}" for value in existing_amounts)
                                        risk = {
                                            "code": "PADDLE_TESSERACT_AMOUNT_MISMATCH",
                                            "severity": "WARNING",
                                            "source_page": page_number,
                                            "amount_cny": paddle_total,
                                            "description": (
                                                f"PaddleOCR 原彩图和红色通道均将同一票面金额标签识别为 "
                                                f"¥{paddle_total:.2f}，与主识别记录 {current_text} 不一致；请按源票核定。"
                                            ),
                                        }
                                    elif not page_expenses:
                                        risk = {
                                            "code": "PADDLE_UNMATCHED_AMOUNT_CUE",
                                            "severity": "WARNING",
                                            "source_page": page_number,
                                            "amount_cny": paddle_total,
                                            "description": (
                                                f"PaddleOCR 两种通道均在同一票面金额标签旁识别到 ¥{paddle_total:.2f}，"
                                                "但主识别未生成费用记录；需确认票据类别和行程归属。"
                                            ),
                                        }
                                else:
                                    risk = {
                                        "code": "PADDLE_CHANNEL_AMOUNT_CONFLICT",
                                        "severity": "WARNING",
                                        "source_page": page_number,
                                        "description": (
                                            f"PaddleOCR 原彩图识别 ¥{rgb_total:.2f}、红色通道识别 "
                                            f"¥{red_total:.2f}；候选金额不一致，需核对源票。"
                                        ),
                                    }
                            if risk:
                                risk_flags = page.get("risk_flags")
                                if not isinstance(risk_flags, list):
                                    risk_flags = []
                                    page["risk_flags"] = risk_flags
                                risk_flags.append(risk)
                                page.setdefault("reasons", []).append(risk["description"])
                                page["review_required"] = True
                                page["paddle_review"]["risk_code"] = risk["code"]
                        page_errors = [
                            f"第 {int(page.get('source_page', 0) or 0)} 页：{page['paddle_review'].get('message', '复读失败')}"
                            for page in trip.page_coverage
                            if page.get("paddle_review", {}).get("status") == "FAILED"
                        ]
                        auto_paddle.update({
                            "status": "COMPLETED_WITH_WARNINGS" if page_errors else "COMPLETED",
                            "error": "；".join(page_errors[:3]),
                            "cache": paddle_result.get("cache", {}),
                            "elapsed_ms": paddle_result.get("elapsed_ms", 0),
                            "cache_warnings": paddle_result.get("cache_warnings", []),
                        })
                    except Exception as exc:
                        auto_paddle.update({
                            "status": "FAILED",
                            "error": f"{type(exc).__name__}: {exc}",
                        })
                        for page in trip.page_coverage:
                            if int(page.get("source_page", 0) or 0) in selected_pages:
                                page["paddle_review"] = {
                                    "status": "FAILED",
                                    "message": str(exc),
                                    "channels": [],
                                }
                coverage_item["paddle_review"] = auto_paddle
            # Parser uses the temporary path for reading, but the original
            # upload name is the useful audit label in the result.
            trip.source_file = original_filename
            for expense in trip.expenses:
                expense.source_file = original_filename
            for invoice in trip.structured_invoices:
                invoice.source["source_file"] = original_filename
                for evidence in invoice.evidence.values():
                    evidence["source_file"] = original_filename
            return trip

        # Preserve the workbench's eight-document batch capacity. Reject
        # overlapping review requests above so a retry in another tab cannot
        # start a second eight-document batch at the same time.
        if saved_uploads:
            worker_count = min(8, len(saved_uploads))
            with ThreadPoolExecutor(max_workers=worker_count) as executor:
                futures = [executor.submit(parse_saved_upload, item) for item in saved_uploads]
                for item, future in zip(saved_uploads, futures):
                    original_filename, _saved_path, coverage_item = item
                    try:
                        trip = future.result()
                        trips.append(trip)
                        coverage_item.update({
                            "status": (
                                "REVIEW_REQUIRED"
                                if assess_trip_delivery(trip)["review_required"]
                                else "PROCESSED"
                            ),
                            "pages": trip.page_coverage,
                            "invoice_validation": assess_trip_delivery(trip)["invoice_validation"],
                            "delivery_status": assess_trip_delivery(trip)["delivery_status"],
                            "delivery_issues": assess_trip_delivery(trip)["delivery_issues"],
                        })
                    except Exception as exc:
                        error = str(exc)
                        parse_errors.append({"file": original_filename, "error": error})
                        coverage_item.update({"status": "FAILED", "error": error})

        coverage_summary = {
            "files_submitted": len(input_coverage),
            "files_processed": sum(
                item.get("status") in {"PROCESSED", "REVIEW_REQUIRED"}
                for item in input_coverage
            ),
            "files_failed": sum(item.get("status") == "FAILED" for item in input_coverage),
            "pages_reported": sum(len(item.get("pages", [])) for item in input_coverage),
            "pages_requiring_review": sum(
                page.get("review_required") is True
                for item in input_coverage
                for page in item.get("pages", [])
            ),
            "invoices_requiring_review": sum(
                (item.get("invoice_validation") or {}).get("review_required", 0)
                for item in input_coverage
            ),
            "trips_requiring_review": sum(
                item.get("delivery_status") == "REVIEW_DRAFT"
                for item in input_coverage
            ),
        }
        partial_result = bool(coverage_summary["files_failed"] or coverage_summary["trips_requiring_review"])

        if not trips:
            return jsonify({
                "error": "没有成功识别任何文件。",
                "ok": False,
                "completion_status": "FAILED",
                "delivery_status": "FAILED",
                "ready_for_use": False,
                "partial_result_available": False,
                "parse_errors": parse_errors,
                "input_coverage": input_coverage,
                "coverage_summary": coverage_summary,
            }), 422

        storage_warning = None
        preferred_name = os.path.basename(_derive_output_path(template_path, output_date))

        def next_available_path(directory, filename):
            """Return a non-conflicting output path without overwriting a workbook."""
            candidate = os.path.join(directory, filename)
            if not os.path.exists(candidate):
                return candidate
            stem, ext = os.path.splitext(filename)
            version = 2
            while True:
                candidate = os.path.join(directory, f"{stem}_v{version}{ext}")
                if not os.path.exists(candidate):
                    return candidate
                version += 1

        # The configured IAV Q3 workbook family always publishes to the PPE
        # acceptance folder, even if its source template was selected elsewhere.
        # Other template families remain beside their selected source workbook.
        preferred_output_path = next_available_path(travel_output_dir, preferred_name)
        if preview_only:
            result = preview_review_workbook(
                template_path=template_path,
                trips=trips,
                output_date=output_date,
                po_no=po_no,
                policy_name=policy_name,
                claim_workbook_path=claim_workbook_path,
            )
            if (
                result.get("ok")
                and len(saved_uploads) == 1
                and os.path.splitext(saved_uploads[0][1])[1].lower() == ".pdf"
                and not parse_errors
            ):
                now = time.time()
                expired = [
                    token for token, session in TRAVEL_REVIEW_SESSIONS.items()
                    if now - session.get("created_at", 0) > 1800
                ]
                for token in expired:
                    expired_session = TRAVEL_REVIEW_SESSIONS.pop(token, None)
                    expired_claim = (expired_session or {}).get("claim_workbook_path")
                    if expired_claim:
                        try:
                            os.unlink(expired_claim)
                        except OSError:
                            pass
                while len(TRAVEL_REVIEW_SESSIONS) >= 10:
                    evicted = TRAVEL_REVIEW_SESSIONS.pop(next(iter(TRAVEL_REVIEW_SESSIONS)))
                    evicted_claim = (evicted or {}).get("claim_workbook_path")
                    if evicted_claim:
                        try:
                            os.unlink(evicted_claim)
                        except OSError:
                            pass
                review_token = uuid.uuid4().hex
                coverage_item = saved_uploads[0][2]
                TRAVEL_REVIEW_SESSIONS[review_token] = {
                    "created_at": now,
                    "source_sha256": coverage_item["source_sha256"],
                    "source_file": saved_uploads[0][0],
                    "trip": copy.deepcopy(trips[0]),
                    "template_path": template_path,
                    "output_date": output_date,
                    "po_no": po_no,
                    "policy_name": policy_name,
                    "claim_workbook_path": claim_workbook_path,
                }
                claim_retained_for_session = bool(claim_workbook_path)
                result["review_token"] = review_token
                result["correction_session_expires_in_seconds"] = 1800
            result.update({
                "parse_errors": parse_errors,
                "input_coverage": input_coverage,
                "coverage_summary": coverage_summary,
            })
            if not result.get("ok"):
                return jsonify(result), 422
        else:
            try:
                os.makedirs(travel_output_dir, exist_ok=True)
                result = create_review_workbook(
                    template_path=template_path,
                    trips=trips,
                    output_date=output_date,
                    output_path=preferred_output_path,
                    po_no=po_no,
                    policy_name=policy_name,
                    claim_workbook_path=claim_workbook_path,
                )
            except PermissionError:
                if uses_configured_acceptance_dir:
                    return jsonify({
                        "ok": False,
                        "completion_status": "FAILED",
                        "delivery_status": "FAILED",
                        "ready_for_use": False,
                        "partial_result_available": False,
                        "output_dir": travel_output_dir,
                        "output_path": preferred_output_path,
                        "error": (
                            "审核结果没有写入指定验收目录；请检查该目录的写入权限。"
                            f"目标目录：{travel_output_dir}"
                        ),
                        "parse_errors": parse_errors,
                        "input_coverage": input_coverage,
                        "coverage_summary": coverage_summary,
                    }), 500
                # If the template folder is read-only, keep the result in the
                # workbench's writable output folder for non-default template families.
                fallback_path = next_available_path(OUTPUT_DIR, preferred_name)
                result = create_review_workbook(
                    template_path=template_path,
                    trips=trips,
                    output_date=output_date,
                    output_path=fallback_path,
                    po_no=po_no,
                    policy_name=policy_name,
                    claim_workbook_path=claim_workbook_path,
                )
                storage_warning = "模板目录没有写权限，新表已另存到工作台 output；原模板未被修改。"
            token = uuid.uuid4().hex
            TRAVEL_REVIEW_FILES[token] = result["output_path"]
            # Keep this in-memory map bounded for a long-running local server.
            while len(TRAVEL_REVIEW_FILES) > 20:
                TRAVEL_REVIEW_FILES.pop(next(iter(TRAVEL_REVIEW_FILES)))
            result.update({
                "download_url": f"/api/travel/invoice-review/download/{token}",
            })
            result.update({
                "parse_errors": parse_errors,
                "input_coverage": input_coverage,
                "coverage_summary": coverage_summary,
            })
        # The generated workbook remains available for review even when the
        # delivery gate fails; only the delivery label changes.
        trip_results = result.get("trips", [])
        coverage_by_file = {item.get("file"): item for item in input_coverage}
        for trip_result in trip_results:
            coverage_item = coverage_by_file.get(trip_result.get("file"))
            if coverage_item is not None:
                coverage_item.update({
                    "status": "REVIEW_REQUIRED" if trip_result.get("review_required") else "PROCESSED",
                    "delivery_status": trip_result.get("delivery_status"),
                    "delivery_issues": trip_result.get("delivery_issues", []),
                    "invoice_validation": trip_result.get("invoice_validation", {}),
                })
        coverage_summary["invoices_requiring_review"] = sum(
            (item.get("invoice_validation") or {}).get("review_required", 0)
            for item in input_coverage
        )
        coverage_summary["trips_requiring_review"] = sum(
            item.get("delivery_status") == "REVIEW_DRAFT" for item in input_coverage
        )
        partial_result = bool(coverage_summary["files_failed"] or coverage_summary["trips_requiring_review"])
        result["completion_status"] = "PARTIAL" if partial_result else "COMPLETE"
        result["delivery_status"] = "REVIEW_DRAFT" if partial_result else "READY_FOR_USE"
        result["ready_for_use"] = not partial_result
        result["delivery_issues"] = [
            {"file": item.get("file"), **issue}
            for item in input_coverage
            for issue in item.get("delivery_issues", [])
        ]
        result["partial_result_available"] = True
        result["ok"] = not partial_result
        if storage_warning:
            result["storage_warning"] = storage_warning
        return jsonify(result), (207 if partial_result else 200)
    except Exception as exc:
        return jsonify({
            "error": str(exc),
            "completion_status": "FAILED",
            "partial_result_available": False,
            "parse_errors": parse_errors,
            "input_coverage": input_coverage,
            "coverage_summary": locals().get("coverage_summary"),
        }), 500
    finally:
        for temp_path in temp_paths:
            try:
                os.unlink(temp_path)
            except OSError:
                pass
        if claim_temp_path and not claim_retained_for_session:
            try:
                os.unlink(claim_temp_path)
            except OSError:
                pass
        TRAVEL_REVIEW_LOCK.release()


@app.route("/api/travel/invoice-review/reocr", methods=["POST"])
def api_travel_invoice_reocr():
    """Re-read up to three selected PDF pages and return review-only evidence."""
    uploaded = request.files.get("file")
    if uploaded is None or not uploaded.filename:
        return jsonify({"ok": False, "error": "请选择一个 PDF 文件。"}), 400
    if os.path.splitext(uploaded.filename)[1].lower() != ".pdf":
        return jsonify({"ok": False, "error": "局部重识别仅支持 PDF 文件。"}), 400

    raw_pages = request.form.get("pages", "")
    page_tokens = [value.strip() for value in raw_pages.split(",") if value.strip()]
    if not page_tokens or any(not re.fullmatch(r"\d+", value) for value in page_tokens):
        return jsonify({"ok": False, "error": "请输入页码，例如 5,8,10。"}), 400
    page_numbers = list(dict.fromkeys(int(value) for value in page_tokens))
    if len(page_numbers) > 3:
        return jsonify({"ok": False, "error": "一次最多重识别 3 页。"}), 400
    if any(value < 1 for value in page_numbers):
        return jsonify({"ok": False, "error": "页码必须从 1 开始。"}), 400

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    safe_name = re.sub(r'[<>:"/\\|?*]', "_", os.path.basename(uploaded.filename)) or "travel.pdf"
    saved_path = os.path.join(OUTPUT_DIR, f".travel-reocr-{uuid.uuid4().hex}-{safe_name}")
    try:
        uploaded.save(saved_path)
        if os.path.getsize(saved_path) > 100 * 1024 * 1024:
            return jsonify({"ok": False, "error": "PDF 不能超过 100 MB。"}), 413

        from travel_invoice_reviewer import reocr_pdf_pages_for_review

        result = reocr_pdf_pages_for_review(saved_path, page_numbers)
        if request.form.get("paddle") == "1":
            from travel_invoice_reviewer import reocr_pdf_pages_with_paddle_for_review

            paddle_result = reocr_pdf_pages_with_paddle_for_review(saved_path, page_numbers)
            tesseract_pages = {
                int(page.get("page")): page
                for page in result.get("pages", [])
                if isinstance(page, dict) and page.get("page") is not None
            }
            for paddle_page in paddle_result.get("pages", []):
                if not isinstance(paddle_page, dict):
                    continue
                page = tesseract_pages.get(int(paddle_page.get("page", 0)))
                if page is None:
                    continue
                candidate = paddle_page.get("candidate")
                if candidate is None and paddle_page.get("method"):
                    candidate = paddle_page
                if isinstance(candidate, dict):
                    page.setdefault("candidates", []).append(candidate)
                    alternate_candidates = candidate.get("alternate_candidates", [])
                    if isinstance(alternate_candidates, list):
                        page["candidates"].extend(
                            alternate for alternate in alternate_candidates
                            if isinstance(alternate, dict)
                        )
                elif paddle_page.get("error"):
                    page.setdefault("warnings", []).append(
                        f"PaddleOCR: {paddle_page['error']}"
                    )
            result["paddle"] = paddle_result
        return jsonify({"ok": True, "file": uploaded.filename, **result})
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except RuntimeError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 503
    except Exception as exc:
        return jsonify({"ok": False, "error": f"局部 OCR 失败：{type(exc).__name__}: {exc}"}), 500
    finally:
        try:
            os.unlink(saved_path)
        except OSError:
            pass


@app.route("/api/travel/invoice-review/apply-corrections", methods=["POST"])
def api_travel_invoice_apply_corrections():
    """Apply reviewer-confirmed fields to an in-memory preview and export it."""
    body = request.get_json(silent=True) or {}
    review_token = str(body.get("review_token", "")).strip()
    session = TRAVEL_REVIEW_SESSIONS.get(review_token)
    if not session:
        return jsonify({"ok": False, "error": "复核会话不存在或已失效，请重新预览一次。"}), 404
    if time.time() - session.get("created_at", 0) > 1800:
        TRAVEL_REVIEW_SESSIONS.pop(review_token, None)
        return jsonify({"ok": False, "error": "复核会话已超过 30 分钟，请重新预览一次。"}), 410

    corrections = body.get("corrections")
    if not isinstance(corrections, list) or not corrections:
        return jsonify({"ok": False, "error": "没有可应用的人工字段修正。"}), 400
    record_ids = [str(item.get("record_id", "")).strip() for item in corrections if isinstance(item, dict)]
    if len(record_ids) != len(corrections) or not all(record_ids) or len(set(record_ids)) != len(record_ids):
        return jsonify({"ok": False, "error": "修正记录 ID 缺失或重复。"}), 400

    from travel_invoice_reviewer import (
        DEFAULT_TRAVEL_OUTPUT_DIR,
        DEFAULT_TRAVEL_TEMPLATE_NAME,
        _derive_output_path,
        apply_manual_travel_corrections,
        assess_trip_delivery,
        create_review_workbook,
    )

    travel_output_dir = _travel_invoice_output_directory(
        session["template_path"], DEFAULT_TRAVEL_OUTPUT_DIR, DEFAULT_TRAVEL_TEMPLATE_NAME,
    )
    uses_configured_acceptance_dir = (
        os.path.normcase(os.path.abspath(travel_output_dir))
        == os.path.normcase(os.path.abspath(DEFAULT_TRAVEL_OUTPUT_DIR))
    )

    trip = copy.deepcopy(session["trip"])
    reviewed_on = datetime.now().date().isoformat()
    manifest = {
        "schema_version": 1,
        "review_status": "approved",
        "review_method": "workbench_manual_source_review",
        "reviewed_on": reviewed_on,
        "source_file": session["source_file"],
        "source_sha256": session["source_sha256"],
        "actions": [],
        "field_corrections": corrections,
    }
    try:
        applied = apply_manual_travel_corrections(
            trip,
            {
                "source_sha256": session["source_sha256"],
                "reviewed_on": reviewed_on,
                "corrections": corrections,
            },
        )
    except (TypeError, ValueError) as exc:
        return jsonify({"ok": False, "error": f"修正未应用：{exc}"}), 400

    preferred_name = os.path.basename(
        _derive_output_path(session["template_path"], session.get("output_date"))
    )

    def next_available_path(directory, filename):
        candidate = os.path.join(directory, filename)
        if not os.path.exists(candidate):
            return candidate
        stem, extension = os.path.splitext(filename)
        version = 2
        while True:
            candidate = os.path.join(directory, f"{stem}_v{version}{extension}")
            if not os.path.exists(candidate):
                return candidate
            version += 1

    preferred_output = next_available_path(travel_output_dir, preferred_name)
    storage_warning = None
    try:
        os.makedirs(travel_output_dir, exist_ok=True)
        result = create_review_workbook(
            template_path=session["template_path"],
            trips=[trip],
            output_date=session.get("output_date"),
            output_path=preferred_output,
            po_no=session.get("po_no"),
            policy_name=session["policy_name"],
            claim_workbook_path=session.get("claim_workbook_path"),
        )
    except PermissionError:
        if uses_configured_acceptance_dir:
            return jsonify({
                "ok": False,
                "completion_status": "FAILED",
                "delivery_status": "FAILED",
                "ready_for_use": False,
                "partial_result_available": False,
                "output_dir": travel_output_dir,
                "output_path": preferred_output,
                "error": (
                    "修正已审核，但结果没有写入指定验收目录；请检查该目录的写入权限。"
                    f"目标目录：{travel_output_dir}"
                ),
            }), 500
        fallback_output = next_available_path(OUTPUT_DIR, preferred_name)
        result = create_review_workbook(
            template_path=session["template_path"],
            trips=[trip],
            output_date=session.get("output_date"),
            output_path=fallback_output,
            po_no=session.get("po_no"),
            policy_name=session["policy_name"],
            claim_workbook_path=session.get("claim_workbook_path"),
        )
        storage_warning = "模板目录没有写权限，新工作簿已另存到工作台 output；原模板未修改。"

    output_digest = hashlib.sha256()
    with open(result["output_path"], "rb") as output_handle:
        for chunk in iter(lambda: output_handle.read(1024 * 1024), b""):
            output_digest.update(chunk)
    audit = {
        **manifest,
        "applied_corrections": applied,
        "output_file": result["output_name"],
        "output_sha256": output_digest.hexdigest(),
        "trip_summary": {
            "name": trip.name,
            "city": trip.city,
            "expense_count": len(trip.expenses),
            "totals": trip.totals,
            "delivery_status": (result.get("trips") or [{}])[0].get("delivery_status"),
        },
    }
    audit_name = os.path.splitext(result["output_name"])[0] + "_review.json"
    audit_directory = os.path.dirname(result["output_path"])
    audit_path = next_available_path(audit_directory, audit_name)
    os.makedirs(audit_directory, exist_ok=True)
    with open(audit_path, "w", encoding="utf-8") as audit_handle:
        json.dump(audit, audit_handle, ensure_ascii=False, indent=2)
        audit_handle.write("\n")

    download_token = uuid.uuid4().hex
    TRAVEL_REVIEW_FILES[download_token] = result["output_path"]
    audit_download_token = uuid.uuid4().hex
    TRAVEL_REVIEW_FILES[audit_download_token] = audit_path
    while len(TRAVEL_REVIEW_FILES) > 20:
        TRAVEL_REVIEW_FILES.pop(next(iter(TRAVEL_REVIEW_FILES)))
    TRAVEL_REVIEW_SESSIONS.pop(review_token, None)
    assessment = assess_trip_delivery(trip)
    result.update({
        "ok": assessment["ready_for_use"],
        "completion_status": "COMPLETE" if assessment["ready_for_use"] else "PARTIAL",
        "delivery_status": assessment["delivery_status"],
        "ready_for_use": assessment["ready_for_use"],
        "partial_result_available": True,
        "applied_corrections": applied,
        "source_sha256": session["source_sha256"],
        "review_audit_path": audit_path,
        "download_url": f"/api/travel/invoice-review/download/{download_token}",
        "review_audit_url": f"/api/travel/invoice-review/download/{audit_download_token}",
    })
    if storage_warning:
        result["storage_warning"] = storage_warning
    return jsonify(result), 200 if assessment["ready_for_use"] else 207


@app.route("/api/travel/standards", methods=["GET", "POST"])
def api_travel_standards():
    """Read or update editable customer travel-standard policies."""
    from travel_invoice_reviewer import load_travel_policies, save_travel_policies

    if request.method == "GET":
        return jsonify(load_travel_policies())
    try:
        saved = save_travel_policies(request.get_json(silent=True) or {})
        return jsonify({"ok": True, **saved})
    except (ValueError, OSError) as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/api/travel/invoice-template-default")
def api_travel_invoice_template_default():
    """Return the newest dated workbook as the next invoice-review template."""
    try:
        from travel_invoice_reviewer import (
            DEFAULT_TRAVEL_OUTPUT_DIR,
            DEFAULT_TRAVEL_TEMPLATE_NAME,
            find_latest_travel_template,
        )

        template_path = find_latest_travel_template()
        if not template_path:
            template_path = os.path.join(DEFAULT_TRAVEL_OUTPUT_DIR, DEFAULT_TRAVEL_TEMPLATE_NAME)
        return jsonify({
            "ok": True,
            "template_path": template_path,
            "output_dir": DEFAULT_TRAVEL_OUTPUT_DIR,
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/travel/invoice-review/download/<token>")
def api_travel_invoice_review_download(token):
    """下载本地工作台最近生成的差旅审核表。"""
    path = TRAVEL_REVIEW_FILES.get(token)
    if not path or not os.path.isfile(path):
        return jsonify({"error": "审核输出文件不存在或下载句柄已过期。"}), 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.route("/api/po/save-config", methods=["POST"])
def api_po_save_config():
    """保存新 PO 配置到 input/po_config_v2.xlsx"""
    body = request.get_json()
    if not body:
        return jsonify({"error": "Request body is required"}), 400

    po_id = body.get("po_id", "").strip()
    project = body.get("project", "CNTS-25165").strip()
    allocation_mode = body.get("allocation_mode", "usage")
    start_month = body.get("start_month", "")
    end_month = body.get("end_month", "")
    priority = body.get("priority", 1)
    levels = body.get("levels", [])

    if not po_id:
        return jsonify({"error": "PO ID is required"}), 400
    if not levels:
        return jsonify({"error": "At least one Level is required"}), 400

    # 读现有配置
    config_path = os.path.join(BASE_DIR, "input", "po_config_v2.xlsx")
    existing_rows = []
    existing_headers = []
    has_allocation_mode_col = False

    if os.path.exists(config_path):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(config_path)
            if "PO_Level_Config" in wb.sheetnames:
                ws = wb["PO_Level_Config"]
                # 读表头
                for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
                    existing_headers = [str(c) if c else "" for c in row]
                has_allocation_mode_col = any(
                    'allocation' in h.lower() and 'mode' in h.lower()
                    for h in existing_headers
                )
                # 读现有数据行
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
                    existing_rows.append(list(row))
            wb.close()
        except Exception:
            pass

    # 写回
    import xlsxwriter
    temp_path = config_path + ".tmp"

    # 统一使用新格式 headers (包含 Project 和 Allocation_Mode)
    new_headers = ["Project", "PO_ID", "Level", "Days_Quota", "Unit_Price",
                   "Value_RMB", "Start_Month", "End_Month", "Priority", "Allocation_Mode"]

    # 迁移旧数据 (无 Project 列 → 默认 CNTS-25165)
    has_project_col = existing_headers and 'project' in str(existing_headers[0]).lower()
    migrated_rows = []
    for row in existing_rows:
        if len(row) < 6:
            continue
        if has_project_col:
            migrated_rows.append(list(row) + [""] * (len(new_headers) - len(row)))
        else:
            # 旧格式: PO_ID 是 col 0，前面插入 Project = CNTS-25165
            new_row = ["CNTS-25165"] + list(row)
            # 补齐缺失列
            while len(new_row) < len(new_headers):
                new_row.append("")
            migrated_rows.append(new_row)

    wb = xlsxwriter.Workbook(temp_path)
    ws = wb.add_worksheet("PO_Level_Config")
    fmt_header = wb.add_format({
        'bold': True, 'bg_color': '#2F5496', 'font_color': 'white',
        'border': 1, 'align': 'center', 'valign': 'vcenter',
    })
    fmt_cell = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    fmt_num = wb.add_format({'border': 1, 'align': 'center', 'num_format': '#,##0'})

    for i, h in enumerate(new_headers):
        ws.write(0, i, h, fmt_header)

    for r, row_data in enumerate(migrated_rows, 1):
        for c in range(len(new_headers)):
            val = row_data[c] if c < len(row_data) else ""
            if c == 3 or c == 4 or c == 5:  # numeric columns
                if isinstance(val, (int, float)) and val != 0:
                    ws.write(r, c, val, fmt_num)
                else:
                    ws.write(r, c, str(val) if val else "", fmt_cell)
            else:
                # Last col = allocation_mode, default to usage
                if c == 9 and not val:
                    ws.write(r, c, "usage", fmt_cell)
                else:
                    ws.write(r, c, str(val) if val else "", fmt_cell)

    # 写新行
    rows_added = 0
    start_row = len(migrated_rows) + 1
    for i, lv in enumerate(levels):
        r = start_row + i
        ws.write(r, 0, project, fmt_cell)
        ws.write(r, 1, po_id, fmt_cell)
        ws.write(r, 2, lv.get("level", ""), fmt_cell)
        ws.write(r, 3, lv.get("days_quota", 0), fmt_num)
        ws.write(r, 4, lv.get("unit_price", 0), fmt_num)
        ws.write(r, 5, lv.get("value_rmb", lv.get("days_quota", 0) * lv.get("unit_price", 0)), fmt_num)
        ws.write(r, 6, start_month, fmt_cell)
        ws.write(r, 7, end_month, fmt_cell)
        ws.write(r, 8, priority, fmt_cell)
        ws.write(r, 9, allocation_mode, fmt_cell)
        rows_added += 1

    for c in range(len(new_headers)):
        ws.set_column(c, c, max(14, len(new_headers[c]) + 4))

    wb.close()

    os.makedirs(os.path.dirname(config_path), exist_ok=True)
    if os.path.exists(config_path):
        os.replace(temp_path, config_path)
    else:
        shutil.move(temp_path, config_path)

    return jsonify({
        "ok": True,
        "po_id": po_id,
        "project": project,
        "rows_added": rows_added,
        "allocation_mode": allocation_mode,
    })


def _load_po_configs_for_preview():
    """临时加载 PO 配置用于预览（不依赖 Excel）"""
    sys.path.insert(0, BASE_DIR)
    import config
    from po_allocator_v2 import (
        POLevelConfig, load_po_level_config,
        load_gongshi_qingdan, load_waiyuan_qingkuang,
        load_budget_report_actuals_and_forecast, load_timesheet_days,
        merge_ground_truth_allocation,
        build_po_level_inventory,
        calculate_level_consumption_rates,
        run_forward_projection,
        resolve_po_start_end_gaps,
    )

    personnel = config.CNTS25165_PERSONNEL
    po_configs = load_po_level_config("CNTS-25165")

    # 加载实际数据
    gongshi = load_gongshi_qingdan(personnel=personnel)
    waiyuan = load_waiyuan_qingkuang(personnel=personnel)
    project_dir = config.PROJECTS.get("CNTS-25165", {}).get("budget_report_dir", "")
    budget_actuals, _ = load_budget_report_actuals_and_forecast(project_dir, personnel)
    ts_days = load_timesheet_days(project_dir, personnel)

    merged = merge_ground_truth_allocation(
        gongshi, waiyuan, budget_actuals, ts_days, personnel, po_configs
    )
    return po_configs, merged, personnel


@app.route("/api/po/preview", methods=["POST"])
def api_po_preview():
    """预览新 PO 的规划结果（不写入 Excel）"""
    body = request.get_json()
    if not body:
        return jsonify({"error": "Request body is required"}), 400

    try:
        sys.path.insert(0, BASE_DIR)

        # 重定向 stdout 避免 po_allocator_v2 的 print 编码错误
        old_stdout = sys.stdout
        sys.stdout = io.StringIO()

        try:
            from po_allocator_v2 import (
                POLevelConfig, load_po_level_config,
                load_gongshi_qingdan, load_waiyuan_qingkuang,
                load_budget_report_actuals_and_forecast, load_timesheet_days,
                merge_ground_truth_allocation,
                build_po_level_inventory,
                calculate_level_consumption_rates,
                run_forward_projection, detect_transitions,
                resolve_po_start_end_gaps,
            )
            import config

            po_id = body.get("po_id", "").strip()
            project = body.get("project", "CNTS-25165").strip()
            allocation_mode = body.get("allocation_mode", "usage")
            start_month = body.get("start_month", "")
            end_month = body.get("end_month", "")
            priority = body.get("priority", 1)
            levels = body.get("levels", [])

            if not po_id or not levels:
                return jsonify({"error": "PO ID and levels are required"}), 400

            base_configs = load_po_level_config(project)
            temp_configs = list(base_configs)

            for lv in levels:
                temp_configs.append(POLevelConfig(
                    project=project,
                    po_id=po_id,
                    level=lv.get("level", ""),
                    days_quota=lv.get("days_quota", 0),
                    unit_price=lv.get("unit_price", 0),
                    value_rmb=lv.get("value_rmb", lv.get("days_quota", 0) * lv.get("unit_price", 0)),
                    start_month=start_month,
                    end_month=end_month,
                    priority=priority,
                ))
                temp_configs[-1].allocation_mode = allocation_mode

            temp_configs, _ = resolve_po_start_end_gaps(temp_configs)

            personnel = config.CNTS25165_PERSONNEL
            gongshi = load_gongshi_qingdan(personnel=personnel)
            waiyuan = load_waiyuan_qingkuang(personnel=personnel)
            project_dir = config.PROJECTS.get("CNTS-25165", {}).get("budget_report_dir", "")
            budget_actuals, _ = load_budget_report_actuals_and_forecast(project_dir, personnel)
            ts_days = load_timesheet_days(project_dir, personnel)

            merged = merge_ground_truth_allocation(
                gongshi, waiyuan, budget_actuals, ts_days, personnel, temp_configs
            )

            inventory = build_po_level_inventory(temp_configs, merged, personnel)
            consumption = calculate_level_consumption_rates(merged, personnel, inventory)

            now_month = datetime.now().strftime("%Y-%m")
            projections = run_forward_projection(
                inventory, consumption, temp_configs, personnel,
                start_month=now_month
            )
            transitions = detect_transitions(projections, inventory)

            previews = []
            for level, proj_list in projections.items():
                for fp in proj_list:
                    if fp.active_po_id == po_id or fp.transition_to_po == po_id:
                        previews.append({
                            "level": level,
                            "month": fp.month,
                            "active_po": fp.active_po_id,
                            "po_remaining": fp.po_remaining_after,
                            "transition": fp.transition,
                            "transition_date": fp.transition_date,
                            "alert": fp.alert,
                        })

            relevant_transitions = [
                t for t in transitions if t["to_po"] == po_id or t["from_po"] == po_id
            ]

            # 从实际推演数据中提取该 PO 的起止月份
            proj_months = sorted(set(p["month"] for p in previews))
            actual_start = proj_months[0] if proj_months else start_month
            actual_end = proj_months[-1] if proj_months else end_month

            return jsonify({
                "ok": True,
                "po_id": po_id,
                "projections": previews,
                "transitions": relevant_transitions,
                "effective_start": actual_start,
                "effective_end": actual_end,
            })

        finally:
            sys.stdout = old_stdout

    except Exception as e:
        return jsonify({"error": f"Preview error: {str(e)}"}), 500


# ============================================================
# 全局错误处理 (确保始终返回 JSON)
# ============================================================

import werkzeug.exceptions

def _json_error(e, code):
    resp = jsonify({"error": e.description if hasattr(e, 'description') else str(e),
                    "http_status": code})
    resp.status_code = code
    return resp

@app.errorhandler(werkzeug.exceptions.HTTPException)
def handle_http_exception(e):
    """捕获所有 HTTP 异常，返回 JSON"""
    resp = jsonify({"error": e.description, "http_status": e.code})
    resp.status_code = e.code
    return resp

# 兜底: 捕获所有未匹配路由
@app.route('/<path:path>', methods=['GET','POST','PUT','DELETE','PATCH','OPTIONS','HEAD'])
def catch_all_unmatched(path):
    return jsonify({"error": f"Endpoint not found: /{path}", "http_status": 404}), 404


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
    parser.add_argument("--host", default="0.0.0.0", help="Bind address (default: 0.0.0.0)")
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

    from werkzeug.serving import run_simple
    run_simple(args.host, args.port, app, use_reloader=False, threaded=True, processes=1)


if __name__ == "__main__":
    main()
