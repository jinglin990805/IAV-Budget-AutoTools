# -*- coding: utf-8 -*-
"""Step 2 audit: PO Master coverage and Monthly Closing reconciliation.

This module records audit facts and review items only.  It never approves a
timesheet candidate and never writes to a source workbook.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
from collections import defaultdict
from datetime import datetime, timezone
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import openpyxl

import config
from fact_store import FactStore
from po_pdf_parser import parse_po_pdf, po_to_config_rows
from project_registry import build_profile


MONTH_SHEETS = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "June": 6,
    "July": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
}


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _norm(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def _canonical_lookup() -> dict[str, str]:
    result: dict[str, str] = {}
    personnel = getattr(config, "CNTS25165_PERSONNEL", {})
    for canonical, info in personnel.items():
        for name in [canonical, *info.get("aliases", [])]:
            result[_norm(name)] = canonical
    return result


def _canonical_name(raw_name: str, lookup: dict[str, str]) -> str:
    value = _norm(raw_name)
    if value in lookup:
        return lookup[value]
    return ""


def _person_key(value: Any) -> str:
    """Normalize display-name punctuation for cross-source reconciliation."""
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]", "", str(value or "").casefold())


def _iter_pdfs(root: Path) -> list[Path]:
    if not root.exists():
        return []
    return sorted(path for path in root.rglob("*") if path.is_file() and path.suffix.lower() == ".pdf")


def _open_review_exists(
    store: FactStore,
    project_id: str,
    queue_type: str,
    candidate_id: str | None = None,
    target_key: str | None = None,
) -> bool:
    row = store.connection.execute(
        """SELECT 1 FROM review_queue
           WHERE project_id=? AND queue_type=? AND status='OPEN'
             AND COALESCE(candidate_id, '')=COALESCE(?, '')
             AND COALESCE(target_key, '')=COALESCE(?, '') LIMIT 1""",
        (project_id, queue_type, candidate_id, target_key),
    ).fetchone()
    return row is not None


def _ensure_review(
    store: FactStore,
    project_id: str,
    run_id: str,
    queue_type: str,
    payload: dict[str, Any],
    *,
    candidate_id: str | None = None,
    target_key: str | None = None,
    severity: str = "WARNING",
) -> None:
    if not _open_review_exists(store, project_id, queue_type, candidate_id, target_key):
        store.add_review_item(
            project_id, run_id, queue_type, payload,
            candidate_id=candidate_id, target_key=target_key, severity=severity,
        )


def _open_exception_exists(store: FactStore, project_id: str, exception_type: str, target_key: str) -> bool:
    row = store.connection.execute(
        """SELECT 1 FROM exceptions
           WHERE project_id=? AND exception_type=? AND target_key=? AND status='OPEN' LIMIT 1""",
        (project_id, exception_type, target_key),
    ).fetchone()
    return row is not None


def _ensure_exception(
    store: FactStore,
    project_id: str,
    exception_type: str,
    message: str,
    *,
    run_id: str,
    target_key: str,
    details: dict[str, Any],
    severity: str = "WARNING",
) -> None:
    existing = store.connection.execute(
        """SELECT exception_id FROM exceptions
           WHERE project_id=? AND exception_type=? AND target_key=? AND status='OPEN'
           LIMIT 1""",
        (project_id, exception_type, target_key),
    ).fetchone()
    if existing:
        store.connection.execute(
            """UPDATE exceptions
               SET run_id=?, severity=?, message=?, details_json=?
               WHERE exception_id=? AND status='OPEN'""",
            (run_id, severity, message, json.dumps(details, ensure_ascii=False, sort_keys=True), existing["exception_id"]),
        )
        store.connection.commit()
        return
    store.add_exception(
        project_id, exception_type, message, run_id=run_id,
        severity=severity, target_key=target_key, details=details,
    )


def _po_suggestions(value: str, known_ids: set[str]) -> list[dict[str, Any]]:
    digits = re.sub(r"\D", "", value)
    stripped = digits.lstrip("0")
    suggestions = []
    for po_id in known_ids:
        target = po_id.lstrip("0")
        ratio = SequenceMatcher(None, stripped, target).ratio()
        if stripped == target:
            ratio = 1.0
        elif stripped and (stripped in target or target in stripped):
            ratio = max(ratio, 0.86)
        if ratio >= 0.78:
            suggestions.append({"po_id": po_id, "similarity": round(ratio, 3)})
    return sorted(suggestions, key=lambda item: (-item["similarity"], item["po_id"]))[:3]


def audit_timesheet_validation(
    store: FactStore,
    project_id: str,
    run_id: str,
) -> dict[str, Any]:
    """Reconcile persisted Timesheet validation with the review gate.

    Older candidates may predate field-level validation review items. Rebuild
    those items from ``parse_json`` so a stale candidate cannot become eligible
    for bulk approval merely because its old queue is empty.
    """
    rows = store.connection.execute(
        """SELECT candidate_id, raw_name, year, month, po_id, total_hours,
                  parse_json, review_status
           FROM timesheet_candidates
           WHERE project_id=? AND review_status<>'SUPERSEDED'""",
        (project_id,),
    ).fetchall()
    blocked_items: list[dict[str, Any]] = []
    warning_items: list[dict[str, Any]] = []
    for row in rows:
        raw_payload = row["parse_json"] or "{}"
        try:
            payload = json.loads(raw_payload) if isinstance(raw_payload, str) else dict(raw_payload)
        except (TypeError, ValueError):
            payload = {}
            validation = {
                "status": "BLOCKED",
                "issues": [{
                    "code": "INVALID_PARSE_JSON",
                    "message": "候选 parse_json 无法解析",
                    "severity": "BLOCKING",
                }],
            }
        else:
            validation = payload.get("validation") or {}

        status = str(validation.get("status", "")).upper()
        issues = [item for item in (validation.get("issues") or []) if isinstance(item, dict)]
        if not validation:
            status = "UNKNOWN"
            issues = [{
                "code": "VALIDATION_MISSING",
                "message": "候选没有字段校验记录，需要重新审核",
                "severity": "WARNING",
            }]
        blocking = [
            item for item in issues
            if str(item.get("severity", "")).upper() == "BLOCKING"
        ]
        item = {
            "candidate_id": row["candidate_id"],
            "raw_name": row["raw_name"],
            "period": f"{row['year']:04d}-{row['month']:02d}" if row["year"] and row["month"] else "",
            "po_id": row["po_id"],
            "hours": row["total_hours"],
            "status": status,
            "issues": issues[:20],
        }
        if status == "BLOCKED" or blocking:
            blocked_items.append(item)
            _ensure_review(
                store, project_id, run_id, "VALIDATION_BLOCKED", item,
                candidate_id=row["candidate_id"],
                target_key=f"validation:{row['candidate_id']}",
                severity="HIGH",
            )
        elif status in {"WARNING", "UNKNOWN"} or issues:
            warning_items.append(item)
            _ensure_review(
                store, project_id, run_id, "VALIDATION_REVIEW", item,
                candidate_id=row["candidate_id"],
                target_key=f"validation:{row['candidate_id']}",
                severity="WARNING",
            )

    if blocked_items:
        _ensure_exception(
            store, project_id, "TIMESHEET_VALIDATION", "工时字段校验存在阻断问题，不能进入正式输出",
            run_id=run_id, target_key=f"timesheet-validation:{project_id}",
            details={"blocked_count": len(blocked_items), "items": blocked_items[:100]},
            severity="HIGH",
        )
    else:
        store.resolve_open_exceptions(
            project_id, "TIMESHEET_VALIDATION", set(),
            "最新 Timesheet validation 未发现阻断项",
        )

    return {
        "candidate_total": len(rows),
        "blocked_count": len(blocked_items),
        "warning_count": len(warning_items),
        "blocked_items": blocked_items[:100],
        "warning_items": warning_items[:100],
    }


def audit_po_master(store: FactStore, project_id: str, run_id: str, profile: Any) -> dict[str, Any]:
    root = Path(profile.source_dirs.get("po_pdf", ""))
    pdfs = _iter_pdfs(root)
    parsed_ids: set[str] = set()
    unparsed_files: list[str] = []

    for path in pdfs:
        source_id = store.record_source_file(project_id, "po_pdf", path, run_id=run_id)
        parse_target = f"po-source:{source_id}"
        info = None
        parse_message = ""
        try:
            info = parse_po_pdf(str(path))
        except Exception as exc:  # parser failures become review items
            parse_message = f"{type(exc).__name__}: {exc}"

        if info:
            parsed_ids.add(info.po_number)
            quotas = [
                {
                    "item": q.item, "qty_days": q.qty_days, "unit": q.unit,
                    "description": q.description, "unit_price": q.unit_price,
                    "total_rmb": q.total_rmb,
                }
                for q in info.level_quotas
            ]
            values = (
                f"pom-{hashlib.sha1((project_id + source_id).encode()).hexdigest()[:24]}",
                project_id, run_id, source_id, info.po_number, "PARSED",
                info.po_date, info.sap_po_number, info.contract_number,
                info.total_excl_vat, info.total_incl_vat, info.vat_rate,
                info.vat_amount, info.delivery_date, json.dumps(quotas, ensure_ascii=False),
                parse_message, _now(), _now(),
            )
        else:
            unparsed_files.append(str(path))
            values = (
                f"pom-{hashlib.sha1((project_id + source_id).encode()).hexdigest()[:24]}",
                project_id, run_id, source_id, "", "UNPARSED", "", "", "",
                None, None, None, None, "", "[]", parse_message or "parser returned no PO data",
                _now(), _now(),
            )

        store.connection.execute(
            """INSERT INTO po_master_records(
                   record_id, project_id, run_id, source_id, po_id, parse_status,
                   po_date, sap_po_number, contract_number, total_excl_vat,
                   total_incl_vat, vat_rate, vat_amount, delivery_date,
                   level_quotas_json, parse_message, created_at, updated_at)
               VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
               ON CONFLICT(project_id, source_id) DO UPDATE SET
                   run_id=excluded.run_id, po_id=excluded.po_id,
                   parse_status=excluded.parse_status, po_date=excluded.po_date,
                   sap_po_number=excluded.sap_po_number, contract_number=excluded.contract_number,
                   total_excl_vat=excluded.total_excl_vat, total_incl_vat=excluded.total_incl_vat,
                   vat_rate=excluded.vat_rate, vat_amount=excluded.vat_amount,
                   delivery_date=excluded.delivery_date, level_quotas_json=excluded.level_quotas_json,
                   parse_message=excluded.parse_message, updated_at=excluded.updated_at""",
            values,
        )

        if not info:
            payload = {"source_id": source_id, "path": str(path), "message": parse_message}
            _ensure_review(
                store, project_id, run_id, "PO_MASTER_PARSE", payload,
                target_key=parse_target, severity="HIGH",
            )
            _ensure_exception(
                store, project_id, "PO_PARSE_FAILED", "PO PDF 无法解析，需人工确认 PO 号和金额",
                run_id=run_id, target_key=parse_target, details=payload, severity="HIGH",
            )

    failed_parse_targets = {f"po-source:{row['source_id']}" for row in (
        store.connection.execute(
            """SELECT source_id FROM po_master_records
               WHERE project_id=? AND run_id=? AND parse_status='UNPARSED'""",
            (project_id, run_id),
        ).fetchall()
    )}
    store.resolve_open_exceptions(
        project_id, "PO_PARSE_FAILED", failed_parse_targets,
        "最新 PO 解析结果已成功覆盖该来源文件",
    )
    store.resolve_open_reviews(
        project_id, "PO_MASTER_PARSE", failed_parse_targets,
        "最新 PO 解析结果已成功覆盖该来源文件",
    )

    exact = potential = unknown = missing = 0
    unknown_items: list[dict[str, Any]] = []
    candidate_rows = store.connection.execute(
        """SELECT candidate_id, raw_name, canonical_name, year, month, po_id, total_hours
           FROM timesheet_candidates
           WHERE project_id=? AND review_status<>'SUPERSEDED'""",
        (project_id,),
    ).fetchall()
    for row in candidate_rows:
        po_id = str(row["po_id"] or "").strip()
        if not po_id:
            missing += 1
            queue_type = "MISSING_PO"
            payload = {"candidate_id": row["candidate_id"], "raw_name": row["raw_name"], "period": f"{row['year']}-{row['month']:02d}" if row["year"] and row["month"] else ""}
            _ensure_review(store, project_id, run_id, queue_type, payload, candidate_id=row["candidate_id"], severity="HIGH")
            continue
        if po_id in parsed_ids:
            exact += 1
            continue
        suggestions = _po_suggestions(po_id, parsed_ids)
        if suggestions:
            potential += 1
            severity = "WARNING"
        else:
            unknown += 1
            severity = "HIGH"
        item = {
            "candidate_id": row["candidate_id"], "raw_name": row["raw_name"],
            "po_id": po_id, "period": f"{row['year']}-{row['month']:02d}" if row["year"] and row["month"] else "",
            "hours": row["total_hours"], "suggestions": suggestions,
        }
        unknown_items.append(item)
        _ensure_review(
            store, project_id, run_id, "UNKNOWN_PO", item,
            candidate_id=row["candidate_id"], severity=severity,
        )

    if unknown_items:
        _ensure_exception(
            store, project_id, "UNKNOWN_PO", "工时候选包含未在 PO Master 中确认的 PO 号",
            run_id=run_id, target_key=f"po-candidates:{project_id}",
            details={"count": len(unknown_items), "items": unknown_items[:100]}, severity="HIGH",
        )

    if missing:
        _ensure_exception(
            store, project_id, "MISSING_PO", "工时候选缺少 PO 号，不能自动进入 PO Usage Status",
            run_id=run_id, target_key=f"missing-po:{project_id}",
            details={"count": missing}, severity="HIGH",
        )
    else:
        store.resolve_open_exceptions(
            project_id, "MISSING_PO", set(),
            "最新工时候选均已包含 PO 号",
        )
    if not unknown_items:
        store.resolve_open_exceptions(
            project_id, "UNKNOWN_PO", set(),
            "最新 PO Master 已覆盖所有工时候选 PO 号",
        )

    return {
        "pdf_files": len(pdfs), "parsed_files": len(pdfs) - len(unparsed_files),
        "unparsed_files": unparsed_files, "parsed_po_ids": sorted(parsed_ids),
        "candidate_exact": exact, "candidate_potential": potential,
        "candidate_unknown": unknown, "candidate_missing": missing,
        "candidate_total": len(candidate_rows),
    }


def _year_hint(path: Path) -> int:
    match = re.search(r"(20\d{2})", path.name)
    return int(match.group(1)) if match else datetime.now().year


def _current_hours_column(ws: Any, sheet_name: str) -> int:
    for row in ws.iter_rows(min_row=1, max_row=min(5, ws.max_row), values_only=True):
        for index, value in enumerate(row):
            text = _norm(value)
            if "this month" in text or text == _norm(sheet_name):
                return index
    return 9 if sheet_name == "Jan" else 11


def _row_period(row: tuple[Any, ...], sheet_name: str, year_hint: int) -> str:
    value = str(row[2] or "") if len(row) > 2 else ""
    match = re.search(r"(20\d{2})[_-](\d{1,2})", value)
    if match:
        return f"{int(match.group(1)):04d}-{int(match.group(2)):02d}"
    return f"{year_hint:04d}-{MONTH_SHEETS[sheet_name]:02d}"


def _project_number(row: tuple[Any, ...]) -> str:
    for value in row:
        match = re.search(r"(CNTS-25165(?:-\d+)?)", str(value or ""), re.I)
        if match:
            return match.group(1).upper()
    return ""


def audit_monthly_closing(store: FactStore, project_id: str, run_id: str, profile: Any) -> dict[str, Any]:
    root = Path(profile.source_dirs.get("monthly_closing", ""))
    files = sorted(path for path in root.glob("Monthly Closing_Data_Base_*.xlsx") if path.is_file())
    if not files:
        _ensure_exception(
            store, project_id, "MONTHLY_CLOSING_MISSING", "未找到 Monthly Closing 数据文件",
            run_id=run_id, target_key=f"monthly-closing-missing:{project_id}", details={"directory": str(root)}, severity="HIGH",
        )
        return {
            "files": [], "mapped_hours": {}, "unresolved_rows": 0,
            "approved_fact_count": 0, "mismatch_count": 0, "mismatches": [],
            "reconciliation_status": "SOURCE_MISSING",
        }

    path = files[-1]
    source_id = store.record_source_file(project_id, "monthly_closing", path, run_id=run_id)
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    lookup = _canonical_lookup()
    mapped_hours: defaultdict[str, float] = defaultdict(float)
    mapped_rows = 0
    unresolved_rows = 0
    unresolved_by_sheet: defaultdict[str, int] = defaultdict(int)
    unresolved_by_period: defaultdict[str, int] = defaultdict(int)
    empty_rows = 0

    for sheet_name, month_number in MONTH_SHEETS.items():
        if sheet_name not in workbook.sheetnames:
            continue
        ws = workbook[sheet_name]
        current_col = _current_hours_column(ws, sheet_name)
        for row_number, row in enumerate(ws.iter_rows(min_row=3, values_only=True), 3):
            row = tuple(row)
            if len(row) < 12:
                continue
            period = _row_period(row, sheet_name, _year_hint(path))
            project_number = _project_number(row)
            raw_name = f"{row[1] or ''} {row[0] or ''}".strip()
            canonical = _canonical_name(raw_name, lookup)
            hours_value = row[current_col] if current_col < len(row) else None
            hours = float(hours_value) if isinstance(hours_value, (int, float)) else None
            unresolved_project = any(str(value or "").strip().upper() in {"#N/A", "#REF!"} for value in row[10:14])
            empty_row = (
                not raw_name and not any(str(value or "").strip() for value in row[:10])
                and (hours is None or abs(hours) <= 0.01)
            )
            if empty_row:
                empty_rows += 1
                continue
            if not project_number and not unresolved_project:
                continue
            mapping_status = "MAPPED" if project_number else "UNRESOLVED_PROJECT"
            if mapping_status == "MAPPED":
                mapped_rows += 1
                if hours is not None and hours > 0:
                    mapped_hours[period] += hours
            else:
                unresolved_rows += 1
                unresolved_by_sheet[sheet_name] += 1
                unresolved_by_period[period] += 1
            fact_id = "mcf-" + hashlib.sha1(f"{project_id}|{source_id}|{sheet_name}|{row_number}".encode()).hexdigest()[:24]
            store.connection.execute(
                """INSERT INTO monthly_closing_facts(
                       fact_id, project_id, run_id, source_id, period, sheet_name,
                       row_number, raw_name, canonical_name, project_number, hours, mapping_status, created_at)
                   VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                   ON CONFLICT(project_id, source_id, sheet_name, row_number) DO UPDATE SET
                       run_id=excluded.run_id, period=excluded.period, raw_name=excluded.raw_name,
                       canonical_name=excluded.canonical_name, project_number=excluded.project_number,
                       hours=excluded.hours, mapping_status=excluded.mapping_status, created_at=excluded.created_at""",
                (fact_id, project_id, run_id, source_id, period, sheet_name, row_number,
                 raw_name, canonical, project_number, hours, mapping_status, _now()),
            )
    workbook.close()

    active_period = max(mapped_hours, default="")
    active_unresolved_rows = unresolved_by_period.get(active_period, 0)
    historical_unresolved_rows = max(unresolved_rows - active_unresolved_rows, 0)
    if unresolved_rows:
        target = f"monthly-closing-mapping:{source_id}"
        payload = {
            "source_id": source_id, "path": str(path),
            "unresolved_rows": unresolved_rows,
            "active_period": active_period,
            "active_unresolved_rows": active_unresolved_rows,
            "historical_unresolved_rows": historical_unresolved_rows,
            "empty_rows_ignored": empty_rows,
            "by_sheet": dict(unresolved_by_sheet),
            "by_period": dict(unresolved_by_period),
        }
        queue_type = (
            "MONTHLY_CLOSING_MAPPING" if active_unresolved_rows
            else "MONTHLY_CLOSING_HISTORICAL_MAPPING"
        )
        severity = "HIGH" if active_unresolved_rows else "WARNING"
        _ensure_review(store, project_id, run_id, queue_type, payload, target_key=target, severity=severity)
        _ensure_exception(
            store, project_id, "MONTHLY_CLOSING_MAPPING", "Monthly Closing 项目映射存在 #N/A/#REF!，无法完成全量对账",
            run_id=run_id, target_key=target, details=payload, severity=severity,
        )
        current_mapping_targets = {target}
    else:
        current_mapping_targets = set()
    current_mapping_queue = (
        "MONTHLY_CLOSING_MAPPING" if active_unresolved_rows
        else "MONTHLY_CLOSING_HISTORICAL_MAPPING" if unresolved_rows
        else ""
    )
    for queue_type in ("MONTHLY_CLOSING_MAPPING", "MONTHLY_CLOSING_HISTORICAL_MAPPING"):
        keep_targets = current_mapping_targets if queue_type == current_mapping_queue else set()
        store.resolve_open_reviews(
            project_id, queue_type, keep_targets,
            "最新 Monthly Closing 来源已重新审计，旧来源问题不再属于当前输入",
        )
    store.resolve_open_exceptions(
        project_id, "MONTHLY_CLOSING_MAPPING", current_mapping_targets,
        "最新 Monthly Closing 来源已重新审计，旧来源问题不再属于当前输入",
    )

    approved = store.connection.execute(
        """SELECT COUNT(*) AS count FROM timesheet_facts
           WHERE project_id=? AND approval_status='APPROVED' AND is_current=1""",
        (project_id,),
    ).fetchone()["count"]
    reconciliation_status = "WAITING_TIMESHEET_APPROVAL"
    mismatch_count = 0
    risk_count = 0
    covered_count = 0
    ignored_external_count = 0
    unreviewed_period_count = 0
    deferred_count = 0
    mismatches: list[dict[str, Any]] = []
    risk_items: list[dict[str, Any]] = []
    covered_items: list[dict[str, Any]] = []
    ignored_external_items: list[dict[str, Any]] = []
    deferred_items: list[dict[str, Any]] = []
    if approved:
        mc_rows = store.connection.execute(
            """SELECT canonical_name, period, SUM(hours) AS hours
               FROM monthly_closing_facts
               WHERE project_id=? AND source_id=? AND mapping_status='MAPPED' AND canonical_name<>''
               GROUP BY canonical_name, period""",
            (project_id, source_id),
        ).fetchall()
        fact_rows = store.connection.execute(
            """SELECT person_key, month, SUM(hours) AS hours
               FROM timesheet_facts
               WHERE project_id=? AND approval_status='APPROVED' AND is_current=1
               GROUP BY person_key, month""",
            (project_id,),
        ).fetchall()
        mc_values = {
            (_person_key(row["canonical_name"]), row["period"]): float(row["hours"] or 0)
            for row in mc_rows
        }
        fact_values = {
            (_person_key(row["person_key"]), row["month"]): float(row["hours"] or 0)
            for row in fact_rows
        }
        person_labels = {
            _person_key(row["canonical_name"]): row["canonical_name"]
            for row in mc_rows
            if row["canonical_name"]
        }
        person_labels.update(
            {
                _person_key(row["person_key"]): row["person_key"]
                for row in fact_rows
                if row["person_key"] and _person_key(row["person_key"]) not in person_labels
            }
        )
        fact_periods = {row["month"] for row in fact_rows}
        raw_deferred_rules = getattr(config, "STEP2_DEFERRED_RECONCILIATION", {}).get(project_id, {})
        deferred_rules = {
            (_person_key(person), period): reason
            for (person, period), reason in raw_deferred_rules.items()
        }
        for key in set(mc_values) | set(fact_values):
            if key[1] not in fact_periods:
                unreviewed_period_count += 1
                continue
            if key not in mc_values:
                ignored_external_count += 1
                ignored_external_items.append({
                    "person": person_labels.get(key[0], key[0]), "period": key[1],
                    "timesheet_hours": fact_values.get(key, 0),
                    "reason": "NOT_IN_MONTHLY_CLOSING_EXTERNAL_EMPLOYEE",
                })
                continue
            difference = round(mc_values.get(key, 0) - fact_values.get(key, 0), 2)
            if abs(difference) > 0.01:
                mismatch_count += 1
                mismatches.append({
                    "person": person_labels.get(key[0], key[0]), "period": key[1],
                    "pdf_hours": fact_values.get(key, 0),
                    "monthly_closing_hours": mc_values.get(key, 0),
                    "difference": difference,
                })
                if difference <= 0.01:
                    continue
                if key in deferred_rules:
                    continue
                _ensure_exception(
                    store, project_id, "MONTHLY_CLOSING_RECONCILIATION", "PDF 工时与 Monthly Closing 存在差异",
                    run_id=run_id, target_key=f"mc-reconcile:{key[0]}:{key[1]}",
                    details={"person": person_labels.get(key[0], key[0]), "period": key[1], "pdf_hours": fact_values.get(key, 0), "monthly_closing_hours": mc_values.get(key, 0), "difference": difference},
                    severity="WARNING",
                )
        for item in mismatches:
            item["risk"] = item["difference"] > 0.01
            deferred_key = (_person_key(item["person"]), item["period"])
            item["deferred"] = deferred_key in deferred_rules
            item["relationship"] = (
                "DEFERRED_REVIEW" if item["deferred"] else (
                    "MONTHLY_CLOSING_EXCEEDS_TIMESHEET"
                    if item["risk"] else "TIMESHEET_COVERS_MONTHLY_CLOSING"
                )
            )
            if item["deferred"]:
                deferred_count += 1
                item["deferred_reason"] = deferred_rules[deferred_key]
                deferred_items.append(item)
                _ensure_review(
                    store, project_id, run_id, "MONTHLY_CLOSING_DEFERRED",
                    item, target_key=f"mc-deferred:{item['person']}:{item['period']}",
                    severity="WARNING",
                )
                _ensure_exception(
                    store, project_id, "MONTHLY_CLOSING_DEFERRED",
                    "Monthly Closing variance is deferred by an explicit business exception",
                    run_id=run_id, target_key=f"mc-deferred:{item['person']}:{item['period']}",
                    details=item, severity="WARNING",
                )
                continue
            if item["risk"]:
                risk_count += 1
                risk_items.append(item)
            else:
                covered_count += 1
                covered_items.append(item)
        active_reconcile_targets = {
            f"mc-reconcile:{item['person']}:{item['period']}"
            for item in risk_items
        }
        store.resolve_open_exceptions(
            project_id, "MONTHLY_CLOSING_RECONCILIATION", active_reconcile_targets,
            "最新对账中该记录已不再构成 Monthly Closing 大于 Timesheet 的风险",
        )
        store.resolve_open_exceptions(
            project_id, "MONTHLY_CLOSING_RISK", set(),
            "最新对账规则已重新计算，该旧风险不再属于当前风险",
        )
        active_deferred_targets = {
            f"mc-deferred:{item['person']}:{item['period']}"
            for item in deferred_items
        }
        store.resolve_open_exceptions(
            project_id, "MONTHLY_CLOSING_DEFERRED", active_deferred_targets,
            "最新对账中该延期复核项已不再存在",
        )
        if risk_count:
            reconciliation_status = "RISK_REVIEW"
        elif deferred_count:
            reconciliation_status = "DEFERRED_REVIEW"
        else:
            reconciliation_status = "MATCH"

    return {
        "files": [str(path)], "mapped_hours": dict(sorted(mapped_hours.items())),
        "mapped_rows": mapped_rows, "unresolved_rows": unresolved_rows,
        "unresolved_by_sheet": dict(unresolved_by_sheet),
        "unresolved_by_period": dict(unresolved_by_period),
        "active_period": active_period,
        "active_unresolved_rows": active_unresolved_rows,
        "historical_unresolved_rows": historical_unresolved_rows,
        "empty_rows_ignored": empty_rows, "approved_fact_count": approved,
        "mismatch_count": mismatch_count,
        "risk_count": risk_count,
        "covered_count": covered_count,
        "deferred_count": deferred_count,
        "ignored_external_count": ignored_external_count,
        "unreviewed_period_count": unreviewed_period_count,
        "mismatches": sorted(mismatches, key=lambda item: (item["period"], item["person"])),
        "risk_items": sorted(risk_items, key=lambda item: (item["period"], item["person"])),
        "covered_items": sorted(covered_items, key=lambda item: (item["period"], item["person"])),
        "deferred_items": sorted(deferred_items, key=lambda item: (item["period"], item["person"])),
        "ignored_external_items": sorted(
            ignored_external_items, key=lambda item: (item["period"], item["person"])
        ),
        "reconciliation_status": reconciliation_status,
    }


def run_audit(project_id: str) -> dict[str, Any]:
    profile = build_profile(project_id)
    with FactStore() as store:
        store.initialize()
        store.register_project(profile)
        run_id = store.create_run(project_id, notes="Step 2 PO Master and Monthly Closing audit")
        validation_summary = audit_timesheet_validation(store, project_id, run_id)
        po_summary = audit_po_master(store, project_id, run_id, profile)
        monthly_summary = audit_monthly_closing(store, project_id, run_id, profile)
        needs_review = bool(
            validation_summary["blocked_count"] or validation_summary["warning_count"] or
            po_summary["unparsed_files"] or po_summary["candidate_potential"] or
            po_summary["candidate_unknown"] or po_summary["candidate_missing"] or
            monthly_summary.get("active_unresolved_rows", monthly_summary.get("unresolved_rows", 0))
            or monthly_summary.get("risk_count")
        )
        status = "REVIEW_REQUIRED" if needs_review else "COMPLETED"
        store.connection.execute(
            "UPDATE runs SET status=?, completed_at=? WHERE run_id=?",
            (status, _now(), run_id),
        )
        return {
            "ok": True, "project": project_id, "run_id": run_id, "status": status,
            "timesheet_validation": validation_summary,
            "po_master": po_summary, "monthly_closing": monthly_summary,
        }


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit PO Master and Monthly Closing for Step 2")
    parser.add_argument("--project", default="CNTS-25165")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    result = run_audit(args.project)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=None if args.json else 2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
