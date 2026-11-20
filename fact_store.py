# -*- coding: utf-8 -*-
"""本地 SQLite 事实存储的第一版基础设施。

第一阶段只初始化 schema、登记项目和来源文件指纹；业务事实的导入由后续
输入适配器完成。所有记录以 project_id + run_id 隔离，Excel 不被此模块修改。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sqlite3
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

import config


SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS schema_meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS projects (
    project_id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    profile_json TEXT NOT NULL,
    enabled INTEGER NOT NULL DEFAULT 0 CHECK(enabled IN (0, 1)),
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS runs (
    run_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL REFERENCES projects(project_id),
    status TEXT NOT NULL DEFAULT 'DRAFT',
    data_as_of TEXT,
    source_cutoff TEXT,
    started_at TEXT NOT NULL,
    completed_at TEXT,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS source_files (
    source_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL REFERENCES projects(project_id),
    run_id TEXT REFERENCES runs(run_id),
    source_type TEXT NOT NULL,
    path TEXT NOT NULL,
    file_name TEXT NOT NULL,
    size_bytes INTEGER,
    mtime_ns INTEGER,
    sha256 TEXT NOT NULL,
    discovered_at TEXT NOT NULL,
    metadata_json TEXT NOT NULL DEFAULT '{}',
    UNIQUE(project_id, path, sha256)
);

CREATE TABLE IF NOT EXISTS person_master (
    person_key TEXT PRIMARY KEY,
    canonical_name TEXT NOT NULL,
    aliases_json TEXT NOT NULL DEFAULT '[]',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS project_person_rates (
    project_id TEXT NOT NULL REFERENCES projects(project_id),
    person_key TEXT NOT NULL REFERENCES person_master(person_key),
    level TEXT,
    customer_day_rate REAL,
    internal_cost_day_rate REAL,
    hours_per_day REAL NOT NULL,
    effective_from TEXT,
    effective_to TEXT,
    source_id TEXT REFERENCES source_files(source_id),
    status TEXT NOT NULL DEFAULT 'CONFIRMED',
    PRIMARY KEY(project_id, person_key, effective_from)
);

CREATE TABLE IF NOT EXISTS timesheet_facts (
    fact_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL REFERENCES projects(project_id),
    run_id TEXT NOT NULL REFERENCES runs(run_id),
    person_key TEXT,
    source_name TEXT,
    month TEXT NOT NULL,
    po_id TEXT,
    hours REAL NOT NULL,
    raw_value REAL,
    raw_unit TEXT,
    source_id TEXT REFERENCES source_files(source_id),
    confidence REAL,
    approval_status TEXT NOT NULL DEFAULT 'PENDING',
    version INTEGER NOT NULL DEFAULT 1,
    is_current INTEGER NOT NULL DEFAULT 1 CHECK(is_current IN (0, 1)),
    notes TEXT
);

CREATE TABLE IF NOT EXISTS cost_facts (
    fact_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL REFERENCES projects(project_id),
    run_id TEXT NOT NULL REFERENCES runs(run_id),
    month TEXT,
    source_type TEXT NOT NULL,
    po_id TEXT,
    cost_layer TEXT NOT NULL,
    currency TEXT NOT NULL,
    original_amount REAL,
    amount REAL NOT NULL,
    booking_status TEXT NOT NULL DEFAULT 'PENDING',
    source_id TEXT REFERENCES source_files(source_id),
    approval_status TEXT NOT NULL DEFAULT 'PENDING',
    notes TEXT,
    gr_ref TEXT,
    invoice_ref TEXT,
    invoice_date TEXT,
    invoice_source_id TEXT REFERENCES source_files(source_id),
    payment_status TEXT NOT NULL DEFAULT 'PENDING',
    payment_ref TEXT,
    payment_date TEXT
);

CREATE TABLE IF NOT EXISTS cost_fact_po_allocations (
    allocation_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL REFERENCES projects(project_id),
    run_id TEXT NOT NULL REFERENCES runs(run_id),
    cost_fact_id TEXT NOT NULL REFERENCES cost_facts(fact_id),
    month TEXT,
    po_id TEXT NOT NULL,
    wbs TEXT NOT NULL,
    hours REAL NOT NULL,
    amount REAL NOT NULL,
    allocation_basis TEXT NOT NULL,
    approval_status TEXT NOT NULL DEFAULT 'PENDING',
    booking_status TEXT NOT NULL DEFAULT 'PENDING',
    gr_ref TEXT,
    invoice_ref TEXT,
    invoice_date TEXT,
    invoice_source_id TEXT REFERENCES source_files(source_id),
    payment_status TEXT NOT NULL DEFAULT 'PENDING',
    payment_ref TEXT,
    payment_date TEXT,
    notes TEXT,
    UNIQUE(cost_fact_id, po_id)
);

CREATE INDEX IF NOT EXISTS idx_cost_alloc_project_po
    ON cost_fact_po_allocations(project_id, po_id, month);

CREATE TABLE IF NOT EXISTS adjustments (
    adjustment_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL REFERENCES projects(project_id),
    run_id TEXT NOT NULL REFERENCES runs(run_id),
    target_type TEXT NOT NULL,
    target_key TEXT NOT NULL,
    field_name TEXT NOT NULL,
    old_value TEXT,
    new_value TEXT,
    reason_category TEXT NOT NULL,
    note TEXT,
    created_at TEXT NOT NULL,
    created_by TEXT,
    approval_status TEXT NOT NULL DEFAULT 'PENDING'
);

CREATE TABLE IF NOT EXISTS exceptions (
    exception_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL REFERENCES projects(project_id),
    run_id TEXT REFERENCES runs(run_id),
    exception_type TEXT NOT NULL,
    severity TEXT NOT NULL DEFAULT 'WARNING',
    status TEXT NOT NULL DEFAULT 'OPEN',
    target_key TEXT,
    message TEXT NOT NULL,
    details_json TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL,
    resolved_at TEXT,
    resolution_note TEXT
);

CREATE TABLE IF NOT EXISTS approvals (
    approval_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL REFERENCES projects(project_id),
    run_id TEXT NOT NULL REFERENCES runs(run_id),
    approval_type TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'PENDING',
    approved_by TEXT,
    approved_at TEXT,
    note TEXT
);

CREATE TABLE IF NOT EXISTS timesheet_candidates (
    candidate_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL REFERENCES projects(project_id),
    run_id TEXT NOT NULL REFERENCES runs(run_id),
    source_id TEXT REFERENCES source_files(source_id),
    raw_name TEXT,
    canonical_name TEXT,
    year INTEGER,
    month INTEGER,
    po_id TEXT,
    total_hours REAL,
    raw_unit TEXT,
    parse_method TEXT NOT NULL,
    version_key TEXT NOT NULL,
    confidence REAL,
    parse_json TEXT NOT NULL DEFAULT '{}',
    review_status TEXT NOT NULL DEFAULT 'PENDING',
    is_preferred INTEGER NOT NULL DEFAULT 0 CHECK(is_preferred IN (0, 1)),
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS review_queue (
    review_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL REFERENCES projects(project_id),
    run_id TEXT NOT NULL REFERENCES runs(run_id),
    queue_type TEXT NOT NULL,
    severity TEXT NOT NULL DEFAULT 'WARNING',
    status TEXT NOT NULL DEFAULT 'OPEN',
    candidate_id TEXT REFERENCES timesheet_candidates(candidate_id),
    target_key TEXT,
    payload_json TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL,
    resolved_at TEXT,
    resolution_note TEXT
);

CREATE TABLE IF NOT EXISTS po_master_records (
    record_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL REFERENCES projects(project_id),
    run_id TEXT NOT NULL REFERENCES runs(run_id),
    source_id TEXT REFERENCES source_files(source_id),
    po_id TEXT,
    parse_status TEXT NOT NULL,
    po_date TEXT,
    sap_po_number TEXT,
    contract_number TEXT,
    total_excl_vat REAL,
    total_incl_vat REAL,
    vat_rate REAL,
    vat_amount REAL,
    delivery_date TEXT,
    level_quotas_json TEXT NOT NULL DEFAULT '[]',
    parse_message TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(project_id, source_id)
);

CREATE TABLE IF NOT EXISTS monthly_closing_facts (
    fact_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL REFERENCES projects(project_id),
    run_id TEXT NOT NULL REFERENCES runs(run_id),
    source_id TEXT REFERENCES source_files(source_id),
    period TEXT,
    sheet_name TEXT NOT NULL,
    row_number INTEGER NOT NULL,
    raw_name TEXT,
    canonical_name TEXT,
    project_number TEXT,
    hours REAL,
    mapping_status TEXT NOT NULL,
    created_at TEXT NOT NULL,
    UNIQUE(project_id, source_id, sheet_name, row_number)
);

CREATE INDEX IF NOT EXISTS idx_runs_project ON runs(project_id, started_at);
CREATE INDEX IF NOT EXISTS idx_sources_project_path ON source_files(project_id, path);
CREATE INDEX IF NOT EXISTS idx_timesheet_project_month ON timesheet_facts(project_id, month);
CREATE INDEX IF NOT EXISTS idx_cost_project_month ON cost_facts(project_id, month);
CREATE INDEX IF NOT EXISTS idx_exception_open ON exceptions(project_id, status);
CREATE INDEX IF NOT EXISTS idx_candidate_version ON timesheet_candidates(project_id, version_key);
CREATE INDEX IF NOT EXISTS idx_review_open ON review_queue(project_id, status);
CREATE INDEX IF NOT EXISTS idx_po_master_project ON po_master_records(project_id, po_id);
CREATE INDEX IF NOT EXISTS idx_monthly_closing_project_period ON monthly_closing_facts(project_id, period);
"""


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _text(value: Any) -> str | None:
    return None if value is None else str(value)


def load_approved_timesheet_facts(
    project_id: str,
    db_path: str | os.PathLike[str] | None = None,
) -> list[dict[str, Any]]:
    """Read current, approved Timesheet facts without mutating the fact store."""
    path = Path(db_path or config.FACT_STORE_PATH)
    if not path.is_file():
        return []
    uri = f"{path.resolve().as_uri()}?mode=ro"
    connection = sqlite3.connect(uri, uri=True)
    connection.row_factory = sqlite3.Row
    try:
        rows = connection.execute(
            """SELECT f.fact_id, f.person_key,
                      COALESCE(p.canonical_name, f.person_key) AS canonical_name,
                      f.month, f.po_id, f.hours, f.raw_unit, f.source_id,
                      f.version
               FROM timesheet_facts AS f
               LEFT JOIN person_master AS p ON p.person_key=f.person_key
               WHERE f.project_id=? AND f.source_name='timesheet_pdf'
                 AND f.approval_status='APPROVED' AND f.is_current=1
               ORDER BY f.month, f.person_key, f.po_id, f.version""",
            (project_id,),
        ).fetchall()
        return [dict(row) for row in rows]
    finally:
        connection.close()


class FactStore:
    def __init__(self, db_path: str | os.PathLike[str] | None = None):
        self.db_path = Path(db_path or config.FACT_STORE_PATH)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.connection = sqlite3.connect(str(self.db_path))
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")

    def close(self) -> None:
        self.connection.close()

    def __enter__(self) -> "FactStore":
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        if exc_type:
            self.connection.rollback()
        else:
            self.connection.commit()
        self.close()

    def initialize(self) -> None:
        self.connection.executescript(SCHEMA_SQL)
        # Schema v2 was created before candidate notes were needed; keep existing
        # local databases compatible without deleting or rewriting any facts.
        candidate_columns = {
            row["name"] for row in self.connection.execute(
                "PRAGMA table_info(timesheet_candidates)"
            )
        }
        if "notes" not in candidate_columns:
            self.connection.execute("ALTER TABLE timesheet_candidates ADD COLUMN notes TEXT")
        cost_columns = {
            row["name"] for row in self.connection.execute(
                "PRAGMA table_info(cost_facts)"
            )
        }
        for column, definition in {
            "gr_ref": "TEXT",
            "invoice_ref": "TEXT",
            "invoice_date": "TEXT",
            "invoice_source_id": "TEXT",
            "payment_status": "TEXT NOT NULL DEFAULT 'PENDING'",
            "payment_ref": "TEXT",
            "payment_date": "TEXT",
        }.items():
            if column not in cost_columns:
                self.connection.execute(
                    f"ALTER TABLE cost_facts ADD COLUMN {column} {definition}"
                )
        self.connection.execute(
            "INSERT INTO schema_meta(key, value) VALUES('schema_version', ?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (str(config.FACT_SCHEMA_VERSION),),
        )
        self.connection.commit()

    def register_project(self, profile: Any, enabled: bool = False) -> None:
        payload = profile.to_dict() if hasattr(profile, "to_dict") else dict(profile)
        timestamp = _now()
        self.connection.execute(
            """INSERT INTO projects(project_id, name, profile_json, enabled, created_at, updated_at)
               VALUES(?, ?, ?, ?, ?, ?)
               ON CONFLICT(project_id) DO UPDATE SET
                   name=excluded.name, profile_json=excluded.profile_json,
                   updated_at=excluded.updated_at""",
            (
                payload["project_id"], payload["name"],
                json.dumps(payload, ensure_ascii=False, sort_keys=True),
                int(enabled), timestamp, timestamp,
            ),
        )
        self.connection.commit()

    def create_run(
        self,
        project_id: str,
        data_as_of: str | None = None,
        source_cutoff: str | None = None,
        notes: str = "",
    ) -> str:
        run_id = f"run-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:8]}"
        self.connection.execute(
            """INSERT INTO runs(run_id, project_id, status, data_as_of, source_cutoff, started_at, notes)
               VALUES(?, ?, 'DRAFT', ?, ?, ?, ?)""",
            (run_id, project_id, data_as_of, source_cutoff, _now(), notes),
        )
        self.connection.commit()
        return run_id

    @staticmethod
    def fingerprint_file(path: str | os.PathLike[str]) -> dict[str, Any]:
        file_path = Path(path)
        digest = hashlib.sha256()
        with file_path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        stat = file_path.stat()
        return {
            "path": str(file_path.resolve()),
            "file_name": file_path.name,
            "size_bytes": stat.st_size,
            "mtime_ns": stat.st_mtime_ns,
            "sha256": digest.hexdigest(),
        }

    def record_source_file(
        self,
        project_id: str,
        source_type: str,
        path: str | os.PathLike[str],
        run_id: str | None = None,
        metadata: dict[str, Any] | None = None,
    ) -> str:
        fingerprint = self.fingerprint_file(path)
        source_id = f"src-{fingerprint['sha256'][:20]}"
        self.connection.execute(
            """INSERT INTO source_files(
                   source_id, project_id, run_id, source_type, path, file_name,
                   size_bytes, mtime_ns, sha256, discovered_at, metadata_json)
               VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
               ON CONFLICT(project_id, path, sha256) DO UPDATE SET
                   run_id=excluded.run_id, source_type=excluded.source_type,
                   size_bytes=excluded.size_bytes, mtime_ns=excluded.mtime_ns,
                   metadata_json=excluded.metadata_json""",
            (
                source_id, project_id, run_id, source_type,
                fingerprint["path"], fingerprint["file_name"],
                fingerprint["size_bytes"], fingerprint["mtime_ns"],
                fingerprint["sha256"], _now(),
                json.dumps(metadata or {}, ensure_ascii=False, sort_keys=True),
            ),
        )
        self.connection.commit()
        return source_id

    def find_source_by_hash(self, project_id: str, sha256: str) -> sqlite3.Row | None:
        return self.connection.execute(
            "SELECT * FROM source_files WHERE project_id=? AND sha256=? ORDER BY discovered_at LIMIT 1",
            (project_id, sha256),
        ).fetchone()

    def find_candidates(self, project_id: str, version_key: str) -> list[sqlite3.Row]:
        return list(self.connection.execute(
            "SELECT * FROM timesheet_candidates WHERE project_id=? AND version_key=? ORDER BY created_at",
            (project_id, version_key),
        ))

    def insert_timesheet_candidate(self, candidate: dict[str, Any]) -> str:
        candidate_id = candidate.get("candidate_id") or f"tsc-{uuid.uuid4().hex[:20]}"
        self.connection.execute(
            """INSERT INTO timesheet_candidates(
                   candidate_id, project_id, run_id, source_id, raw_name, canonical_name,
                   year, month, po_id, total_hours, raw_unit, parse_method, version_key,
                   confidence, parse_json, review_status, is_preferred, created_at)
               VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                candidate_id, candidate["project_id"], candidate["run_id"],
                candidate.get("source_id"), candidate.get("raw_name"),
                candidate.get("canonical_name"), candidate.get("year"), candidate.get("month"),
                candidate.get("po_id"), candidate.get("total_hours"), candidate.get("raw_unit", "hour"),
                candidate.get("parse_method", "pdf_text"), candidate["version_key"],
                candidate.get("confidence"),
                json.dumps(candidate.get("parse_json", {}), ensure_ascii=False, sort_keys=True),
                candidate.get("review_status", "PENDING"), int(candidate.get("is_preferred", False)),
                candidate.get("created_at", _now()),
            ),
        )
        self.connection.commit()
        return candidate_id

    def add_review_item(
        self,
        project_id: str,
        run_id: str,
        queue_type: str,
        payload: dict[str, Any],
        candidate_id: str | None = None,
        target_key: str | None = None,
        severity: str = "WARNING",
    ) -> str:
        review_id = f"review-{uuid.uuid4().hex[:16]}"
        self.connection.execute(
            """INSERT INTO review_queue(
                   review_id, project_id, run_id, queue_type, severity, candidate_id,
                   target_key, payload_json, created_at)
               VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                review_id, project_id, run_id, queue_type, severity, candidate_id,
                target_key, json.dumps(payload, ensure_ascii=False, sort_keys=True), _now(),
            ),
        )
        self.connection.commit()
        return review_id

    def update_timesheet_candidate(self, candidate_id: str, candidate: dict[str, Any]) -> None:
        self.connection.execute(
            """UPDATE timesheet_candidates SET
                   raw_name=?, canonical_name=?, year=?, month=?, po_id=?, total_hours=?,
                   raw_unit=?, parse_method=?, version_key=?, confidence=?, parse_json=?, review_status=?
               WHERE candidate_id=?""",
            (
                candidate.get("raw_name"), candidate.get("canonical_name"), candidate.get("year"),
                candidate.get("month"), candidate.get("po_id"), candidate.get("total_hours"),
                candidate.get("raw_unit", "hour"), candidate.get("parse_method", "ocr"),
                candidate["version_key"],
                candidate.get("confidence"),
                json.dumps(candidate.get("parse_json", {}), ensure_ascii=False, sort_keys=True),
                candidate.get("review_status", "REVIEW"), candidate_id,
            ),
        )
        self.connection.commit()

    def resolve_candidate_reviews(self, candidate_id: str, note: str = "") -> None:
        self.connection.execute(
            """UPDATE review_queue SET status='RESOLVED', resolved_at=?, resolution_note=?
               WHERE candidate_id=? AND status='OPEN'""",
            (_now(), note, candidate_id),
        )
        self.connection.commit()

    def resolve_open_reviews(
        self,
        project_id: str,
        queue_type: str,
        keep_target_keys: set[str] | None = None,
        note: str = "",
    ) -> int:
        """Close non-candidate review items no longer emitted by the latest run."""
        keep = keep_target_keys or set()
        rows = self.connection.execute(
            """SELECT review_id, target_key FROM review_queue
               WHERE project_id=? AND queue_type=? AND status='OPEN'
                 AND candidate_id IS NULL""",
            (project_id, queue_type),
        ).fetchall()
        stale = [row["review_id"] for row in rows if row["target_key"] not in keep]
        if not stale:
            return 0
        timestamp = _now()
        self.connection.executemany(
            """UPDATE review_queue
               SET status='RESOLVED', resolved_at=?, resolution_note=?
               WHERE review_id=? AND status='OPEN'""",
            [(timestamp, note, review_id) for review_id in stale],
        )
        self.connection.commit()
        return len(stale)

    def dismiss_candidate(self, candidate_id: str, note: str = "") -> None:
        """Soft-delete a review candidate while preserving its audit trail."""
        self.connection.execute(
            "UPDATE timesheet_candidates SET review_status='DISMISSED', is_preferred=0, notes=? WHERE candidate_id=?",
            (note, candidate_id),
        )
        self.connection.execute(
            """UPDATE review_queue SET status='RESOLVED', resolved_at=?, resolution_note=?
               WHERE candidate_id=? AND status='OPEN'""",
            (_now(), note or "候选条目已从审核队列移除", candidate_id),
        )
        self.connection.commit()

    def supersede_candidate(self, candidate_id: str, replacement_id: str, note: str) -> None:
        self.connection.execute(
            "UPDATE timesheet_candidates SET review_status='SUPERSEDED', is_preferred=0, notes=? WHERE candidate_id=?",
            (f"{note}; replacement={replacement_id}", candidate_id),
        )
        self.connection.execute(
            """UPDATE review_queue SET status='RESOLVED', resolved_at=?, resolution_note=?
               WHERE candidate_id=? AND status='OPEN'""",
            (_now(), f"候选被完全重复记录替代：{replacement_id}", candidate_id),
        )
        self.connection.commit()

    def record_adjustment(
        self,
        project_id: str,
        run_id: str,
        target_type: str,
        target_key: str,
        field_name: str,
        old_value: Any,
        new_value: Any,
        reason_category: str,
        note: str = "",
        created_by: str = "system",
        approval_status: str = "APPROVED",
    ) -> str:
        """Persist one before/after field change.

        ``adjustments`` is the only durable record of a reviewer's correction:
        re-scanning a PDF may refresh ``timesheet_candidates.parse_json``, so a
        correction that is not written here is lost.
        """
        adjustment_id = f"adj-{uuid.uuid4().hex[:20]}"
        self.connection.execute(
            """INSERT INTO adjustments(
                   adjustment_id, project_id, run_id, target_type, target_key,
                   field_name, old_value, new_value, reason_category, note,
                   created_at, created_by, approval_status)
               VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                adjustment_id, project_id, run_id, target_type, target_key,
                field_name, _text(old_value), _text(new_value),
                reason_category, note, _now(), created_by, approval_status,
            ),
        )
        self.connection.commit()
        return adjustment_id

    def _ocr_proposed_values(self, candidate_id: str) -> dict[str, Any]:
        """Recover what OCR proposed for one candidate, before any review edit.

        ``timesheet_candidates.total_hours`` is overwritten with the approved
        value once a reviewer confirms, so ``parse_json.fields`` -- written at
        candidate creation and never rewritten on approval -- is the only
        pristine copy of the OCR output.
        """
        row = self.connection.execute(
            """SELECT total_hours, po_id, canonical_name, raw_name, parse_json
               FROM timesheet_candidates WHERE candidate_id=?""",
            (candidate_id,),
        ).fetchone()
        if row is None:
            return {}
        try:
            parsed = json.loads(row["parse_json"] or "{}")
        except (TypeError, ValueError):
            parsed = {}
        fields = parsed.get("fields") or {}
        return {
            "hours": fields.get("hours") or row["total_hours"],
            "po_id": fields.get("po_id") or row["po_id"],
            "person": row["canonical_name"] or row["raw_name"] or fields.get("name"),
            "period": fields.get("period"),
        }

    def _record_timesheet_corrections(
        self,
        candidate_id: str,
        project_id: str,
        run_id: str,
        person_key: str,
        month: str,
        po_id: str,
        hours: float,
    ) -> list[str]:
        """Record how the approved values differ from what OCR proposed.

        Each row is tagged with whether OCR actually produced a *different*
        value (``ocr_field_correction`` -- a genuine extraction failure worth
        turning into a regression test) or produced nothing at all
        (``review_fill`` -- a recovery, not a misread).
        """
        proposed = self._ocr_proposed_values(candidate_id)
        if not proposed:
            return []

        def same(left: Any, right: Any) -> bool:
            if left is None or right is None:
                return left == right
            try:
                return abs(float(left) - float(right)) < 1e-9
            except (TypeError, ValueError):
                return str(left).strip() == str(right).strip()

        pairs = [
            ("hours", proposed.get("hours"), hours),
            ("po_id", proposed.get("po_id"), po_id),
            ("person_key", proposed.get("person"), person_key),
            ("month", proposed.get("period"), month),
        ]
        recorded: list[str] = []
        for field_name, old_value, new_value in pairs:
            if same(old_value, new_value):
                continue
            empty = old_value is None or str(old_value).strip() in ("", "None")
            recorded.append(self.record_adjustment(
                project_id, run_id, "timesheet_candidate", candidate_id,
                field_name, old_value, new_value,
                "review_fill" if empty else "ocr_field_correction",
                note=(
                    "复核时补填：OCR 未产出该字段"
                    if empty
                    else "复核修正：OCR 原值与确认值不一致"
                ),
                created_by="timesheet_review",
            ))
        return recorded

    def save_approved_timesheet(
        self,
        candidate_id: str,
        project_id: str,
        run_id: str,
        person_key: str,
        month: str,
        po_id: str,
        hours: float,
        source_id: str | None,
        raw_unit: str = "hour",
        note: str = "",
    ) -> str:
        fact_id = f"tsf-{candidate_id}"
        proposed = self._ocr_proposed_values(candidate_id)
        raw_hours = proposed.get("hours")
        try:
            raw_value = float(raw_hours)
        except (TypeError, ValueError):
            raw_value = float(hours)
        self.connection.execute(
            """INSERT INTO timesheet_facts(
                   fact_id, project_id, run_id, person_key, source_name, month, po_id,
                   hours, raw_value, raw_unit, source_id, confidence, approval_status,
                   version, is_current, notes)
               VALUES(?, ?, ?, ?, 'timesheet_pdf', ?, ?, ?, ?, ?, ?, 1.0, 'APPROVED', 1, 1, ?)
               ON CONFLICT(fact_id) DO UPDATE SET
                   project_id=excluded.project_id, run_id=excluded.run_id,
                   person_key=excluded.person_key, month=excluded.month, po_id=excluded.po_id,
                   hours=excluded.hours, raw_value=excluded.raw_value,
                   raw_unit=excluded.raw_unit, source_id=excluded.source_id,
                   approval_status='APPROVED', is_current=1, notes=excluded.notes""",
            (
                fact_id, project_id, run_id, person_key, month, po_id,
                float(hours), raw_value, raw_unit, source_id, note,
            ),
        )
        self.connection.commit()
        self._record_timesheet_corrections(
            candidate_id, project_id, run_id, person_key, month, po_id, hours,
        )
        return fact_id

    def upsert_person_rate(
        self,
        project_id: str,
        person_key: str,
        canonical_name: str,
        aliases: list[str],
        level: str,
        customer_day_rate: float,
        effective_from: str,
        hours_per_day: float,
        internal_cost_day_rate: float | None = None,
    ) -> None:
        timestamp = _now()
        self.connection.execute(
            """INSERT INTO person_master(person_key, canonical_name, aliases_json, created_at, updated_at)
               VALUES(?, ?, ?, ?, ?)
               ON CONFLICT(person_key) DO UPDATE SET
                   canonical_name=excluded.canonical_name,
                   aliases_json=excluded.aliases_json, updated_at=excluded.updated_at""",
            (person_key, canonical_name, json.dumps(sorted(set(aliases)), ensure_ascii=False), timestamp, timestamp),
        )
        self.connection.execute(
            """INSERT INTO project_person_rates(
                   project_id, person_key, level, customer_day_rate, internal_cost_day_rate,
                   hours_per_day, effective_from, status)
               VALUES(?, ?, ?, ?, ?, ?, ?, 'CONFIRMED')
               ON CONFLICT(project_id, person_key, effective_from) DO UPDATE SET
                   level=excluded.level, customer_day_rate=excluded.customer_day_rate,
                   internal_cost_day_rate=excluded.internal_cost_day_rate,
                   hours_per_day=excluded.hours_per_day, status='CONFIRMED'""",
            (
                project_id, person_key, level, float(customer_day_rate),
                internal_cost_day_rate, float(hours_per_day), effective_from,
            ),
        )
        self.connection.commit()

    def add_exception(
        self,
        project_id: str,
        exception_type: str,
        message: str,
        run_id: str | None = None,
        severity: str = "WARNING",
        target_key: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> str:
        exception_id = f"exc-{uuid.uuid4().hex[:16]}"
        self.connection.execute(
            """INSERT INTO exceptions(
                   exception_id, project_id, run_id, exception_type, severity,
                   target_key, message, details_json, created_at)
               VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                exception_id, project_id, run_id, exception_type, severity,
                target_key, message,
                json.dumps(details or {}, ensure_ascii=False, sort_keys=True), _now(),
            ),
        )
        self.connection.commit()
        return exception_id

    def resolve_open_exceptions(
        self,
        project_id: str,
        exception_type: str,
        keep_target_keys: set[str] | None = None,
        note: str = "",
    ) -> int:
        """Close open exception records not present in the latest audit result.

        Exception rows are historical audit records, so resolving them keeps
        the evidence while preventing a later run from showing stale risks as
        current.  ``keep_target_keys`` identifies the issues still active in
        the latest run.
        """
        keep = keep_target_keys or set()
        rows = self.connection.execute(
            """SELECT exception_id, target_key FROM exceptions
               WHERE project_id=? AND exception_type=? AND status='OPEN'""",
            (project_id, exception_type),
        ).fetchall()
        stale = [row["exception_id"] for row in rows if row["target_key"] not in keep]
        if not stale:
            return 0
        timestamp = _now()
        self.connection.executemany(
            """UPDATE exceptions
               SET status='RESOLVED', resolved_at=?, resolution_note=?
               WHERE exception_id=? AND status='OPEN'""",
            [(timestamp, note, exception_id) for exception_id in stale],
        )
        self.connection.commit()
        return len(stale)

    def health(self) -> dict[str, Any]:
        version = self.connection.execute(
            "SELECT value FROM schema_meta WHERE key='schema_version'"
        ).fetchone()
        tables = {}
        for table in (
            "projects", "runs", "source_files", "person_master",
            "timesheet_facts", "cost_facts", "adjustments", "exceptions", "approvals",
            "timesheet_candidates", "review_queue", "po_master_records",
            "monthly_closing_facts", "cost_fact_po_allocations",
        ):
            row = self.connection.execute(f"SELECT COUNT(*) AS count FROM {table}").fetchone()
            tables[table] = row["count"]
        return {
            "db_path": str(self.db_path),
            "exists": self.db_path.exists(),
            "schema_version": int(version["value"]) if version else None,
            "tables": tables,
        }


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="IAV local fact store")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("init", help="initialize SQLite schema")
    register = sub.add_parser("register", help="register configured projects")
    register.add_argument("--project", default=None)
    sub.add_parser("health", help="show schema and table counts")
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    from project_registry import build_profile, list_profiles

    with FactStore() as store:
        store.initialize()
        if args.command == "register":
            profiles = [build_profile(args.project)] if args.project else list_profiles()
            for profile in profiles:
                store.register_project(profile)
            print(f"Registered {len(profiles)} project profile(s) in {store.db_path}")
        elif args.command == "health":
            print(json.dumps(store.health(), ensure_ascii=False, indent=2))
        else:
            print(f"Initialized {store.db_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
