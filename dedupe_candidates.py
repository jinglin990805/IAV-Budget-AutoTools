# -*- coding: utf-8 -*-
"""整理候选层的完全重复记录，保留审计链，不删除任何记录。"""

from __future__ import annotations

import config
from fact_store import FactStore


def dedupe_exact_candidates(project_id: str) -> dict[str, int]:
    with FactStore() as store:
        store.initialize()
        rows = store.connection.execute(
            """SELECT source_id, COUNT(1) AS count
               FROM timesheet_candidates WHERE project_id=?
               GROUP BY source_id HAVING COUNT(1)>1""",
            (project_id,),
        ).fetchall()
        superseded = 0
        for group in rows:
            candidates = store.connection.execute(
                """SELECT * FROM timesheet_candidates
                   WHERE project_id=? AND source_id=?
                   ORDER BY CASE WHEN po_id IS NOT NULL AND po_id!='' THEN 0 ELSE 1 END,
                            CASE WHEN total_hours IS NOT NULL THEN 0 ELSE 1 END,
                            created_at DESC""",
                (project_id, group["source_id"]),
            ).fetchall()
            preferred = candidates[0]
            for duplicate in candidates[1:]:
                if duplicate["review_status"] == "SUPERSEDED":
                    continue
                store.supersede_candidate(
                    duplicate["candidate_id"], preferred["candidate_id"],
                    "同一 source SHA-256 的完全重复解析候选，保留字段更完整/更新者",
                )
                superseded += 1
        return {"duplicate_groups": len(rows), "superseded": superseded}


if __name__ == "__main__":
    print(dedupe_exact_candidates("CNTS-25165"))
