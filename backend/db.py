# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import sqlite3
import time
from pathlib import Path
from typing import Any

BACKEND_DIR = Path(__file__).resolve().parent
DEFAULT_DB_PATH = BACKEND_DIR / "data" / "testpilot.db"
DB_PATH = Path(os.getenv("TESTPILOT_DB_PATH", str(DEFAULT_DB_PATH)))
DB_PATH.parent.mkdir(parents=True, exist_ok=True)


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn


def init_db() -> None:
    with _connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS analysis_runs (
                run_id TEXT PRIMARY KEY,
                kind TEXT NOT NULL,
                project_id TEXT NOT NULL,
                folder_id TEXT NOT NULL,
                created_at REAL NOT NULL,
                source_names_json TEXT NOT NULL,
                matrix_json TEXT NOT NULL,
                agent1_summary_json TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS testcases (
                record_id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id TEXT NOT NULL,
                project_id TEXT NOT NULL,
                folder_id TEXT NOT NULL,
                scope TEXT NOT NULL,
                sort_order INTEGER NOT NULL DEFAULT 0,

                tc_id TEXT NOT NULL,
                name TEXT NOT NULL,
                pre_condition TEXT NOT NULL DEFAULT '',
                importance TEXT NOT NULL DEFAULT '',
                steps_json TEXT NOT NULL DEFAULT '[]',
                test_data TEXT NOT NULL DEFAULT '',
                expected_result TEXT NOT NULL DEFAULT '',

                actual_result TEXT NOT NULL DEFAULT '',
                run1 TEXT NOT NULL DEFAULT '',
                run2 TEXT NOT NULL DEFAULT '',
                run3 TEXT NOT NULL DEFAULT '',
                current_result TEXT NOT NULL DEFAULT '',
                note TEXT NOT NULL DEFAULT '',
                error_code TEXT NOT NULL DEFAULT '',
                qc_writer TEXT NOT NULL DEFAULT '',
                sprint_writer TEXT NOT NULL DEFAULT '',
                qc_executor TEXT NOT NULL DEFAULT '',
                sprint_executor TEXT NOT NULL DEFAULT '',
                reviewer TEXT NOT NULL DEFAULT '',
                review_date TEXT NOT NULL DEFAULT '',
                review_content TEXT NOT NULL DEFAULT '',
                need_auto TEXT NOT NULL DEFAULT '',
                automated TEXT NOT NULL DEFAULT '',
                smoke TEXT NOT NULL DEFAULT '',
                regression TEXT NOT NULL DEFAULT '',
                outdated TEXT NOT NULL DEFAULT '',
                outdated_date TEXT NOT NULL DEFAULT '',

                type TEXT NOT NULL DEFAULT 'Chức năng',
                source_rule_id TEXT NOT NULL DEFAULT '',
                source_requirement TEXT NOT NULL DEFAULT '',
                feature_group TEXT NOT NULL DEFAULT '',
                feature_name TEXT NOT NULL DEFAULT '',
                category TEXT NOT NULL DEFAULT '',
                screen TEXT NOT NULL DEFAULT '',
                validation_json TEXT NOT NULL DEFAULT '{}',

                created_at REAL NOT NULL,
                updated_at REAL NOT NULL,

                FOREIGN KEY(run_id) REFERENCES analysis_runs(run_id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_testcases_folder
              ON testcases(project_id, folder_id, scope, sort_order);

            CREATE INDEX IF NOT EXISTS idx_testcases_run
              ON testcases(run_id, sort_order);
            """
        )


def save_run(
    *,
    run_id: str,
    kind: str,
    project_id: str,
    folder_id: str,
    source_names: list[str],
    matrix: dict,
    agent1_summary: dict,
) -> None:
    with _connect() as conn:
        conn.execute(
            """
            INSERT OR REPLACE INTO analysis_runs
            (run_id, kind, project_id, folder_id, created_at, source_names_json, matrix_json, agent1_summary_json)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                kind,
                project_id,
                folder_id,
                time.time(),
                json.dumps(source_names, ensure_ascii=False),
                json.dumps(matrix, ensure_ascii=False),
                json.dumps(agent1_summary, ensure_ascii=False),
            ),
        )


def _tc_value(tc: dict, key: str, default: Any = "") -> Any:
    value = tc.get(key, default)
    return default if value is None else value


def _insert_testcase_row(conn: sqlite3.Connection, *, run_id: str, project_id: str, folder_id: str, scope: str, sort_order: int, tc: dict) -> int:
    now = time.time()
    columns = [
        "run_id", "project_id", "folder_id", "scope", "sort_order",
        "tc_id", "name", "pre_condition", "importance", "steps_json", "test_data", "expected_result",
        "actual_result", "run1", "run2", "run3", "current_result", "note", "error_code",
        "qc_writer", "sprint_writer", "qc_executor", "sprint_executor", "reviewer",
        "review_date", "review_content", "need_auto", "automated", "smoke", "regression",
        "outdated", "outdated_date",
        "type", "source_rule_id", "source_requirement", "feature_group", "feature_name",
        "category", "screen", "validation_json", "created_at", "updated_at",
    ]
    values = [
        run_id, project_id, folder_id, scope, sort_order,
        _tc_value(tc, "id"), _tc_value(tc, "name"), _tc_value(tc, "preCondition"),
        _tc_value(tc, "importance"), json.dumps(_tc_value(tc, "steps", []), ensure_ascii=False),
        _tc_value(tc, "testData"), _tc_value(tc, "expectedResult"),
        _tc_value(tc, "actualResult"), _tc_value(tc, "run1"), _tc_value(tc, "run2"),
        _tc_value(tc, "run3"), _tc_value(tc, "currentResult"), _tc_value(tc, "note"),
        _tc_value(tc, "errorCode"), _tc_value(tc, "qcWriter"), _tc_value(tc, "sprintWriter"),
        _tc_value(tc, "qcExecutor"), _tc_value(tc, "sprintExecutor"), _tc_value(tc, "reviewer"),
        _tc_value(tc, "reviewDate"), _tc_value(tc, "reviewContent"), _tc_value(tc, "needAuto"),
        _tc_value(tc, "automated"), _tc_value(tc, "smoke"), _tc_value(tc, "regression"),
        _tc_value(tc, "outdated"), _tc_value(tc, "outdatedDate"),
        _tc_value(tc, "type", "Chức năng"), _tc_value(tc, "sourceRuleId"),
        _tc_value(tc, "sourceRequirement"), _tc_value(tc, "featureGroup"),
        _tc_value(tc, "featureName"), _tc_value(tc, "category"), _tc_value(tc, "screen"),
        json.dumps(_tc_value(tc, "validation", {}), ensure_ascii=False), now, now,
    ]
    placeholders = ",".join("?" for _ in columns)
    sql = f"INSERT INTO testcases ({','.join(columns)}) VALUES ({placeholders})"
    cur = conn.execute(sql, values)
    return int(cur.lastrowid)


def replace_run_testcases(
    *,
    run_id: str,
    project_id: str,
    folder_id: str,
    scope: str,
    testcases: list[dict],
) -> list[dict]:
    with _connect() as conn:
        conn.execute(
            "DELETE FROM testcases WHERE project_id=? AND folder_id=? AND scope=?",
            (project_id, folder_id, scope),
        )
        for idx, tc in enumerate(testcases):
            _insert_testcase_row(
                conn, run_id=run_id, project_id=project_id, folder_id=folder_id,
                scope=scope, sort_order=idx, tc=tc
            )
        _renumber_folder_conn(conn, project_id, folder_id, scope)
    return list_run_testcases(run_id)

def _row_to_tc(row: sqlite3.Row) -> dict:
    try:
        steps = json.loads(row["steps_json"] or "[]")
    except Exception:
        steps = []
    try:
        validation = json.loads(row["validation_json"] or "{}")
    except Exception:
        validation = {}
    return {
        "recordId": row["record_id"],
        "runId": row["run_id"],
        "projectId": row["project_id"],
        "folderId": row["folder_id"],
        "scope": row["scope"],
        "sortOrder": row["sort_order"],
        "id": row["tc_id"],
        "name": row["name"],
        "preCondition": row["pre_condition"],
        "importance": row["importance"],
        "steps": steps,
        "testData": row["test_data"],
        "expectedResult": row["expected_result"],
        "actualResult": row["actual_result"],
        "run1": row["run1"],
        "run2": row["run2"],
        "run3": row["run3"],
        "currentResult": row["current_result"],
        "note": row["note"],
        "errorCode": row["error_code"],
        "qcWriter": row["qc_writer"],
        "sprintWriter": row["sprint_writer"],
        "qcExecutor": row["qc_executor"],
        "sprintExecutor": row["sprint_executor"],
        "reviewer": row["reviewer"],
        "reviewDate": row["review_date"],
        "reviewContent": row["review_content"],
        "needAuto": row["need_auto"],
        "automated": row["automated"],
        "smoke": row["smoke"],
        "regression": row["regression"],
        "outdated": row["outdated"],
        "outdatedDate": row["outdated_date"],
        "type": row["type"],
        "sourceRuleId": row["source_rule_id"],
        "sourceRequirement": row["source_requirement"],
        "featureGroup": row["feature_group"],
        "featureName": row["feature_name"],
        "category": row["category"],
        "screen": row["screen"],
        "validation": validation,
    }


def get_run(run_id: str) -> dict | None:
    with _connect() as conn:
        row = conn.execute("SELECT * FROM analysis_runs WHERE run_id=?", (run_id,)).fetchone()
    if not row:
        return None
    return {
        "run_id": row["run_id"],
        "kind": row["kind"],
        "project_id": row["project_id"],
        "folder_id": row["folder_id"],
        "created_at": row["created_at"],
        "source_names": json.loads(row["source_names_json"] or "[]"),
        "matrix": json.loads(row["matrix_json"] or "{}"),
        "agent1_summary": json.loads(row["agent1_summary_json"] or "{}"),
    }


def list_run_testcases(run_id: str) -> list[dict]:
    with _connect() as conn:
        rows = conn.execute(
            "SELECT * FROM testcases WHERE run_id=? ORDER BY sort_order, record_id",
            (run_id,),
        ).fetchall()
    return [_row_to_tc(row) for row in rows]


def list_folder_testcases(project_id: str, folder_id: str, scope: str) -> list[dict]:
    with _connect() as conn:
        rows = conn.execute(
            """
            SELECT * FROM testcases
            WHERE project_id=? AND folder_id=? AND scope=?
            ORDER BY sort_order, record_id
            """,
            (project_id, folder_id, scope),
        ).fetchall()
    return [_row_to_tc(row) for row in rows]



def list_project_testcases(project_id: str, scope: str | None = None) -> list[dict]:
    with _connect() as conn:
        if scope:
            rows = conn.execute(
                """
                SELECT * FROM testcases
                WHERE project_id=? AND scope=?
                ORDER BY scope, screen COLLATE NOCASE, folder_id, sort_order, record_id
                """,
                (project_id, scope),
            ).fetchall()
        else:
            rows = conn.execute(
                """
                SELECT * FROM testcases
                WHERE project_id=?
                ORDER BY scope, screen COLLATE NOCASE, folder_id, sort_order, record_id
                """,
                (project_id,),
            ).fetchall()
    return [_row_to_tc(row) for row in rows]


def get_latest_run_for_folder(project_id: str, folder_id: str, scope: str) -> dict | None:
    with _connect() as conn:
        row = conn.execute(
            """
            SELECT * FROM analysis_runs
            WHERE project_id=? AND folder_id=? AND kind=?
            ORDER BY created_at DESC LIMIT 1
            """,
            (project_id, folder_id, scope),
        ).fetchone()
    if not row:
        return None
    return {
        "run_id": row["run_id"],
        "kind": row["kind"],
        "project_id": row["project_id"],
        "folder_id": row["folder_id"],
        "created_at": row["created_at"],
        "source_names": json.loads(row["source_names_json"] or "[]"),
        "matrix": json.loads(row["matrix_json"] or "{}"),
        "agent1_summary": json.loads(row["agent1_summary_json"] or "{}"),
    }



def get_latest_run_for_project_scope(project_id: str, scope: str) -> dict | None:
    with _connect() as conn:
        row = conn.execute(
            """
            SELECT * FROM analysis_runs
            WHERE project_id=? AND kind=?
            ORDER BY created_at DESC LIMIT 1
            """,
            (project_id, scope),
        ).fetchone()
    if not row:
        return None
    return {
        "run_id": row["run_id"],
        "kind": row["kind"],
        "project_id": row["project_id"],
        "folder_id": row["folder_id"],
        "created_at": row["created_at"],
        "source_names": json.loads(row["source_names_json"] or "[]"),
        "matrix": json.loads(row["matrix_json"] or "{}"),
        "agent1_summary": json.loads(row["agent1_summary_json"] or "{}"),
    }


def _renumber_folder_conn(conn: sqlite3.Connection, project_id: str, folder_id: str, scope: str) -> None:
    rows = conn.execute(
        """
        SELECT record_id FROM testcases
        WHERE project_id=? AND folder_id=? AND scope=?
        ORDER BY sort_order, record_id
        """,
        (project_id, folder_id, scope),
    ).fetchall()
    for idx, row in enumerate(rows, start=1):
        conn.execute(
            "UPDATE testcases SET tc_id=?, updated_at=? WHERE record_id=?",
            (f"TC_{idx:03d}", time.time(), row["record_id"]),
        )


def renumber_folder_testcase_ids(project_id: str, folder_id: str, scope: str) -> None:
    with _connect() as conn:
        _renumber_folder_conn(conn, project_id, folder_id, scope)


def renumber_all_testcase_ids() -> None:
    with _connect() as conn:
        groups = conn.execute(
            "SELECT DISTINCT project_id, folder_id, scope FROM testcases"
        ).fetchall()
        for group in groups:
            _renumber_folder_conn(conn, group["project_id"], group["folder_id"], group["scope"])


def _update_columns_from_tc(payload: dict) -> dict[str, Any]:
    mapping = {
        # tc_id is system-managed (TC_001..TC_NNN) and is intentionally not editable.
        "name": "name",
        "preCondition": "pre_condition",
        "importance": "importance",
        "testData": "test_data",
        "expectedResult": "expected_result",
        "actualResult": "actual_result",
        "run1": "run1",
        "run2": "run2",
        "run3": "run3",
        "currentResult": "current_result",
        "note": "note",
        "errorCode": "error_code",
        "qcWriter": "qc_writer",
        "sprintWriter": "sprint_writer",
        "qcExecutor": "qc_executor",
        "sprintExecutor": "sprint_executor",
        "reviewer": "reviewer",
        "reviewDate": "review_date",
        "reviewContent": "review_content",
        "needAuto": "need_auto",
        "automated": "automated",
        "smoke": "smoke",
        "regression": "regression",
        "outdated": "outdated",
        "outdatedDate": "outdated_date",
        "type": "type",
        "sourceRuleId": "source_rule_id",
        "sourceRequirement": "source_requirement",
        "featureGroup": "feature_group",
        "featureName": "feature_name",
        "category": "category",
        "screen": "screen",
        "sortOrder": "sort_order",
    }
    updates: dict[str, Any] = {}
    for src, dst in mapping.items():
        if src in payload:
            updates[dst] = payload[src]
    if "steps" in payload:
        updates["steps_json"] = json.dumps(payload["steps"] or [], ensure_ascii=False)
    if "validation" in payload:
        updates["validation_json"] = json.dumps(payload["validation"] or {}, ensure_ascii=False)
    return updates


def update_testcase(record_id: int, payload: dict) -> dict | None:
    updates = _update_columns_from_tc(payload)
    updates["updated_at"] = time.time()
    if not updates:
        return get_testcase(record_id)

    cols = ", ".join(f"{col}=?" for col in updates)
    values = list(updates.values()) + [record_id]
    with _connect() as conn:
        cur = conn.execute(f"UPDATE testcases SET {cols} WHERE record_id=?", values)
        if cur.rowcount <= 0:
            return None
    return get_testcase(record_id)


def get_testcase(record_id: int) -> dict | None:
    with _connect() as conn:
        row = conn.execute("SELECT * FROM testcases WHERE record_id=?", (record_id,)).fetchone()
    return _row_to_tc(row) if row else None


def create_testcase(
    *,
    project_id: str,
    folder_id: str,
    scope: str,
    run_id: str,
    payload: dict,
) -> dict:
    with _connect() as conn:
        max_order = conn.execute(
            """
            SELECT COALESCE(MAX(sort_order), -1) AS max_order
            FROM testcases WHERE project_id=? AND folder_id=? AND scope=?
            """,
            (project_id, folder_id, scope),
        ).fetchone()["max_order"]
        record_id = _insert_testcase_row(
            conn, run_id=run_id, project_id=project_id, folder_id=folder_id,
            scope=scope, sort_order=int(max_order) + 1, tc=payload
        )
        _renumber_folder_conn(conn, project_id, folder_id, scope)
    return get_testcase(record_id)

def delete_testcase(record_id: int) -> bool:
    with _connect() as conn:
        row = conn.execute(
            "SELECT project_id, folder_id, scope FROM testcases WHERE record_id=?",
            (record_id,),
        ).fetchone()
        if not row:
            return False
        cur = conn.execute("DELETE FROM testcases WHERE record_id=?", (record_id,))
        if cur.rowcount > 0:
            _renumber_folder_conn(conn, row["project_id"], row["folder_id"], row["scope"])
            return True
        return False


init_db()
# One-time-safe migration for existing V1.3/V1.5 databases.
renumber_all_testcase_ids()
