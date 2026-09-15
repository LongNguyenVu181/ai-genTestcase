# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import sqlite3
import time
import shutil
from pathlib import Path
from typing import Any

BACKEND_DIR = Path(__file__).resolve().parent
# V1.9.2: keep data OUTSIDE versioned source folders so rebuilding/upgrading the app
# does not make projects/testcases appear to disappear. Override with TESTPILOT_DB_PATH if needed.
DEFAULT_DATA_DIR = Path(os.getenv("TESTPILOT_DATA_DIR", str(Path.home() / ".testpilot-ai")))
DEFAULT_DB_PATH = DEFAULT_DATA_DIR / "testpilot.db"
DB_PATH = Path(os.getenv("TESTPILOT_DB_PATH", str(DEFAULT_DB_PATH))).expanduser()
DB_PATH.parent.mkdir(parents=True, exist_ok=True)
LEGACY_DB_PATH = BACKEND_DIR / "data" / "testpilot.db"
if not DB_PATH.exists() and LEGACY_DB_PATH.exists() and LEGACY_DB_PATH.resolve() != DB_PATH.resolve():
    try:
        shutil.copy2(LEGACY_DB_PATH, DB_PATH)
    except Exception:
        pass


def _connect() -> sqlite3.Connection:
    # Background analysis writes progress while the UI polls/reads SQLite. A longer
    # busy timeout prevents transient "database is locked" failures on Windows.
    conn = sqlite3.connect(DB_PATH, timeout=30.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA busy_timeout = 30000")
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

            CREATE INDEX IF NOT EXISTS idx_testcases_project_tree
              ON testcases(project_id, scope, screen, folder_id);

            CREATE TABLE IF NOT EXISTS analysis_jobs (
                job_id TEXT PRIMARY KEY,
                kind TEXT NOT NULL,
                project_id TEXT NOT NULL,
                folder_id TEXT NOT NULL,
                status TEXT NOT NULL,
                stage TEXT NOT NULL DEFAULT 'queued',
                message TEXT NOT NULL DEFAULT '',
                progress INTEGER NOT NULL DEFAULT 0,
                run_id TEXT,
                source_names_json TEXT NOT NULL DEFAULT '[]',
                error_json TEXT NOT NULL DEFAULT '{}',
                created_at REAL NOT NULL,
                updated_at REAL NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_analysis_jobs_status
              ON analysis_jobs(status, updated_at);

            CREATE TABLE IF NOT EXISTS projects (
                project_id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                description TEXT NOT NULL DEFAULT '',
                project_type TEXT NOT NULL DEFAULT 'both',
                data_json TEXT NOT NULL DEFAULT '{}',
                created_at REAL NOT NULL,
                updated_at REAL NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_projects_updated
              ON projects(updated_at DESC);
            """
        )


def upsert_project(project: dict) -> dict:
    project_id = str(project.get("id") or project.get("project_id") or "").strip()
    if not project_id:
        raise ValueError("project id is required")
    name = str(project.get("name") or project_id).strip()
    description = str(project.get("description") or "").strip()
    project_type = str(project.get("type") or project.get("project_type") or "both").strip().lower()
    if project_type not in {"web", "api", "both"}:
        project_type = "both"
    created_at = float(project.get("createdAt") or project.get("created_at") or time.time())
    updated_at = float(project.get("updatedAt") or project.get("updated_at") or time.time())
    payload = dict(project)
    payload["id"] = project_id
    payload["name"] = name
    payload["description"] = description
    payload["type"] = project_type
    payload["createdAt"] = created_at
    payload["updatedAt"] = updated_at
    with _connect() as conn:
        conn.execute(
            """
            INSERT INTO projects(project_id, name, description, project_type, data_json, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(project_id) DO UPDATE SET
              name=excluded.name,
              description=excluded.description,
              project_type=excluded.project_type,
              data_json=excluded.data_json,
              updated_at=excluded.updated_at
            """,
            (project_id, name, description, project_type, json.dumps(payload, ensure_ascii=False), created_at, updated_at),
        )
    return get_project(project_id) or payload


def get_project(project_id: str) -> dict | None:
    with _connect() as conn:
        row = conn.execute("SELECT * FROM projects WHERE project_id=?", (project_id,)).fetchone()
    if not row:
        return None
    try:
        data = json.loads(row["data_json"] or "{}")
    except Exception:
        data = {}
    data.update({
        "id": row["project_id"],
        "name": row["name"],
        "description": row["description"],
        "type": row["project_type"],
        "createdAt": row["created_at"],
        "updatedAt": row["updated_at"],
    })
    return data


def list_projects() -> list[dict]:
    # Read all project metadata in one SQLite query. The previous implementation
    # opened a new connection for every row, which becomes noticeably slower over
    # a Cloudflare tunnel and on Windows antivirus-scanned folders.
    with _connect() as conn:
        rows = conn.execute(
            """SELECT project_id, name, description, project_type, data_json, created_at, updated_at
               FROM projects
               ORDER BY updated_at DESC, created_at DESC"""
        ).fetchall()
    projects: list[dict] = []
    for row in rows:
        try:
            data = json.loads(row["data_json"] or "{}")
        except Exception:
            data = {}
        data.update({
            "id": row["project_id"],
            "name": row["name"],
            "description": row["description"],
            "type": row["project_type"],
            "createdAt": row["created_at"],
            "updatedAt": row["updated_at"],
        })
        projects.append(data)
    return projects


def delete_project(project_id: str) -> bool:
    with _connect() as conn:
        cur = conn.execute("DELETE FROM projects WHERE project_id=?", (project_id,))
        return cur.rowcount > 0


def create_analysis_job(
    *,
    job_id: str,
    kind: str,
    project_id: str,
    folder_id: str,
    source_names: list[str],
) -> None:
    now = time.time()
    with _connect() as conn:
        conn.execute(
            """
            INSERT INTO analysis_jobs
            (job_id, kind, project_id, folder_id, status, stage, message, progress,
             run_id, source_names_json, error_json, created_at, updated_at)
            VALUES (?, ?, ?, ?, 'queued', 'queued', ?, 0, NULL, ?, '{}', ?, ?)
            """,
            (
                job_id, kind, project_id, folder_id,
                "Đã tiếp nhận tài liệu. Đang xếp hàng phân tích.",
                json.dumps(source_names, ensure_ascii=False),
                now, now,
            ),
        )


def update_analysis_job(
    job_id: str,
    *,
    status: str | None = None,
    stage: str | None = None,
    message: str | None = None,
    progress: int | None = None,
    run_id: str | None = None,
    error_payload: dict | None = None,
) -> None:
    fields: list[str] = ["updated_at=?"]
    values: list[Any] = [time.time()]
    if status is not None:
        fields.append("status=?")
        values.append(status)
    if stage is not None:
        fields.append("stage=?")
        values.append(stage)
    if message is not None:
        fields.append("message=?")
        values.append(message)
    if progress is not None:
        fields.append("progress=?")
        values.append(max(0, min(100, int(progress))))
    if run_id is not None:
        fields.append("run_id=?")
        values.append(run_id)
    if error_payload is not None:
        fields.append("error_json=?")
        values.append(json.dumps(error_payload, ensure_ascii=False))
    values.append(job_id)
    with _connect() as conn:
        conn.execute(f"UPDATE analysis_jobs SET {', '.join(fields)} WHERE job_id=?", values)


def get_analysis_job(job_id: str) -> dict | None:
    with _connect() as conn:
        row = conn.execute("SELECT * FROM analysis_jobs WHERE job_id=?", (job_id,)).fetchone()
    if not row:
        return None
    try:
        source_names = json.loads(row["source_names_json"] or "[]")
    except Exception:
        source_names = []
    try:
        error_payload = json.loads(row["error_json"] or "{}")
    except Exception:
        error_payload = {}
    return {
        "job_id": row["job_id"],
        "kind": row["kind"],
        "project_id": row["project_id"],
        "folder_id": row["folder_id"],
        "status": row["status"],
        "stage": row["stage"],
        "message": row["message"],
        "progress": row["progress"],
        "run_id": row["run_id"],
        "source_names": source_names,
        "error": error_payload,
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }


def fail_orphaned_analysis_jobs() -> int:
    # A process restart destroys in-memory background tasks. Make that state explicit
    # instead of leaving jobs forever in queued/processing.
    now = time.time()
    payload = json.dumps({"message": "Server đã khởi động lại khi tác vụ đang chạy. Vui lòng phân tích lại."}, ensure_ascii=False)
    with _connect() as conn:
        cur = conn.execute(
            """
            UPDATE analysis_jobs
            SET status='failed', stage='interrupted',
                message='Tác vụ bị gián đoạn do server khởi động lại.',
                error_json=?, updated_at=?
            WHERE status IN ('queued', 'processing')
            """,
            (payload, now),
        )
        return int(cur.rowcount)


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



def list_project_tree_summary(project_id: str) -> list[dict]:
    """Fast explorer summary without deserializing every testcase row."""
    with _connect() as conn:
        rows = conn.execute(
            """
            SELECT scope, screen, folder_id, COUNT(*) AS testcase_count
            FROM testcases
            WHERE project_id=?
            GROUP BY scope, screen, folder_id
            ORDER BY scope, screen COLLATE NOCASE, folder_id
            """,
            (project_id,),
        ).fetchall()
    return [
        {
            "scope": row["scope"],
            "screen": row["screen"],
            "folder_id": row["folder_id"],
            "count": int(row["testcase_count"] or 0),
        }
        for row in rows
    ]


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
# Normalize system-managed testcase IDs at startup.
renumber_all_testcase_ids()
# Any queued/processing job belonged to a previous process and cannot still be running.
fail_orphaned_analysis_jobs()
