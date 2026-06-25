from __future__ import annotations

import json
import re
import secrets
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any


TASK_STATUSES = ("대기", "진행중", "완료", "보류")
TASK_PRIORITIES = ("낮음", "보통", "높음")


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def init_crm_db(connection: sqlite3.Connection) -> None:
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS crm_accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            name TEXT NOT NULL UNIQUE,
            account_type TEXT,
            contact_name TEXT,
            phone TEXT,
            email TEXT,
            memo TEXT,
            active INTEGER NOT NULL DEFAULT 1
        )
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS crm_tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            public_id TEXT UNIQUE,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            account_id INTEGER,
            account_name TEXT,
            title TEXT NOT NULL,
            description TEXT,
            assignee_user_id INTEGER,
            assignee_name TEXT,
            requester_user_id INTEGER,
            requester_name TEXT,
            due_at TEXT,
            status TEXT NOT NULL DEFAULT '대기',
            priority TEXT NOT NULL DEFAULT '보통',
            source TEXT NOT NULL DEFAULT 'app',
            source_message_event_id INTEGER,
            completed_at TEXT
        )
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS crm_task_comments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id INTEGER NOT NULL,
            created_at TEXT NOT NULL,
            author_user_id INTEGER,
            author_name TEXT,
            comment_type TEXT,
            body TEXT
        )
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS crm_saved_views (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            scope TEXT NOT NULL DEFAULT 'tasks',
            filters_json TEXT NOT NULL,
            sort_key TEXT,
            is_default INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(user_id, scope, name)
        )
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS crm_daily_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            log_date TEXT NOT NULL,
            user_id INTEGER NOT NULL,
            employee_name TEXT,
            author_user_id INTEGER,
            author_name TEXT,
            work_summary TEXT,
            completed_work TEXT,
            ongoing_work TEXT,
            blockers TEXT,
            next_plan TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(log_date, user_id)
        )
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS crm_message_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            platform TEXT NOT NULL,
            sender_key TEXT,
            sender_name TEXT,
            text TEXT,
            payload_json TEXT,
            result TEXT NOT NULL,
            error TEXT,
            task_id INTEGER
        )
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS crm_messenger_users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            platform TEXT NOT NULL,
            sender_key TEXT NOT NULL,
            display_name TEXT,
            user_id INTEGER NOT NULL,
            username TEXT,
            active INTEGER NOT NULL DEFAULT 1,
            UNIQUE(platform, sender_key)
        )
        """
    )
    connection.execute("CREATE INDEX IF NOT EXISTS idx_crm_accounts_name ON crm_accounts(name)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_crm_tasks_status ON crm_tasks(status)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_crm_tasks_due ON crm_tasks(due_at)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_crm_tasks_assignee ON crm_tasks(assignee_user_id)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_crm_task_comments_task ON crm_task_comments(task_id)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_crm_message_events_created ON crm_message_events(created_at)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_crm_saved_views_user_scope ON crm_saved_views(user_id, scope)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_crm_daily_logs_date ON crm_daily_logs(log_date)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_crm_daily_logs_user_date ON crm_daily_logs(user_id, log_date)")


def ensure_webhook_token(path: Path) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        token = path.read_text(encoding="utf-8").strip()
        if len(token) >= 24:
            return token
    token = secrets.token_urlsafe(32)
    path.write_text(token, encoding="utf-8")
    return token


def rotate_webhook_token(path: Path) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    token = secrets.token_urlsafe(32)
    path.write_text(token, encoding="utf-8")
    return token


def _connect(db_path: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    return connection


def _row_dict(row: sqlite3.Row | None) -> dict[str, Any] | None:
    return dict(row) if row else None


def _rows(rows: list[sqlite3.Row]) -> list[dict[str, Any]]:
    return [dict(row) for row in rows]


def _clean(value: object) -> str:
    return str(value or "").strip()


def _user_name(user: dict[str, Any]) -> str:
    return _clean(user.get("display_name")) or _clean(user.get("username")) or "사용자"


def _user_id(user: dict[str, Any]) -> int | None:
    try:
        return int(user.get("id") or 0) or None
    except (TypeError, ValueError):
        return None


def _find_user(connection: sqlite3.Connection, value: object) -> sqlite3.Row | None:
    text = _clean(value)
    if not text:
        return None
    if text.isdigit():
        row = connection.execute(
            "SELECT id, username, display_name FROM users WHERE id = ? AND active = 1",
            (int(text),),
        ).fetchone()
        if row:
            return row
    return connection.execute(
        """
        SELECT id, username, display_name
          FROM users
         WHERE active = 1
           AND (username = ? OR display_name = ? OR display_name LIKE ?)
         ORDER BY CASE WHEN username = ? OR display_name = ? THEN 0 ELSE 1 END, id
         LIMIT 1
        """,
        (text, text, f"%{text}%", text, text),
    ).fetchone()


def _ensure_account(connection: sqlite3.Connection, name: str) -> sqlite3.Row:
    account_name = _clean(name)
    if not account_name:
        raise ValueError("거래처명을 입력해주세요.")
    existing = connection.execute("SELECT * FROM crm_accounts WHERE name = ?", (account_name,)).fetchone()
    if existing:
        return existing
    now = now_text()
    cursor = connection.execute(
        """
        INSERT INTO crm_accounts (created_at, updated_at, name, account_type, active)
        VALUES (?, ?, ?, '거래처', 1)
        """,
        (now, now, account_name),
    )
    return connection.execute("SELECT * FROM crm_accounts WHERE id = ?", (cursor.lastrowid,)).fetchone()


def _task_public_id(task_id: int) -> str:
    return f"TASK-{task_id:04d}"


def _normalize_public_id(value: object) -> str:
    text = _clean(value).upper()
    match = re.search(r"(?:TASK-?)?(\d+)", text)
    if not match:
        return text
    return _task_public_id(int(match.group(1)))


def _task_by_public_id(connection: sqlite3.Connection, value: object) -> sqlite3.Row | None:
    public_id = _normalize_public_id(value)
    return connection.execute("SELECT * FROM crm_tasks WHERE public_id = ?", (public_id,)).fetchone()


def _normalize_status(value: object) -> str:
    status = _clean(value)
    if status in TASK_STATUSES:
        return status
    aliases = {
        "확인": "진행중",
        "진행": "진행중",
        "완료": "완료",
        "보류": "보류",
        "대기": "대기",
    }
    return aliases.get(status, "대기")


def _normalize_priority(value: object) -> str:
    priority = _clean(value)
    return priority if priority in TASK_PRIORITIES else "보통"


def _parse_due_text(value: object) -> str:
    text = _clean(value)
    if not text:
        return ""
    today = date.today()
    time_match = re.search(r"(\d{1,2})\s*시(?:\s*(\d{1,2})\s*분?)?", text)
    if not time_match:
        time_match = re.search(r"(?<!\d)(\d{1,2}):(\d{2})(?!\d)", text)
    hour = int(time_match.group(1)) if time_match else None
    minute = int(time_match.group(2) or 0) if time_match else 0
    if hour is not None:
        lowered = text.lower()
        if ("오후" in text or "pm" in lowered) and hour < 12:
            hour += 12
        elif ("오전" in text or "am" in lowered) and hour == 12:
            hour = 0
        elif "오전" not in text and "오후" not in text and "am" not in lowered and "pm" not in lowered and 1 <= hour <= 7:
            hour += 12

    target: date | None = None
    if "오늘" in text:
        target = today
    elif "내일" in text:
        target = today + timedelta(days=1)
    else:
        iso_match = re.search(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", text)
        short_match = re.search(r"(\d{1,2})\s*(?:/|월)\s*(\d{1,2})", text)
        try:
            if iso_match:
                target = date(int(iso_match.group(1)), int(iso_match.group(2)), int(iso_match.group(3)))
            elif short_match:
                target = date(today.year, int(short_match.group(1)), int(short_match.group(2)))
        except ValueError:
            target = None

    if not target:
        return text
    if hour is None:
        return target.isoformat()
    return f"{target.isoformat()} {hour:02d}:{minute:02d}"


def user_can_touch_task(task: sqlite3.Row, user: dict[str, Any], can_manage: bool) -> bool:
    if can_manage:
        return True
    current_id = _user_id(user)
    if current_id and task["assignee_user_id"] == current_id:
        return True
    user_names = {_clean(user.get("username")), _clean(user.get("display_name"))}
    return bool(_clean(task["assignee_name"]) in user_names)


def _normalize_log_date(value: object) -> str:
    text = _clean(value)
    if not text:
        return date.today().isoformat()
    try:
        return date.fromisoformat(text[:10]).isoformat()
    except ValueError as exc:
        raise ValueError("일지 날짜는 YYYY-MM-DD 형식으로 입력해주세요.") from exc


def _staff_display_name(row: sqlite3.Row | dict[str, Any]) -> str:
    return _clean(row["display_name"]) or _clean(row["username"]) or f"직원 {row['id']}"


def _daily_log_public(
    row: sqlite3.Row | None,
    staff: sqlite3.Row | dict[str, Any],
    log_date: str,
    task_counts: dict[int, dict[str, int]],
) -> dict[str, Any]:
    user_id = int(staff["id"])
    counts = task_counts.get(user_id, {})
    if row:
        item = dict(row)
    else:
        item = {
            "id": "",
            "log_date": log_date,
            "user_id": user_id,
            "employee_name": _staff_display_name(staff),
            "author_user_id": "",
            "author_name": "",
            "work_summary": "",
            "completed_work": "",
            "ongoing_work": "",
            "blockers": "",
            "next_plan": "",
            "created_at": "",
            "updated_at": "",
        }
    item["user_id"] = user_id
    item["username"] = _clean(staff["username"])
    item["display_name"] = _staff_display_name(staff)
    item["role"] = _clean(staff["role"])
    item["submitted"] = bool(item.get("id"))
    item["open_tasks"] = int(counts.get("open_tasks", 0))
    item["completed_today"] = int(counts.get("completed_today", 0))
    item["due_today"] = int(counts.get("due_today", 0))
    return item


def _has_daily_log_issue(value: object) -> bool:
    text = _clean(value)
    if not text:
        return False
    normalized = re.sub(r"[\s./_-]+", "", text.lower())
    no_issue_markers = {
        "없음",
        "이슈없음",
        "특이사항없음",
        "특이이슈없음",
        "문제없음",
        "해당없음",
        "없습니다",
        "none",
        "no",
        "na",
        "n/a",
    }
    return normalized not in no_issue_markers


def list_crm_daily_logs(db_path: Path, log_date: object = "", user_id: int | None = None) -> dict[str, Any]:
    target_date = _normalize_log_date(log_date)
    connection = _connect(db_path)
    try:
        staff_params: list[Any] = []
        staff_where = "WHERE active = 1"
        if user_id:
            staff_where += " AND id = ?"
            staff_params.append(int(user_id))
        staff_rows = connection.execute(
            f"""
            SELECT id, username, display_name, role
              FROM users
             {staff_where}
             ORDER BY CASE role WHEN 'admin' THEN 0 WHEN 'sub_admin' THEN 1 ELSE 2 END,
                      display_name COLLATE NOCASE,
                      username COLLATE NOCASE
            """,
            staff_params,
        ).fetchall()
        log_params: list[Any] = [target_date]
        log_where = "WHERE logs.log_date = ?"
        if user_id:
            log_where += " AND logs.user_id = ?"
            log_params.append(int(user_id))
        log_rows = connection.execute(
            f"""
            SELECT logs.*
              FROM crm_daily_logs logs
             {log_where}
             ORDER BY logs.updated_at DESC, logs.id DESC
            """,
            log_params,
        ).fetchall()
        counts_rows = connection.execute(
            """
            SELECT assignee_user_id AS user_id,
                   SUM(CASE WHEN status != '완료' THEN 1 ELSE 0 END) AS open_tasks,
                   SUM(CASE WHEN status = '완료' AND substr(completed_at, 1, 10) = ? THEN 1 ELSE 0 END) AS completed_today,
                   SUM(CASE WHEN status != '완료' AND substr(due_at, 1, 10) = ? THEN 1 ELSE 0 END) AS due_today
              FROM crm_tasks
             WHERE assignee_user_id IS NOT NULL
             GROUP BY assignee_user_id
            """,
            (target_date, target_date),
        ).fetchall()
        logs_by_user = {int(row["user_id"]): row for row in log_rows}
        task_counts = {
            int(row["user_id"]): {
                "open_tasks": int(row["open_tasks"] or 0),
                "completed_today": int(row["completed_today"] or 0),
                "due_today": int(row["due_today"] or 0),
            }
            for row in counts_rows
            if row["user_id"] is not None
        }
        logs = [_daily_log_public(logs_by_user.get(int(staff["id"])), staff, target_date, task_counts) for staff in staff_rows]
        submitted = sum(1 for item in logs if item["submitted"])
        issue_count = sum(1 for item in logs if _has_daily_log_issue(item.get("blockers")))
        return {
            "date": target_date,
            "staff": [
                {
                    "id": int(staff["id"]),
                    "username": _clean(staff["username"]),
                    "display_name": _staff_display_name(staff),
                    "role": _clean(staff["role"]),
                }
                for staff in staff_rows
            ],
            "logs": logs,
            "summary": {
                "total_staff": len(logs),
                "submitted": submitted,
                "missing": max(len(logs) - submitted, 0),
                "issues": issue_count,
                "completed_today": sum(int(item.get("completed_today") or 0) for item in logs),
                "due_today": sum(int(item.get("due_today") or 0) for item in logs),
            },
        }
    finally:
        connection.close()


def save_crm_daily_log(
    db_path: Path,
    payload: dict[str, Any],
    user: dict[str, Any],
    can_manage: bool = False,
) -> int:
    current_user_id = _user_id(user)
    target_user_id = int(payload.get("user_id") or current_user_id or 0)
    if not target_user_id:
        raise ValueError("일지를 작성할 직원을 선택해주세요.")
    if not can_manage and target_user_id != current_user_id:
        raise PermissionError("본인 일지만 작성할 수 있습니다.")
    log_date = _normalize_log_date(payload.get("log_date"))
    fields = {
        "work_summary": _clean(payload.get("work_summary")),
        "completed_work": _clean(payload.get("completed_work")),
        "ongoing_work": _clean(payload.get("ongoing_work")),
        "blockers": _clean(payload.get("blockers")),
        "next_plan": _clean(payload.get("next_plan")),
    }
    if not any(fields.values()):
        raise ValueError("일지 내용을 하나 이상 입력해주세요.")
    connection = _connect(db_path)
    try:
        employee = connection.execute(
            "SELECT id, username, display_name FROM users WHERE id = ? AND active = 1",
            (target_user_id,),
        ).fetchone()
        if not employee:
            raise ValueError("일지를 작성할 직원을 찾지 못했습니다.")
        now = now_text()
        employee_name = _staff_display_name(employee)
        author_name = _user_name(user)
        existing = connection.execute(
            "SELECT id FROM crm_daily_logs WHERE log_date = ? AND user_id = ?",
            (log_date, target_user_id),
        ).fetchone()
        if existing:
            log_id = int(existing["id"])
            connection.execute(
                """
                UPDATE crm_daily_logs
                   SET employee_name = ?, author_user_id = ?, author_name = ?,
                       work_summary = ?, completed_work = ?, ongoing_work = ?,
                       blockers = ?, next_plan = ?, updated_at = ?
                 WHERE id = ?
                """,
                (
                    employee_name,
                    current_user_id,
                    author_name,
                    fields["work_summary"],
                    fields["completed_work"],
                    fields["ongoing_work"],
                    fields["blockers"],
                    fields["next_plan"],
                    now,
                    log_id,
                ),
            )
        else:
            cursor = connection.execute(
                """
                INSERT INTO crm_daily_logs
                    (log_date, user_id, employee_name, author_user_id, author_name,
                     work_summary, completed_work, ongoing_work, blockers, next_plan,
                     created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    log_date,
                    target_user_id,
                    employee_name,
                    current_user_id,
                    author_name,
                    fields["work_summary"],
                    fields["completed_work"],
                    fields["ongoing_work"],
                    fields["blockers"],
                    fields["next_plan"],
                    now,
                    now,
                ),
            )
            log_id = int(cursor.lastrowid)
        connection.commit()
        return log_id
    finally:
        connection.close()


def crm_dashboard_payload(db_path: Path) -> dict[str, Any]:
    connection = _connect(db_path)
    try:
        today = date.today().isoformat()
        stats = {
            "accounts": connection.execute("SELECT COUNT(*) FROM crm_accounts WHERE active = 1").fetchone()[0],
            "open_tasks": connection.execute("SELECT COUNT(*) FROM crm_tasks WHERE status != '완료'").fetchone()[0],
            "due_today": connection.execute(
                "SELECT COUNT(*) FROM crm_tasks WHERE status != '완료' AND substr(due_at, 1, 10) = ?",
                (today,),
            ).fetchone()[0],
            "overdue": connection.execute(
                "SELECT COUNT(*) FROM crm_tasks WHERE status != '완료' AND length(due_at) >= 10 AND substr(due_at, 1, 10) < ?",
                (today,),
            ).fetchone()[0],
        }
        status_rows = connection.execute(
            "SELECT status, COUNT(*) AS count FROM crm_tasks WHERE status != '완료' GROUP BY status ORDER BY status"
        ).fetchall()
        priority_tasks = connection.execute(
            """
            SELECT *
              FROM crm_tasks
             WHERE status != '완료'
             ORDER BY CASE priority WHEN '높음' THEN 0 WHEN '보통' THEN 1 ELSE 2 END,
                      CASE WHEN due_at IS NULL OR due_at = '' THEN 1 ELSE 0 END,
                      due_at,
                      updated_at DESC
             LIMIT 8
            """
        ).fetchall()
        recent_events = connection.execute(
            """
            SELECT events.*
              FROM crm_message_events events
              LEFT JOIN crm_tasks tasks ON tasks.id = events.task_id
             WHERE events.task_id IS NULL OR tasks.status != '완료'
             ORDER BY events.id DESC
             LIMIT 8
            """
        ).fetchall()
        project_rows = connection.execute(
            """
            SELECT CASE
                     WHEN account_id IS NOT NULL THEN 'account:' || account_id
                     ELSE 'direct:' || COALESCE(NULLIF(account_name, ''), '직원 지시 업무')
                   END AS project_key,
                   COALESCE(NULLIF(account_name, ''), '직원 지시 업무') AS project_name,
                   COUNT(*) AS total_tasks,
                   SUM(CASE WHEN status != '완료' THEN 1 ELSE 0 END) AS open_tasks,
                   SUM(CASE WHEN status = '완료' THEN 1 ELSE 0 END) AS completed_tasks,
                   SUM(CASE WHEN status = '대기' THEN 1 ELSE 0 END) AS waiting_tasks,
                   SUM(CASE WHEN status = '진행중' THEN 1 ELSE 0 END) AS progress_tasks,
                   SUM(CASE WHEN status = '보류' THEN 1 ELSE 0 END) AS hold_tasks,
                   SUM(CASE WHEN status != '완료' AND substr(due_at, 1, 10) = ? THEN 1 ELSE 0 END) AS due_today,
                   SUM(CASE WHEN status != '완료' AND length(due_at) >= 10 AND substr(due_at, 1, 10) < ? THEN 1 ELSE 0 END) AS overdue,
                   SUM(CASE WHEN status != '완료' AND priority = '높음' THEN 1 ELSE 0 END) AS high_priority,
                   GROUP_CONCAT(DISTINCT NULLIF(assignee_name, '')) AS assignee_names,
                   MIN(CASE WHEN status != '완료' AND due_at IS NOT NULL AND due_at != '' THEN due_at END) AS next_due_at,
                   MAX(updated_at) AS latest_update
              FROM crm_tasks
             GROUP BY project_key, project_name
            HAVING SUM(CASE WHEN status != '완료' THEN 1 ELSE 0 END) > 0
             ORDER BY open_tasks DESC,
                      overdue DESC,
                      due_today DESC,
                      latest_update DESC
             LIMIT 12
            """,
            (today, today),
        ).fetchall()
        project_task_rows = connection.execute(
            """
            SELECT CASE
                     WHEN account_id IS NOT NULL THEN 'account:' || account_id
                     ELSE 'direct:' || COALESCE(NULLIF(account_name, ''), '직원 지시 업무')
                   END AS project_key,
                   *
              FROM crm_tasks
             WHERE status != '완료'
             ORDER BY CASE status WHEN '대기' THEN 0 WHEN '진행중' THEN 1 WHEN '보류' THEN 2 ELSE 3 END,
                      CASE priority WHEN '높음' THEN 0 WHEN '보통' THEN 1 ELSE 2 END,
                      CASE WHEN due_at IS NULL OR due_at = '' THEN 1 ELSE 0 END,
                      due_at,
                      updated_at DESC
             LIMIT 400
            """
        ).fetchall()
        project_tasks: dict[str, list[dict[str, Any]]] = {}
        for row in _rows(project_task_rows):
            project_tasks.setdefault(str(row.get("project_key") or ""), []).append(row)
        project_progress = []
        for row in _rows(project_rows):
            total_tasks = int(row.get("total_tasks") or 0)
            completed_tasks = int(row.get("completed_tasks") or 0)
            row["progress_percent"] = round((completed_tasks / total_tasks) * 100) if total_tasks else 0
            row["tasks"] = project_tasks.get(str(row.get("project_key") or ""), [])[:20]
            project_progress.append(row)
        return {
            "stats": stats,
            "status_counts": _rows(status_rows),
            "priority_tasks": _rows(priority_tasks),
            "recent_events": _rows(recent_events),
            "project_progress": project_progress,
        }
    finally:
        connection.close()


def list_crm_accounts(db_path: Path, query: str = "", limit: int = 200) -> list[dict[str, Any]]:
    connection = _connect(db_path)
    try:
        search = f"%{_clean(query)}%"
        params: list[Any] = []
        where = "WHERE accounts.active = 1"
        if _clean(query):
            where += " AND (accounts.name LIKE ? OR accounts.contact_name LIKE ? OR accounts.memo LIKE ?)"
            params.extend([search, search, search])
        params.append(max(1, min(int(limit), 1000)))
        rows = connection.execute(
            f"""
            SELECT accounts.*,
                   SUM(CASE WHEN tasks.status != '완료' THEN 1 ELSE 0 END) AS open_task_count,
                   COUNT(tasks.id) AS task_count
              FROM crm_accounts accounts
              LEFT JOIN crm_tasks tasks ON tasks.account_id = accounts.id
             {where}
             GROUP BY accounts.id
             ORDER BY accounts.updated_at DESC, accounts.id DESC
             LIMIT ?
            """,
            params,
        ).fetchall()
        return _rows(rows)
    finally:
        connection.close()


def save_crm_account(db_path: Path, payload: dict[str, Any]) -> int:
    account_id = int(payload.get("id") or 0)
    name = _clean(payload.get("name"))
    if not name:
        raise ValueError("거래처명을 입력해주세요.")
    values = {
        "name": name,
        "account_type": _clean(payload.get("account_type")) or "거래처",
        "contact_name": _clean(payload.get("contact_name")),
        "phone": _clean(payload.get("phone")),
        "email": _clean(payload.get("email")),
        "memo": _clean(payload.get("memo")),
        "active": 1,
    }
    connection = _connect(db_path)
    try:
        now = now_text()
        if account_id:
            connection.execute(
                """
                UPDATE crm_accounts
                   SET updated_at = ?, name = ?, account_type = ?, contact_name = ?,
                       phone = ?, email = ?, memo = ?, active = ?
                 WHERE id = ?
                """,
                (
                    now,
                    values["name"],
                    values["account_type"],
                    values["contact_name"],
                    values["phone"],
                    values["email"],
                    values["memo"],
                    values["active"],
                    account_id,
                ),
            )
        else:
            cursor = connection.execute(
                """
                INSERT INTO crm_accounts
                    (created_at, updated_at, name, account_type, contact_name, phone, email, memo, active)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    now,
                    now,
                    values["name"],
                    values["account_type"],
                    values["contact_name"],
                    values["phone"],
                    values["email"],
                    values["memo"],
                    values["active"],
                ),
            )
            account_id = int(cursor.lastrowid)
        connection.commit()
        return account_id
    finally:
        connection.close()


def list_crm_tasks(
    db_path: Path,
    query: str = "",
    status: str = "",
    assignee: str = "",
    assignee_user_id: int | None = None,
    priority: str = "",
    due: str = "",
    source: str = "",
    open_only: bool = False,
    sort: str = "smart",
    limit: int = 300,
) -> list[dict[str, Any]]:
    connection = _connect(db_path)
    try:
        today = date.today().isoformat()
        params: list[Any] = []
        conditions = ["1 = 1"]
        if _clean(query):
            search = f"%{_clean(query)}%"
            conditions.append(
                """
                (title LIKE ?
                 OR COALESCE(NULLIF(account_name, ''), '직원 지시 업무') LIKE ?
                 OR description LIKE ?
                 OR public_id LIKE ?
                 OR assignee_name LIKE ?
                 OR requester_name LIKE ?)
                """
            )
            params.extend([search, search, search, search, search, search])
        if _clean(status):
            conditions.append("status = ?")
            params.append(_normalize_status(status))
        elif open_only:
            conditions.append("status != '완료'")
        if _clean(assignee):
            search = f"%{_clean(assignee)}%"
            conditions.append("assignee_name LIKE ?")
            params.append(search)
        if assignee_user_id:
            conditions.append("assignee_user_id = ?")
            params.append(int(assignee_user_id))
        if _clean(priority):
            conditions.append("priority = ?")
            params.append(_normalize_priority(priority))
        due_filter = _clean(due)
        if due_filter == "today":
            conditions.append("length(due_at) >= 10 AND substr(due_at, 1, 10) = ?")
            params.append(today)
        elif due_filter == "overdue":
            conditions.append("length(due_at) >= 10 AND substr(due_at, 1, 10) < ?")
            params.append(today)
        elif due_filter == "upcoming":
            conditions.append("length(due_at) >= 10 AND substr(due_at, 1, 10) > ?")
            params.append(today)
        elif due_filter == "none":
            conditions.append("(due_at IS NULL OR due_at = '')")
        source_filter = _clean(source)
        if source_filter:
            if source_filter == "messenger":
                conditions.append("source LIKE 'messenger:%'")
            else:
                conditions.append("source = ?")
                params.append(source_filter)
        sort_key = _clean(sort) or "smart"
        order_by = """
            CASE status WHEN '대기' THEN 0 WHEN '진행중' THEN 1 WHEN '보류' THEN 2 ELSE 3 END,
            CASE WHEN due_at IS NULL OR due_at = '' THEN 1 ELSE 0 END,
            due_at,
            updated_at DESC
        """
        if sort_key == "due":
            order_by = """
                CASE WHEN due_at IS NULL OR due_at = '' THEN 1 ELSE 0 END,
                due_at,
                CASE priority WHEN '높음' THEN 0 WHEN '보통' THEN 1 ELSE 2 END,
                updated_at DESC
            """
        elif sort_key == "updated":
            order_by = "updated_at DESC, id DESC"
        params.append(max(1, min(int(limit), 2000)))
        rows = connection.execute(
            f"""
            SELECT *
              FROM crm_tasks
             WHERE {' AND '.join(conditions)}
             ORDER BY {order_by}
             LIMIT ?
            """,
            params,
        ).fetchall()
        return _rows(rows)
    finally:
        connection.close()


def list_crm_task_comments(db_path: Path, task_id: int, limit: int = 120) -> list[dict[str, Any]]:
    connection = _connect(db_path)
    try:
        task = connection.execute("SELECT id FROM crm_tasks WHERE id = ?", (int(task_id),)).fetchone()
        if not task:
            raise ValueError("업무를 찾지 못했습니다.")
        rows = connection.execute(
            """
            SELECT *
              FROM crm_task_comments
             WHERE task_id = ?
             ORDER BY id DESC
             LIMIT ?
            """,
            (int(task_id), max(1, min(int(limit), 300))),
        ).fetchall()
        return _rows(rows)
    finally:
        connection.close()


def _saved_view_scope(value: object) -> str:
    scope = _clean(value)
    return scope if scope in {"tasks"} else "tasks"


def _saved_view_filters(value: object) -> dict[str, Any]:
    if isinstance(value, dict):
        raw = value
    elif isinstance(value, str) and value.strip():
        raw = json.loads(value)
    else:
        raw = {}
    allowed = {"q", "status", "assignee_user_id", "priority", "due", "source", "open_only", "sort"}
    return {key: _clean(raw.get(key)) for key in allowed if key in raw}


def list_crm_saved_views(db_path: Path, user_id: int, scope: str = "tasks") -> list[dict[str, Any]]:
    connection = _connect(db_path)
    try:
        rows = connection.execute(
            """
            SELECT *
              FROM crm_saved_views
             WHERE user_id = ? AND scope = ?
             ORDER BY is_default DESC, updated_at DESC, name COLLATE NOCASE
            """,
            (int(user_id), _saved_view_scope(scope)),
        ).fetchall()
        views = []
        for row in rows:
            item = dict(row)
            try:
                item["filters"] = json.loads(item.get("filters_json") or "{}")
            except (TypeError, json.JSONDecodeError):
                item["filters"] = {}
            views.append(item)
        return views
    finally:
        connection.close()


def save_crm_saved_view(db_path: Path, payload: dict[str, Any], user: dict[str, Any]) -> int:
    user_id = _user_id(user)
    if not user_id:
        raise ValueError("로그인 사용자를 확인하지 못했습니다.")
    name = _clean(payload.get("name"))
    if not name:
        raise ValueError("저장뷰 이름을 입력해주세요.")
    if len(name) > 40:
        raise ValueError("저장뷰 이름은 40자 이하로 입력해주세요.")
    scope = _saved_view_scope(payload.get("scope"))
    filters = _saved_view_filters(payload.get("filters"))
    sort_key = _clean(payload.get("sort_key") or filters.get("sort"))
    is_default = 1 if payload.get("is_default") else 0
    view_id = int(payload.get("id") or 0)
    now = now_text()
    connection = _connect(db_path)
    try:
        if is_default:
            connection.execute(
                "UPDATE crm_saved_views SET is_default = 0, updated_at = ? WHERE user_id = ? AND scope = ?",
                (now, user_id, scope),
            )
        existing = None
        if view_id:
            existing = connection.execute(
                "SELECT id FROM crm_saved_views WHERE id = ? AND user_id = ?",
                (view_id, user_id),
            ).fetchone()
        if not existing:
            existing = connection.execute(
                "SELECT id FROM crm_saved_views WHERE user_id = ? AND scope = ? AND name = ?",
                (user_id, scope, name),
            ).fetchone()
        if existing:
            view_id = int(existing["id"])
            connection.execute(
                """
                UPDATE crm_saved_views
                   SET name = ?, scope = ?, filters_json = ?, sort_key = ?, is_default = ?, updated_at = ?
                 WHERE id = ? AND user_id = ?
                """,
                (
                    name,
                    scope,
                    json.dumps(filters, ensure_ascii=False),
                    sort_key,
                    is_default,
                    now,
                    view_id,
                    user_id,
                ),
            )
        else:
            cursor = connection.execute(
                """
                INSERT INTO crm_saved_views
                    (user_id, name, scope, filters_json, sort_key, is_default, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    user_id,
                    name,
                    scope,
                    json.dumps(filters, ensure_ascii=False),
                    sort_key,
                    is_default,
                    now,
                    now,
                ),
            )
            view_id = int(cursor.lastrowid)
        connection.commit()
        return view_id
    finally:
        connection.close()


def delete_crm_saved_view(db_path: Path, view_id: int, user: dict[str, Any]) -> None:
    user_id = _user_id(user)
    if not user_id:
        raise ValueError("로그인 사용자를 확인하지 못했습니다.")
    connection = _connect(db_path)
    try:
        cursor = connection.execute(
            "DELETE FROM crm_saved_views WHERE id = ? AND user_id = ?",
            (int(view_id), user_id),
        )
        if cursor.rowcount == 0:
            raise ValueError("삭제할 저장뷰를 찾지 못했습니다.")
        connection.commit()
    finally:
        connection.close()


def save_crm_task(db_path: Path, payload: dict[str, Any], user: dict[str, Any]) -> int:
    task_id = int(payload.get("id") or 0)
    title = _clean(payload.get("title"))
    if not title:
        raise ValueError("업무 제목을 입력해주세요.")
    connection = _connect(db_path)
    try:
        now = now_text()
        account_id = int(payload.get("account_id") or 0)
        account_name = _clean(payload.get("account_name"))
        if account_id:
            account = connection.execute("SELECT * FROM crm_accounts WHERE id = ?", (account_id,)).fetchone()
            if not account:
                raise ValueError("거래처를 찾지 못했습니다.")
        else:
            account = _ensure_account(connection, account_name) if account_name else None
        if account:
            account_id = int(account["id"])
            account_name = account["name"]

        assignee = _find_user(connection, payload.get("assignee_user_id") or payload.get("assignee_name"))
        assignee_user_id = int(assignee["id"]) if assignee else None
        assignee_name = assignee["display_name"] if assignee else _clean(payload.get("assignee_name"))
        status = _normalize_status(payload.get("status"))
        completed_at = now if status == "완료" else ""

        values = (
            now,
            account_id or None,
            account_name,
            title,
            _clean(payload.get("description")),
            assignee_user_id,
            assignee_name,
            _user_id(user),
            _user_name(user),
            _parse_due_text(payload.get("due_at")),
            status,
            _normalize_priority(payload.get("priority")),
            _clean(payload.get("source")) or "app",
            completed_at,
        )
        if task_id:
            connection.execute(
                """
                UPDATE crm_tasks
                   SET updated_at = ?, account_id = ?, account_name = ?, title = ?,
                       description = ?, assignee_user_id = ?, assignee_name = ?,
                       requester_user_id = ?, requester_name = ?, due_at = ?, status = ?,
                       priority = ?, source = ?, completed_at = ?
                 WHERE id = ?
                """,
                values + (task_id,),
            )
        else:
            cursor = connection.execute(
                """
                INSERT INTO crm_tasks
                    (created_at, updated_at, account_id, account_name, title, description,
                     assignee_user_id, assignee_name, requester_user_id, requester_name,
                     due_at, status, priority, source, completed_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (now,) + values,
            )
            task_id = int(cursor.lastrowid)
            connection.execute(
                "UPDATE crm_tasks SET public_id = ? WHERE id = ?",
                (_task_public_id(task_id), task_id),
            )
        connection.commit()
        return task_id
    finally:
        connection.close()


def change_crm_task_status(
    db_path: Path,
    task_id: int,
    status: str,
    user: dict[str, Any],
    comment: str = "",
    can_manage: bool = False,
) -> int:
    connection = _connect(db_path)
    try:
        task = connection.execute("SELECT * FROM crm_tasks WHERE id = ?", (task_id,)).fetchone()
        if not task:
            raise ValueError("업무를 찾지 못했습니다.")
        if not user_can_touch_task(task, user, can_manage):
            raise PermissionError("본인 담당 업무만 처리할 수 있습니다.")
        normalized = _normalize_status(status)
        now = now_text()
        completed_at = now if normalized == "완료" else ""
        connection.execute(
            "UPDATE crm_tasks SET status = ?, updated_at = ?, completed_at = ? WHERE id = ?",
            (normalized, now, completed_at, task_id),
        )
        body = _clean(comment) or normalized
        connection.execute(
            """
            INSERT INTO crm_task_comments (task_id, created_at, author_user_id, author_name, comment_type, body)
            VALUES (?, ?, ?, ?, 'status', ?)
            """,
            (task_id, now, _user_id(user), _user_name(user), body),
        )
        connection.commit()
        return task_id
    finally:
        connection.close()


def add_crm_task_comment(
    db_path: Path,
    task_id: int,
    body: str,
    user: dict[str, Any],
    can_manage: bool = False,
    comment_type: str = "comment",
) -> int:
    text = _clean(body)
    if not text:
        raise ValueError("댓글 내용을 입력해주세요.")
    connection = _connect(db_path)
    try:
        task = connection.execute("SELECT * FROM crm_tasks WHERE id = ?", (task_id,)).fetchone()
        if not task:
            raise ValueError("업무를 찾지 못했습니다.")
        if not user_can_touch_task(task, user, can_manage):
            raise PermissionError("본인 담당 업무만 처리할 수 있습니다.")
        now = now_text()
        connection.execute(
            """
            INSERT INTO crm_task_comments (task_id, created_at, author_user_id, author_name, comment_type, body)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (task_id, now, _user_id(user), _user_name(user), comment_type, text),
        )
        connection.execute("UPDATE crm_tasks SET updated_at = ? WHERE id = ?", (now, task_id))
        connection.commit()
        return task_id
    finally:
        connection.close()


def list_crm_message_events(db_path: Path, limit: int = 100) -> list[dict[str, Any]]:
    connection = _connect(db_path)
    try:
        rows = connection.execute(
            """
            SELECT events.*, tasks.public_id
              FROM crm_message_events events
              LEFT JOIN crm_tasks tasks ON tasks.id = events.task_id
             ORDER BY events.id DESC
             LIMIT ?
            """,
            (max(1, min(int(limit), 500)),),
        ).fetchall()
        return _rows(rows)
    finally:
        connection.close()


def list_crm_messenger_users(db_path: Path) -> dict[str, Any]:
    connection = _connect(db_path)
    try:
        mappings = connection.execute(
            """
            SELECT mapped.*, users.display_name AS workhub_display_name
              FROM crm_messenger_users mapped
              JOIN users ON users.id = mapped.user_id
             ORDER BY mapped.platform, mapped.display_name
            """
        ).fetchall()
        users = connection.execute(
            "SELECT id, username, display_name FROM users WHERE active = 1 ORDER BY display_name, username"
        ).fetchall()
        return {"mappings": _rows(mappings), "users": _rows(users)}
    finally:
        connection.close()


def save_crm_messenger_user(db_path: Path, payload: dict[str, Any]) -> int:
    platform = _clean(payload.get("platform")) or "kakao"
    sender_key = _clean(payload.get("sender_key"))
    if not sender_key:
        raise ValueError("메신저 사용자 키를 입력해주세요.")
    user_id = int(payload.get("user_id") or 0)
    if not user_id:
        raise ValueError("연결할 Workhub 사용자를 선택해주세요.")
    connection = _connect(db_path)
    try:
        user = connection.execute(
            "SELECT id, username, display_name FROM users WHERE id = ? AND active = 1",
            (user_id,),
        ).fetchone()
        if not user:
            raise ValueError("Workhub 사용자를 찾지 못했습니다.")
        now = now_text()
        connection.execute(
            """
            INSERT INTO crm_messenger_users
                (created_at, updated_at, platform, sender_key, display_name, user_id, username, active)
            VALUES (?, ?, ?, ?, ?, ?, ?, 1)
            ON CONFLICT(platform, sender_key) DO UPDATE SET
                updated_at = excluded.updated_at,
                display_name = excluded.display_name,
                user_id = excluded.user_id,
                username = excluded.username,
                active = 1
            """,
            (
                now,
                now,
                platform,
                sender_key,
                _clean(payload.get("display_name")) or user["display_name"],
                int(user["id"]),
                user["username"],
            ),
        )
        connection.commit()
        row = connection.execute(
            "SELECT id FROM crm_messenger_users WHERE platform = ? AND sender_key = ?",
            (platform, sender_key),
        ).fetchone()
        return int(row["id"])
    finally:
        connection.close()


def _extract_incoming_message(payload: dict[str, Any]) -> dict[str, str]:
    if isinstance(payload.get("userRequest"), dict):
        request = payload.get("userRequest") or {}
        user = request.get("user") or {}
        properties = user.get("properties") or {}
        sender_key = (
            _clean(user.get("id"))
            or _clean(properties.get("plusfriendUserKey"))
            or _clean(properties.get("botUserKey"))
            or _clean(properties.get("appUserId"))
        )
        return {
            "platform": "kakao",
            "sender_key": sender_key,
            "sender_name": _clean(properties.get("nickname")) or sender_key,
            "text": _clean(request.get("utterance")),
            "reply_type": "kakao",
        }
    return {
        "platform": _clean(payload.get("platform")) or "generic",
        "sender_key": _clean(payload.get("sender_key") or payload.get("sender") or payload.get("user_id")),
        "sender_name": _clean(payload.get("sender_name") or payload.get("name")),
        "text": _clean(payload.get("text") or payload.get("message") or payload.get("utterance")),
        "reply_type": "json",
    }


def _find_mapping(connection: sqlite3.Connection, platform: str, sender_key: str) -> sqlite3.Row | None:
    return connection.execute(
        """
        SELECT mapped.*, users.display_name AS workhub_display_name, users.role, users.permissions
          FROM crm_messenger_users mapped
          JOIN users ON users.id = mapped.user_id
         WHERE mapped.platform = ?
           AND mapped.sender_key = ?
           AND mapped.active = 1
        """,
        (platform, sender_key),
    ).fetchone()


def _log_event(
    connection: sqlite3.Connection,
    incoming: dict[str, str],
    payload: dict[str, Any],
    result: str,
    error: str = "",
    task_id: int | None = None,
) -> int:
    cursor = connection.execute(
        """
        INSERT INTO crm_message_events
            (created_at, platform, sender_key, sender_name, text, payload_json, result, error, task_id)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            now_text(),
            incoming["platform"],
            incoming["sender_key"],
            incoming["sender_name"],
            incoming["text"],
            json.dumps(payload, ensure_ascii=False),
            result,
            error,
            task_id,
        ),
    )
    return int(cursor.lastrowid)


def _create_task_from_command(
    connection: sqlite3.Connection,
    text: str,
    event_id: int,
    requester: sqlite3.Row,
    platform: str = "kakao",
) -> tuple[int, str, list[str]]:
    match = re.match(r"^업무등록\s+@?([^\s/]+)\s+(.+?)\s*/\s*(.+?)\s*/\s*(.+)$", text)
    if not match:
        raise ValueError("업무등록 형식: 업무등록 @담당자 거래처명 / 업무내용 / 기한")
    assignee_text, account_name, title, due_at = [part.strip() for part in match.groups()]
    assignee = _find_user(connection, assignee_text)
    if not assignee:
        raise ValueError(f"담당자 '{assignee_text}'를 찾지 못했습니다.")
    account = _ensure_account(connection, account_name)
    now = now_text()
    cursor = connection.execute(
        """
        INSERT INTO crm_tasks
            (created_at, updated_at, account_id, account_name, title, description,
             assignee_user_id, assignee_name, requester_user_id, requester_name,
             due_at, status, priority, source, source_message_event_id)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '대기', '보통', ?, ?)
        """,
        (
            now,
            now,
            int(account["id"]),
            account["name"],
            title,
            title,
            int(assignee["id"]),
            assignee["display_name"],
            int(requester["user_id"]),
            requester["workhub_display_name"] or requester["display_name"] or requester["username"],
            _parse_due_text(due_at),
            f"messenger:{platform or 'kakao'}",
            event_id,
        ),
    )
    task_id = int(cursor.lastrowid)
    public_id = _task_public_id(task_id)
    connection.execute("UPDATE crm_tasks SET public_id = ? WHERE id = ?", (public_id, task_id))
    connection.execute(
        """
        INSERT INTO crm_task_comments (task_id, created_at, author_user_id, author_name, comment_type, body)
        VALUES (?, ?, ?, ?, 'messenger', ?)
        """,
        (
            task_id,
            now,
            int(requester["user_id"]),
            requester["workhub_display_name"] or requester["display_name"] or requester["username"],
            "메신저에서 업무가 등록되었습니다.",
        ),
    )
    return task_id, f"{public_id} 업무가 등록됐습니다.", _task_quick_replies(public_id)


def _task_quick_replies(public_id: str) -> list[str]:
    return [
        f"확인 {public_id}",
        f"완료 {public_id}",
        f"보류 {public_id} 사유",
        f"댓글 {public_id} 내용",
        "내업무",
    ]


def _help_message() -> tuple[str, list[str]]:
    return (
        "사용 가능한 CRM 명령입니다.\n"
        "업무등록 @담당자 거래처명 / 업무내용 / 기한\n"
        "확인 TASK-0001\n"
        "완료 TASK-0001\n"
        "보류 TASK-0001 사유\n"
        "댓글 TASK-0001 내용\n"
        "내업무",
        ["내업무", "업무등록 @관리자 거래처명 / 업무내용 / 오늘 5시"],
    )


def _my_tasks_message(connection: sqlite3.Connection, requester: sqlite3.Row) -> tuple[str, list[str]]:
    rows = connection.execute(
        """
        SELECT public_id, account_name, title, due_at, status, priority
          FROM crm_tasks
         WHERE status != '완료'
           AND assignee_user_id = ?
         ORDER BY CASE status WHEN '대기' THEN 0 WHEN '진행중' THEN 1 WHEN '보류' THEN 2 ELSE 3 END,
                  CASE WHEN due_at IS NULL OR due_at = '' THEN 1 ELSE 0 END,
                  due_at,
                  updated_at DESC
         LIMIT 5
        """,
        (int(requester["user_id"]),),
    ).fetchall()
    if not rows:
        return "현재 배정된 미완료 CRM 업무가 없습니다.", ["도움말"]
    lines = ["내 미완료 CRM 업무"]
    for row in rows:
        due = row["due_at"] or "기한 없음"
        account_name = row["account_name"] or "거래처 미지정"
        lines.append(f"{row['public_id']} · {row['status']} · {account_name} · {row['title']} · {due}")
    first_public_id = rows[0]["public_id"]
    return "\n".join(lines), _task_quick_replies(first_public_id)


def _update_task_from_command(
    connection: sqlite3.Connection,
    text: str,
    requester: sqlite3.Row,
) -> tuple[int, str, list[str]]:
    match = re.match(r"^(확인|완료|보류|댓글)\s+([A-Za-z0-9-]+)(?:\s+(.+))?$", text)
    if not match:
        raise ValueError("지원 명령: 확인 TASK-0001, 완료 TASK-0001, 보류 TASK-0001 사유, 댓글 TASK-0001 내용")
    command, public_id, body = match.groups()
    task = _task_by_public_id(connection, public_id)
    if not task:
        raise ValueError("업무 번호를 찾지 못했습니다.")
    now = now_text()
    author = requester["workhub_display_name"] or requester["display_name"] or requester["username"]
    if command == "댓글":
        comment = _clean(body)
        if not comment:
            raise ValueError("댓글 내용을 입력해주세요.")
        connection.execute(
            """
            INSERT INTO crm_task_comments (task_id, created_at, author_user_id, author_name, comment_type, body)
            VALUES (?, ?, ?, ?, 'messenger', ?)
            """,
            (int(task["id"]), now, int(requester["user_id"]), author, comment),
        )
        connection.execute("UPDATE crm_tasks SET updated_at = ? WHERE id = ?", (now, int(task["id"])))
        return int(task["id"]), f"{task['public_id']} 댓글이 등록됐습니다.", _task_quick_replies(task["public_id"])

    next_status = "진행중" if command == "확인" else command
    comment = _clean(body) or command
    completed_at = now if next_status == "완료" else ""
    connection.execute(
        "UPDATE crm_tasks SET status = ?, updated_at = ?, completed_at = ? WHERE id = ?",
        (next_status, now, completed_at, int(task["id"])),
    )
    connection.execute(
        """
        INSERT INTO crm_task_comments (task_id, created_at, author_user_id, author_name, comment_type, body)
        VALUES (?, ?, ?, ?, 'messenger', ?)
        """,
        (int(task["id"]), now, int(requester["user_id"]), author, comment),
    )
    return int(task["id"]), f"{task['public_id']} 상태가 {next_status}(으)로 변경됐습니다.", _task_quick_replies(task["public_id"])


def _message_response(reply_type: str, ok: bool, message: str, quick_replies: list[str] | None = None) -> dict[str, Any]:
    quick_replies = [item for item in (quick_replies or []) if _clean(item)][:10]
    if reply_type == "kakao":
        response = {
            "version": "2.0",
            "template": {
                "outputs": [
                    {
                        "simpleText": {
                            "text": message,
                        }
                    }
                ]
            },
        }
        if quick_replies:
            response["template"]["quickReplies"] = [
                {"label": item[:14], "action": "message", "messageText": item}
                for item in quick_replies
            ]
        return response
    return {"ok": ok, "message": message, "quick_replies": quick_replies}


def handle_crm_messenger_webhook(db_path: Path, payload: dict[str, Any]) -> dict[str, Any]:
    incoming = _extract_incoming_message(payload)
    if not incoming["sender_key"]:
        incoming["sender_key"] = "unknown"
    connection = _connect(db_path)
    try:
        mapping = _find_mapping(connection, incoming["platform"], incoming["sender_key"])
        if not mapping:
            _log_event(connection, incoming, payload, "거절", "등록되지 않은 메신저 사용자입니다.")
            connection.commit()
            return _message_response(incoming["reply_type"], False, "등록된 직원만 CRM 명령을 사용할 수 있습니다.")
        text = incoming["text"]
        if not text:
            _log_event(connection, incoming, payload, "오류", "빈 메시지입니다.")
            connection.commit()
            return _message_response(incoming["reply_type"], False, "명령어를 입력해주세요.")

        event_id = _log_event(connection, incoming, payload, "처리중")
        try:
            task_id: int | None = None
            quick_replies: list[str] = []
            if text == "도움말" or text.lower() == "help":
                message, quick_replies = _help_message()
            elif text == "내업무":
                message, quick_replies = _my_tasks_message(connection, mapping)
            elif text.startswith("업무등록"):
                task_id, message, quick_replies = _create_task_from_command(
                    connection,
                    text,
                    event_id,
                    mapping,
                    incoming["platform"],
                )
            else:
                task_id, message, quick_replies = _update_task_from_command(connection, text, mapping)
            connection.execute(
                "UPDATE crm_message_events SET result = '성공', task_id = ? WHERE id = ?",
                (task_id, event_id),
            )
            connection.commit()
            return _message_response(incoming["reply_type"], True, message, quick_replies)
        except Exception as exc:  # noqa: BLE001
            connection.execute(
                "UPDATE crm_message_events SET result = '오류', error = ? WHERE id = ?",
                (str(exc), event_id),
            )
            connection.commit()
            return _message_response(incoming["reply_type"], False, str(exc))
    finally:
        connection.close()
