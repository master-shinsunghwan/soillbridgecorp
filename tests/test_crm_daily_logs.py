from __future__ import annotations

import sqlite3
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import workhub_crm as crm  # noqa: E402


def build_db(tmp_path: Path) -> Path:
    db_path = tmp_path / "workhub.db"
    connection = sqlite3.connect(db_path)
    try:
        connection.execute(
            """
            CREATE TABLE users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL,
                display_name TEXT,
                role TEXT,
                active INTEGER NOT NULL DEFAULT 1
            )
            """
        )
        connection.executemany(
            "INSERT INTO users (id, username, display_name, role, active) VALUES (?, ?, ?, ?, 1)",
            [
                (1, "kim", "김대리", "user"),
                (2, "park", "박과장", "sub_admin"),
            ],
        )
        crm.init_crm_db(connection)
        connection.commit()
    finally:
        connection.close()
    return db_path


class CrmDailyLogTests(unittest.TestCase):
    def test_crm_daily_log_upserts_and_lists_by_employee(self) -> None:
        with self.subTest("upsert and list"):
            import tempfile

            with tempfile.TemporaryDirectory() as directory:
                db_path = build_db(Path(directory))
                user = {"id": 1, "username": "kim", "display_name": "김대리"}

                log_id = crm.save_crm_daily_log(
                    db_path,
                    {
                        "log_date": "2026-06-25",
                        "user_id": 1,
                        "work_summary": "발주 확인",
                        "completed_work": "오전 발주 마감",
                    },
                    user,
                )

                self.assertGreater(log_id, 0)
                connection = sqlite3.connect(db_path)
                try:
                    connection.executemany(
                        """
                        INSERT INTO crm_tasks (
                            public_id, created_at, updated_at, title, assignee_user_id, assignee_name,
                            due_at, status, priority, source, completed_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        [
                            (
                                "CRM-20260625-0001",
                                "2026-06-25T08:00:00",
                                "2026-06-25T08:00:00",
                                "오늘 마감 발주 확인",
                                1,
                                "김대리",
                                "2026-06-25",
                                "진행",
                                "높음",
                                "app",
                                "",
                            ),
                            (
                                "CRM-20260625-0002",
                                "2026-06-25T08:30:00",
                                "2026-06-25T09:30:00",
                                "오전 CS 처리",
                                1,
                                "김대리",
                                "",
                                "완료",
                                "보통",
                                "app",
                                "2026-06-25T09:30:00",
                            ),
                        ],
                    )
                    connection.commit()
                finally:
                    connection.close()

                payload = crm.list_crm_daily_logs(db_path, "2026-06-25")
                self.assertEqual(payload["summary"]["submitted"], 1)
                self.assertEqual(payload["summary"]["missing"], 1)
                self.assertEqual([item["display_name"] for item in payload["logs"]], ["박과장", "김대리"])
                kim_log = next(item for item in payload["logs"] if item["user_id"] == 1)
                self.assertTrue(kim_log["submitted"])
                self.assertEqual(kim_log["completed_work"], "오전 발주 마감")
                self.assertEqual(kim_log["open_tasks"], 1)
                self.assertEqual(kim_log["due_today"], 1)
                self.assertEqual(kim_log["completed_today"], 1)
                self.assertEqual(kim_log["due_today_tasks"][0]["title"], "오늘 마감 발주 확인")
                self.assertEqual(kim_log["completed_today_tasks"][0]["title"], "오전 CS 처리")

                updated_id = crm.save_crm_daily_log(
                    db_path,
                    {
                        "log_date": "2026-06-25",
                        "user_id": 1,
                        "work_summary": "발주/CS 확인",
                        "blockers": "특이 이슈 없음",
                        "next_plan": "미처리 CS 확인",
                    },
                    user,
                )
                self.assertEqual(updated_id, log_id)
                updated_payload = crm.list_crm_daily_logs(db_path, "2026-06-25", user_id=1)
                self.assertEqual(updated_payload["logs"][0]["work_summary"], "발주/CS 확인")
                self.assertEqual(updated_payload["summary"]["issues"], 0)

    def test_crm_daily_log_editing_other_employee_requires_manage_permission(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory() as directory:
            db_path = build_db(Path(directory))
            user = {"id": 1, "username": "kim", "display_name": "김대리"}

            with self.assertRaises(PermissionError):
                crm.save_crm_daily_log(
                    db_path,
                    {
                        "log_date": "2026-06-25",
                        "user_id": 2,
                        "work_summary": "대신 작성",
                    },
                    user,
                    can_manage=False,
                )

            log_id = crm.save_crm_daily_log(
                db_path,
                {
                    "log_date": "2026-06-25",
                    "user_id": 2,
                    "work_summary": "관리자 대리 작성",
                },
                user,
                can_manage=True,
            )
            self.assertGreater(log_id, 0)


if __name__ == "__main__":
    unittest.main()
