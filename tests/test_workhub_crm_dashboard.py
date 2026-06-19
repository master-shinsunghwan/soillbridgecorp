from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


import workhub_crm as crm  # noqa: E402


class WorkhubCrmDashboardTests(unittest.TestCase):
    def test_dashboard_payload_keeps_completed_tasks_out_of_active_workflow(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            db_path = Path(directory) / "workhub.db"
            connection = crm._connect(db_path)
            try:
                crm.init_crm_db(connection)
                now = crm.now_text()
                open_task = connection.execute(
                    """
                    INSERT INTO crm_tasks (
                        public_id, created_at, updated_at, account_name, title,
                        assignee_name, due_at, status, priority, source
                    ) VALUES ('TASK-0001', ?, ?, '소일브릿지', '진행 업무',
                              '담당자', ?, '진행중', '높음', 'app')
                    """,
                    (now, now, now),
                ).lastrowid
                completed_task = connection.execute(
                    """
                    INSERT INTO crm_tasks (
                        public_id, created_at, updated_at, account_name, title,
                        assignee_name, due_at, status, priority, source
                    ) VALUES ('TASK-0002', ?, ?, '소일브릿지', '완료 업무',
                              '담당자', ?, '완료', '보통', 'app')
                    """,
                    (now, now, now),
                ).lastrowid
                connection.execute(
                    """
                    INSERT INTO crm_message_events (
                        created_at, platform, sender_key, sender_name, text,
                        payload_json, result, task_id
                    ) VALUES (?, 'kakao', 'open', '담당자', '진행 이벤트',
                              '{}', '성공', ?)
                    """,
                    (now, open_task),
                )
                connection.execute(
                    """
                    INSERT INTO crm_message_events (
                        created_at, platform, sender_key, sender_name, text,
                        payload_json, result, task_id
                    ) VALUES (?, 'kakao', 'done', '담당자', '완료 이벤트',
                              '{}', '성공', ?)
                    """,
                    (now, completed_task),
                )
                connection.commit()
            finally:
                connection.close()

            payload = crm.crm_dashboard_payload(db_path)

        self.assertEqual(payload["stats"]["open_tasks"], 1)
        self.assertNotIn("완료", {row["status"] for row in payload["status_counts"]})
        self.assertEqual([task["title"] for task in payload["project_progress"][0]["tasks"]], ["진행 업무"])
        self.assertEqual([event["text"] for event in payload["recent_events"]], ["진행 이벤트"])


if __name__ == "__main__":
    unittest.main()
