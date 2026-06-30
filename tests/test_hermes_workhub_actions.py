from __future__ import annotations

import importlib
import os
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


class HermesWorkhubActionTests(unittest.TestCase):
    def load_app(self):
        for module_name in ("workhub_delivery_app", "workhub_crm"):
            sys.modules.pop(module_name, None)
        tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(tempdir.cleanup)
        os.environ["WORKHUB_DATA_DIR"] = tempdir.name
        app = importlib.import_module("workhub_delivery_app")
        app.init_db()
        return app

    def admin_user(self, app):
        return {
            "id": 1,
            "username": "codexadmin",
            "display_name": "Codex Admin",
            "role": "admin",
            "permissions": app.ALL_PERMISSIONS,
        }

    def test_hermes_action_catalog_marks_crm_task_creation_available(self) -> None:
        app = self.load_app()
        catalog = app.hermes_workhub_action_catalog(self.admin_user(app))

        task_action = next(item for item in catalog if item["id"] == "crm.create_task")

        self.assertTrue(task_action["available"])
        self.assertEqual(task_action["required_permission"], "crm_manage")
        self.assertTrue(task_action["mutates"])

    def test_hermes_action_preview_and_execute_create_crm_task(self) -> None:
        app = self.load_app()
        user = self.admin_user(app)

        preview = app.preview_hermes_workhub_action(
            "crm.create_task",
            {
                "title": "헤르메스 로컬 실행 테스트",
                "description": "승인형 실행 API가 업무를 생성하는지 확인",
                "assignee_name": "김테스트",
                "due_at": "2026-07-01",
                "priority": "높음",
                "status": "대기",
            },
            user,
        )
        result = app.execute_hermes_workhub_action("crm.create_task", preview["params"], user)
        tasks = app.list_crm_tasks(app.DB_PATH, query="헤르메스 로컬 실행 테스트", limit=10)
        history = app.list_hermes_history()

        self.assertEqual(preview["preview"]["title"], "업무관리 새 업무 등록")
        self.assertTrue(result["ok"])
        self.assertRegex(result["public_id"], r"^TASK-\d{4}$")
        self.assertEqual(len(tasks), 1)
        self.assertEqual(tasks[0]["title"], "헤르메스 로컬 실행 테스트")
        self.assertEqual(tasks[0]["source"], "hermes")
        self.assertEqual(tasks[0]["priority"], "높음")
        self.assertEqual(tasks[0]["status"], "대기")
        self.assertTrue(any(item.get("action_id") == "crm.create_task" for item in history))

    def test_hermes_action_execute_requires_underlying_permission(self) -> None:
        app = self.load_app()
        user = {
            "id": 2,
            "username": "limited",
            "display_name": "Limited User",
            "role": "user",
            "permissions": ["hermes_use", "hermes_automation"],
        }

        preview = app.preview_hermes_workhub_action("crm.create_task", {"title": "권한 테스트"}, user)

        self.assertFalse(preview["can_execute"])
        with self.assertRaises(PermissionError):
            app.execute_hermes_workhub_action("crm.create_task", preview["params"], user)


if __name__ == "__main__":
    unittest.main()
