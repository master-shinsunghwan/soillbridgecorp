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


def load_app(tmp_path: Path):
    os.environ["WORKHUB_DATA_DIR"] = str(tmp_path)
    sys.modules.pop("workhub_delivery_app", None)
    return importlib.import_module("workhub_delivery_app")


def admin_user() -> dict[str, object]:
    return {
        "id": 1,
        "username": "admin",
        "display_name": "관리자",
        "role": "admin",
        "permissions": ["ledger_edit", "mail_send", "notice_manage", "backup_manage"],
    }


def director_user() -> dict[str, object]:
    return {
        "id": 2,
        "username": "ssh19",
        "display_name": "신성환 실장",
        "role": "user",
        "permissions": ["ledger_edit", "mail_send", "notice_manage", "backup_manage"],
    }


def staff_user() -> dict[str, object]:
    return {
        "id": 3,
        "username": "staff",
        "display_name": "직원",
        "role": "user",
        "permissions": ["ledger_edit", "mail_send", "notice_manage", "backup_manage"],
    }


class WorkhubAutomationCenterTests(unittest.TestCase):
    def insert_management_record(self, app, **overrides) -> int:
        app.init_db()
        payload = {
            "created_at": app.now_text(),
            "source_file": "test.xlsx",
            "source_sheet": "Sheet1",
            "source_row": 1,
            "sales_vendor": "쿠팡 로켓",
            "purchase_vendor": "테스트 매입처",
            "transaction_type": "",
            "order_date": "2026-07-01",
            "ship_date": "2026-07-01",
            "product_name": " 테스트 상품 ",
            "quantity": "",
            "courier": "CJ 대한통운",
            "invoice_number": "123",
        }
        payload.update(overrides)
        connection = app.connect_db()
        try:
            columns = ", ".join(payload.keys())
            placeholders = ", ".join("?" for _ in payload)
            cursor = connection.execute(
                f"INSERT INTO management_records ({columns}) VALUES ({placeholders})",
                list(payload.values()),
            )
            connection.commit()
            return int(cursor.lastrowid)
        finally:
            connection.close()

    def test_automation_center_lists_admin_actions(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))

            payload = app.automation_center_payload(admin_user())
            action_ids = {action["id"] for action in payload["actions"]}

            self.assertIn("management_rules", action_ids)
            self.assertIn("bulk_db_change", action_ids)
            self.assertIn("mail_failure_ops", action_ids)
            self.assertIn("cs_bulk_mail", action_ids)
            self.assertIn("notice_auto", action_ids)
            self.assertIn("backup_ops", action_ids)

    def test_automation_center_visibility_is_limited_to_admin_and_director(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))

            self.assertTrue(app.can_view_automation_center(admin_user()))
            self.assertTrue(app.can_view_automation_center(director_user()))
            self.assertFalse(app.can_view_automation_center(staff_user()))
            with self.assertRaises(PermissionError):
                app.automation_center_payload(staff_user())

    def test_management_rules_preview_and_execute_updates_records(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            record_id = self.insert_management_record(app)

            preview = app.automation_center_preview("management_rules", {}, admin_user())["preview"]
            result = app.automation_center_execute("management_rules", {}, admin_user())["result"]

            self.assertGreaterEqual(preview["row_count"], 3)
            self.assertGreaterEqual(result["updated"], 3)
            row = app.get_management_record(record_id)
            self.assertEqual(row["quantity"], "1")
            self.assertEqual(row["courier"], "CJ대한통운")
            self.assertEqual(row["transaction_type"], "매입/매출")
            self.assertEqual(row["product_name"], "테스트 상품")

    def test_bulk_db_change_preview_and_execute_creates_backup(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            record_id = self.insert_management_record(app)
            user = admin_user()

            preview = app.automation_center_preview(
                "bulk_db_change",
                {"field": "sales_vendor", "find": "쿠팡 로켓", "replace": "쿠팡"},
                user,
            )["preview"]
            executed = app.automation_center_execute(
                "bulk_db_change",
                {"field": "sales_vendor", "find": "쿠팡 로켓", "replace": "쿠팡"},
                user,
            )

            self.assertEqual(preview["row_count"], 1)
            self.assertTrue(executed["backup_name"].startswith("workhub_backup_"))
            self.assertEqual(app.get_management_record(record_id)["sales_vendor"], "쿠팡")

    def test_notice_auto_creates_portal_notice(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            user = admin_user()

            result = app.automation_center_execute("notice_auto", {}, user)["result"]
            notices = app.list_portal_notices()

            self.assertIn("공지사항", result["summary"])
            self.assertEqual(len(notices), 1)
            self.assertEqual(notices[0]["title"], "업무 자동화 점검 공지")


if __name__ == "__main__":
    unittest.main()
