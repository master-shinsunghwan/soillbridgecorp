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


def user_by_name(app, username: str) -> dict[str, object]:
    connection = app.connect_db()
    try:
        row = connection.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        assert row is not None
        return dict(row)
    finally:
        connection.close()


class UserAdminAccountTests(unittest.TestCase):
    def test_admin_can_delete_existing_user_and_clear_sessions(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            app.init_db()
            admin = user_by_name(app, "admin")
            user_id = app.save_user_account(
                {
                    "username": "worker",
                    "display_name": "Worker",
                    "role": "user",
                    "password": "SafePass!2345",
                    "active": True,
                    "permissions": ["ledger_edit"],
                },
                admin,
            )
            token = app.create_login_session("worker")

            deleted_id = app.delete_user_account(user_id, admin)

            self.assertEqual(deleted_id, user_id)
            self.assertTrue(all(user["username"] != "worker" for user in app.list_users()))
            connection = app.connect_db()
            try:
                session_count = connection.execute(
                    "SELECT COUNT(*) FROM login_sessions WHERE token_hash = ?",
                    (app.token_digest(token),),
                ).fetchone()[0]
                deleted = connection.execute(
                    "SELECT username, display_name, role FROM deleted_user_accounts WHERE original_user_id = ?",
                    (user_id,),
                ).fetchone()
                deleted_columns = {
                    row["name"]
                    for row in connection.execute("PRAGMA table_info(deleted_user_accounts)").fetchall()
                }
            finally:
                connection.close()
            self.assertEqual(session_count, 0)
            self.assertIsNotNone(deleted)
            self.assertEqual(deleted["username"], "worker")
            self.assertEqual(deleted["display_name"], "Worker")
            self.assertEqual(deleted["role"], "user")
            self.assertNotIn("password_hash", deleted_columns)
            self.assertTrue(any(item["username"] == "worker" for item in app.list_deleted_user_accounts()))

    def test_admin_cannot_delete_current_login_account(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            app.init_db()
            admin = user_by_name(app, "admin")

            with self.assertRaises(ValueError):
                app.delete_user_account(admin["id"], admin)

            self.assertTrue(any(user["username"] == "admin" for user in app.list_users()))


if __name__ == "__main__":
    unittest.main()
