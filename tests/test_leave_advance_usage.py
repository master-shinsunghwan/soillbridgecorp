from __future__ import annotations

import importlib
import os
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


def load_app(tmp_path: Path):
    os.environ["WORKHUB_DATA_DIR"] = str(tmp_path)
    sys.modules.pop("workhub_delivery_app", None)
    app = importlib.import_module("workhub_delivery_app")
    app.init_db()
    return app


def get_user(app, username: str) -> dict[str, str]:
    connection = app.connect_db()
    try:
        row = connection.execute(
            "SELECT id, username, display_name, role, permissions FROM users WHERE username = ?",
            (username,),
        ).fetchone()
        assert row is not None
        return dict(row)
    finally:
        connection.close()


class LeaveAdvanceUsageTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(self.tempdir.cleanup)
        self.app = load_app(Path(self.tempdir.name))
        self.admin = get_user(self.app, "admin")
        for payload in (
            {
                "username": "director-shin",
                "display_name": "신성환 실장",
                "role": "sub_admin",
                "password": "SafePass!4567",
                "active": True,
                "permissions": ["leave_view", "leave_approve_director"],
            },
            {
                "username": "employee",
                "display_name": "일반 직원",
                "role": "user",
                "password": "SafePass!5678",
                "active": True,
                "permissions": ["leave_view"],
            },
        ):
            self.app.save_user_account(payload, self.admin)
        self.director = get_user(self.app, "director-shin")
        self.employee = get_user(self.app, "employee")
        self.target_year = date.today().year + 1

    def advance_payload(self, *, day: int = 5, unit: str = "FULL_DAY") -> dict[str, object]:
        return {
            "user_id": self.employee["id"],
            "usage_date": f"{self.target_year}-01-{day:02d}",
            "unit": unit,
            "note": "다음 연도 선사용",
        }

    def test_only_shin_director_and_admin_can_manage_advance_usage(self) -> None:
        self.assertTrue(self.app.can_manage_next_year_leave(self.director))
        self.assertTrue(self.app.can_manage_next_year_leave(self.admin))
        self.assertFalse(self.app.can_manage_next_year_leave(self.employee))

        with self.assertRaisesRegex(ValueError, "권한"):
            self.app.add_next_year_leave_usage(self.advance_payload(), self.employee)

    def test_registers_half_day_and_rejects_duplicate_or_wrong_year(self) -> None:
        entry_id = self.app.add_next_year_leave_usage(
            self.advance_payload(unit="HALF_DAY"),
            self.director,
        )

        rows = self.app.list_next_year_leave_usage(self.director)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["id"], entry_id)
        self.assertEqual(rows[0]["days"], 0.5)
        self.assertEqual(rows[0]["status"], "PENDING")
        self.assertEqual(rows[0]["created_by_name"], "신성환 실장")

        with self.assertRaisesRegex(ValueError, "중복"):
            self.app.add_next_year_leave_usage(self.advance_payload(unit="HALF_DAY"), self.director)

        wrong_year = self.advance_payload(day=6)
        wrong_year["usage_date"] = f"{date.today().year}-12-31"
        with self.assertRaisesRegex(ValueError, "다음 연도"):
            self.app.add_next_year_leave_usage(wrong_year, self.director)

    def test_pending_advance_usage_can_be_deleted_but_applied_usage_cannot(self) -> None:
        entry_id = self.app.add_next_year_leave_usage(self.advance_payload(), self.director)

        self.app.delete_next_year_leave_usage(entry_id, self.director)
        self.assertEqual(self.app.list_next_year_leave_usage(self.director), [])

        applied_id = self.app.add_next_year_leave_usage(self.advance_payload(day=7), self.director)
        self.app.apply_annual_leave_accrual(self.target_year, actor=self.admin, default_days=15)
        with self.assertRaisesRegex(ValueError, "차감 완료"):
            self.app.delete_next_year_leave_usage(applied_id, self.director)

    def test_accrual_applies_advance_usage_once_and_creates_approved_history(self) -> None:
        self.app.add_next_year_leave_usage(self.advance_payload(day=8), self.director)
        self.app.add_next_year_leave_usage(self.advance_payload(day=9, unit="HALF_DAY"), self.director)

        first_count = self.app.apply_annual_leave_accrual(self.target_year, actor=self.admin, default_days=15)
        second_count = self.app.apply_annual_leave_accrual(self.target_year, actor=self.admin, default_days=15)

        connection = self.app.connect_db()
        try:
            balance = connection.execute(
                "SELECT total_days, used_days, remaining_days, accrual_year FROM leave_balances WHERE user_id = ?",
                (self.employee["id"],),
            ).fetchone()
            applied_count = connection.execute(
                "SELECT COUNT(*) AS count FROM leave_advance_usage WHERE user_id = ? AND applied_at IS NOT NULL",
                (self.employee["id"],),
            ).fetchone()["count"]
            approved = connection.execute(
                "SELECT requested_days, status, approval_step FROM leave_requests WHERE user_id = ? ORDER BY start_date",
                (self.employee["id"],),
            ).fetchall()
        finally:
            connection.close()

        self.assertGreater(first_count, 0)
        self.assertEqual(second_count, 0)
        self.assertEqual(balance["total_days"], 15)
        self.assertEqual(balance["used_days"], 1.5)
        self.assertEqual(balance["remaining_days"], 13.5)
        self.assertEqual(balance["accrual_year"], self.target_year)
        self.assertEqual(applied_count, 2)
        self.assertEqual([(row["requested_days"], row["status"], row["approval_step"]) for row in approved], [
            (1.0, "APPROVED", "COMPLETED"),
            (0.5, "APPROVED", "COMPLETED"),
        ])

    def test_accrual_rolls_back_when_advance_usage_exceeds_total(self) -> None:
        self.app.add_next_year_leave_usage(self.advance_payload(day=10), self.director)
        self.app.add_next_year_leave_usage(self.advance_payload(day=11), self.director)

        with self.assertRaisesRegex(ValueError, "초과"):
            self.app.apply_annual_leave_accrual(self.target_year, actor=self.admin, default_days=1)

        connection = self.app.connect_db()
        try:
            balance = connection.execute(
                "SELECT accrual_year FROM leave_balances WHERE user_id = ?",
                (self.employee["id"],),
            ).fetchone()
            pending = connection.execute(
                "SELECT COUNT(*) AS count FROM leave_advance_usage WHERE user_id = ? AND applied_at IS NULL",
                (self.employee["id"],),
            ).fetchone()["count"]
        finally:
            connection.close()

        self.assertTrue(balance is None or balance["accrual_year"] is None)
        self.assertEqual(pending, 2)
