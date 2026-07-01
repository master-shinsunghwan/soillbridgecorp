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


def leave_notification_messages(app, user: dict[str, str]) -> list[str]:
    return [str(item["message"]) for item in app.list_leave_notifications(user)]


class LeaveAdvancedWorkflowTests(unittest.TestCase):
    def make_users(self, app):
        app.init_db()
        admin = get_user(app, "admin")
        users = {
            "requester": {
                "username": "requester",
                "display_name": "신청자",
                "role": "user",
                "password": "SafePass!2345",
                "active": True,
                "permissions": ["leave_view"],
            },
            "team": {
                "username": "teamlead",
                "display_name": "팀장",
                "role": "sub_admin",
                "password": "SafePass!3456",
                "active": True,
                "permissions": ["leave_view", "leave_approve_team"],
            },
            "director": {
                "username": "director",
                "display_name": "실장",
                "role": "sub_admin",
                "password": "SafePass!4567",
                "active": True,
                "permissions": ["leave_view", "leave_approve_director", "leave_director_override"],
            },
            "ceo": {
                "username": "ceo",
                "display_name": "대표",
                "role": "admin",
                "password": "SafePass!5678",
                "active": True,
                "permissions": ["leave_view", "leave_approve_ceo"],
            },
        }
        for payload in users.values():
            app.save_user_account(payload, admin)
        return {key: get_user(app, value["username"]) for key, value in users.items()}

    def set_balance(self, app, user: dict[str, str], total: float = 15) -> None:
        admin = get_user(app, "admin")
        app.set_leave_balance({"user_id": user["id"], "total_days": total, "used_days": 0}, admin)

    def test_public_holiday_is_excluded_and_pending_request_reserves_balance(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            users = self.make_users(app)
            self.set_balance(app, users["requester"], total=5)
            app.save_company_holiday("2026-01-01", "신정")

            request_id = app.create_leave_request(
                users["requester"],
                {
                    "leave_type_id": app.get_leave_type_id("annual"),
                    "unit": "FULL_DAY",
                    "start_date": "2026-01-01",
                    "end_date": "2026-01-02",
                    "reason": "가족 일정",
                },
            )

            connection = app.connect_db()
            try:
                request = connection.execute("SELECT * FROM leave_requests WHERE id = ?", (request_id,)).fetchone()
                balance = connection.execute(
                    "SELECT total_days, used_days, reserved_days, remaining_days FROM leave_balances WHERE user_id = ?",
                    (users["requester"]["id"],),
                ).fetchone()
            finally:
                connection.close()

            self.assertEqual(request["requested_days"], 1.0)
            self.assertEqual(request["approval_step"], "TEAM_LEAD")
            self.assertEqual(request["team_status"], "PENDING")
            self.assertEqual(request["director_status"], "WAITING")
            self.assertEqual(request["ceo_status"], "WAITING")
            self.assertEqual(balance["used_days"], 0)
            self.assertEqual(balance["reserved_days"], 1.0)
            self.assertEqual(balance["remaining_days"], 4.0)

            team_notifications = app.list_leave_notifications(users["team"])
            self.assertTrue(any(item["request_id"] == request_id for item in team_notifications))

    def test_leave_request_and_approval_notifications_reach_requester_and_managers(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            users = self.make_users(app)
            self.set_balance(app, users["requester"], total=5)

            request_id = app.create_leave_request(
                users["requester"],
                {
                    "leave_type_id": app.get_leave_type_id("annual"),
                    "unit": "FULL_DAY",
                    "start_date": "2026-01-05",
                    "end_date": "2026-01-06",
                    "reason": "가족 일정",
                },
            )

            requester_messages = leave_notification_messages(app, users["requester"])
            team_messages = leave_notification_messages(app, users["team"])
            self.assertTrue(any("접수" in message and "팀장 확인" in message for message in requester_messages))
            self.assertTrue(any("신청자님" in message and "팀장 확인" in message for message in team_messages))

            app.decide_leave_request(request_id, users["team"], "approve", "팀장 확인")

            requester_messages = leave_notification_messages(app, users["requester"])
            director_messages = leave_notification_messages(app, users["director"])
            self.assertTrue(any("팀장 확인 승인" in message and "실장 확인" in message for message in requester_messages))
            self.assertTrue(any("신청자님" in message and "실장 확인" in message for message in director_messages))

            app.decide_leave_request(request_id, users["director"], "approve", "실장 확인")

            requester_messages = leave_notification_messages(app, users["requester"])
            ceo_messages = leave_notification_messages(app, users["ceo"])
            self.assertTrue(any("실장 확인 승인" in message and "대표 확인" in message for message in requester_messages))
            self.assertTrue(any("신청자님" in message and "대표 확인" in message for message in ceo_messages))

            app.decide_leave_request(request_id, users["ceo"], "approve", "대표 승인")

            requester_messages = leave_notification_messages(app, users["requester"])
            admin_messages = leave_notification_messages(app, get_user(app, "admin"))
            self.assertTrue(any("최종 승인" in message for message in requester_messages))
            self.assertTrue(any("신청자님" in message and "최종 승인" in message for message in admin_messages))

    def test_leave_cancel_notifies_requester_and_managers(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            users = self.make_users(app)
            self.set_balance(app, users["requester"], total=5)

            request_id = app.create_leave_request(
                users["requester"],
                {
                    "leave_type_id": app.get_leave_type_id("annual"),
                    "unit": "HALF_DAY",
                    "start_date": "2026-01-09",
                    "end_date": "2026-01-09",
                    "reason": "오전 반차",
                },
            )

            app.cancel_leave_request(request_id, users["requester"], "일정 변경")

            requester_messages = leave_notification_messages(app, users["requester"])
            team_messages = leave_notification_messages(app, users["team"])
            self.assertTrue(any("취소" in message and "2026-01-09" in message for message in requester_messages))
            self.assertTrue(any("취소" in message and "신청자님" in message for message in team_messages))

    def test_three_step_approval_finalizes_only_after_ceo_approval(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            users = self.make_users(app)
            self.set_balance(app, users["requester"], total=5)
            request_id = app.create_leave_request(
                users["requester"],
                {
                    "leave_type_id": app.get_leave_type_id("annual"),
                    "unit": "FULL_DAY",
                    "start_date": "2026-01-05",
                    "end_date": "2026-01-07",
                    "reason": "연차",
                },
            )

            app.decide_leave_request(request_id, users["team"], "approve", "팀장 확인")
            app.decide_leave_request(request_id, users["director"], "approve", "실장 확인")

            connection = app.connect_db()
            try:
                mid = connection.execute("SELECT status, approval_step FROM leave_requests WHERE id = ?", (request_id,)).fetchone()
                balance_mid = connection.execute(
                    "SELECT used_days, reserved_days, remaining_days FROM leave_balances WHERE user_id = ?",
                    (users["requester"]["id"],),
                ).fetchone()
            finally:
                connection.close()

            self.assertEqual(mid["status"], "PENDING")
            self.assertEqual(mid["approval_step"], "CEO")
            self.assertEqual(balance_mid["used_days"], 0)
            self.assertEqual(balance_mid["reserved_days"], 3.0)
            self.assertEqual(balance_mid["remaining_days"], 2.0)

            app.decide_leave_request(request_id, users["ceo"], "approve", "대표 승인")

            connection = app.connect_db()
            try:
                done = connection.execute("SELECT status, approval_step FROM leave_requests WHERE id = ?", (request_id,)).fetchone()
                balance_done = connection.execute(
                    "SELECT used_days, reserved_days, remaining_days FROM leave_balances WHERE user_id = ?",
                    (users["requester"]["id"],),
                ).fetchone()
            finally:
                connection.close()

            self.assertEqual(done["status"], "APPROVED")
            self.assertEqual(done["approval_step"], "COMPLETED")
            self.assertEqual(balance_done["used_days"], 3.0)
            self.assertEqual(balance_done["reserved_days"], 0)
            self.assertEqual(balance_done["remaining_days"], 2.0)

    def test_director_can_override_remaining_approval_steps(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            users = self.make_users(app)
            self.set_balance(app, users["requester"], total=5)
            request_id = app.create_leave_request(
                users["requester"],
                {
                    "leave_type_id": app.get_leave_type_id("annual"),
                    "unit": "FULL_DAY",
                    "start_date": "2026-01-05",
                    "end_date": "2026-01-06",
                    "reason": "긴급",
                },
            )

            app.decide_leave_request(request_id, users["director"], "override", "전결 처리")

            connection = app.connect_db()
            try:
                request = connection.execute("SELECT status, approval_step, director_status, ceo_status FROM leave_requests WHERE id = ?", (request_id,)).fetchone()
                balance = connection.execute(
                    "SELECT used_days, reserved_days, remaining_days FROM leave_balances WHERE user_id = ?",
                    (users["requester"]["id"],),
                ).fetchone()
            finally:
                connection.close()

            self.assertEqual(request["status"], "APPROVED")
            self.assertEqual(request["approval_step"], "COMPLETED")
            self.assertEqual(request["director_status"], "OVERRIDDEN")
            self.assertEqual(request["ceo_status"], "OVERRIDDEN")
            self.assertEqual(balance["used_days"], 2.0)
            self.assertEqual(balance["reserved_days"], 0)
            self.assertEqual(balance["remaining_days"], 3.0)

    def test_requester_can_cancel_pending_request_and_release_reservation(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            users = self.make_users(app)
            self.set_balance(app, users["requester"], total=5)
            request_id = app.create_leave_request(
                users["requester"],
                {
                    "leave_type_id": app.get_leave_type_id("annual"),
                    "unit": "FULL_DAY",
                    "start_date": "2026-01-05",
                    "end_date": "2026-01-05",
                    "reason": "취소 예정",
                },
            )

            app.cancel_leave_request(request_id, users["requester"], "일정 변경")

            connection = app.connect_db()
            try:
                request = connection.execute("SELECT status, cancel_reason FROM leave_requests WHERE id = ?", (request_id,)).fetchone()
                balance = connection.execute(
                    "SELECT used_days, reserved_days, remaining_days FROM leave_balances WHERE user_id = ?",
                    (users["requester"]["id"],),
                ).fetchone()
            finally:
                connection.close()

            self.assertEqual(request["status"], "CANCELED")
            self.assertEqual(request["cancel_reason"], "일정 변경")
            self.assertEqual(balance["used_days"], 0)
            self.assertEqual(balance["reserved_days"], 0)
            self.assertEqual(balance["remaining_days"], 5.0)

    def test_auto_accrual_creates_yearly_annual_balance(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            users = self.make_users(app)

            updated = app.apply_annual_leave_accrual(2026, actor=get_user(app, "admin"))

            connection = app.connect_db()
            try:
                balance = connection.execute(
                    "SELECT total_days, used_days, reserved_days, remaining_days FROM leave_balances WHERE user_id = ?",
                    (users["requester"]["id"],),
                ).fetchone()
            finally:
                connection.close()

            self.assertGreaterEqual(updated, 1)
            self.assertEqual(balance["total_days"], 15.0)
            self.assertEqual(balance["used_days"], 0)
            self.assertEqual(balance["reserved_days"], 0)
            self.assertEqual(balance["remaining_days"], 15.0)

    def test_leave_payload_sends_staff_list_to_approvers_and_managers(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            users = self.make_users(app)
            admin = get_user(app, "admin")

            team_payload = app.leave_payload(users["team"])
            admin_payload = app.leave_payload(admin)

            self.assertTrue(team_payload["can_approve"])
            self.assertFalse(team_payload["can_manage"])
            self.assertGreaterEqual(len(team_payload["users"]), 1)
            self.assertTrue(any(row["username"] == "requester" for row in team_payload["users"]))
            self.assertTrue(admin_payload["can_manage"])
            self.assertGreaterEqual(len(admin_payload["users"]), len(team_payload["users"]))

    def test_leave_payload_sends_staff_usage_dates_to_managers(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            users = self.make_users(app)
            admin = get_user(app, "admin")
            self.set_balance(app, users["requester"], total=5)

            app.add_historical_leave_usage(
                {
                    "user_id": users["requester"]["id"],
                    "usage_dates": "2026-01-23\n2026-02-14 반차",
                    "note": "관리자 확인용",
                },
                admin,
            )

            team_payload = app.leave_payload(users["team"])
            admin_payload = app.leave_payload(admin)
            requester_usage = admin_payload["admin_usage_requests"][str(users["requester"]["id"])]

            self.assertEqual(team_payload["admin_usage_requests"], {})
            self.assertEqual(len(requester_usage), 2)
            self.assertEqual(requester_usage[0]["start_date"], "2026-02-14")
            self.assertEqual(requester_usage[0]["requested_days"], 0.5)
            self.assertEqual(requester_usage[1]["start_date"], "2026-01-23")


if __name__ == "__main__":
    unittest.main()
