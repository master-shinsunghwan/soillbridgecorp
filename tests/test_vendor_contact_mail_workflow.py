from __future__ import annotations

import importlib
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


def load_app(tmp_path: Path):
    os.environ["WORKHUB_DATA_DIR"] = str(tmp_path)
    sys.modules.pop("workhub_delivery_app", None)
    return importlib.import_module("workhub_delivery_app")


def make_vendor_contacts_workbook(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(["거래처구분", "업체명", "메일주소"])
    worksheet.append(["매입처", "탑스미넬", "purchase@example.com"])
    worksheet.append(["매출처", "쿠팡", "sales@example.com"])
    workbook.save(path)


class VendorContactMailWorkflowTests(unittest.TestCase):
    def test_vendor_contacts_import_to_db_with_purchase_sales_type(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            tmp_path = Path(directory)
            app = load_app(tmp_path)
            workbook_path = tmp_path / "vendor_contacts.xlsx"
            make_vendor_contacts_workbook(workbook_path)

            contacts, saved_count = app.import_vendor_contacts_from_workbook(workbook_path)

            self.assertEqual(saved_count, 2)
            self.assertEqual(
                {
                    (contact["vendor_type"], contact["vendor_name"], contact["email"])
                    for contact in contacts
                },
                {
                    ("purchase", "탑스미넬", "purchase@example.com"),
                    ("sales", "쿠팡", "sales@example.com"),
                },
            )
            self.assertEqual(app.find_vendor_contact("탑스미넬", "purchase")["email"], "purchase@example.com")
            self.assertEqual(app.find_vendor_contact("쿠팡", "sales")["email"], "sales@example.com")

    def test_purchase_vendor_cs_mail_prompt_excludes_headquarters(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            app.save_vendor_contact("탑스미넬", "purchase@example.com", vendor_type="purchase")

            self.assertFalse(app.is_purchase_vendor_cs_target("(주)소일브릿지(본사)"))
            self.assertTrue(app.is_purchase_vendor_cs_target("탑스미넬"))

            prompt = app.vendor_cs_mail_prompt(
                {
                    "id": 17,
                    "purchase_vendor": "탑스미넬",
                    "original_info": "2026-06-18 / 1234567890",
                    "original_invoice": "1234567890",
                    "product_name": "테스트 상품",
                    "receiver_name": "홍길동",
                    "receiver_phone": "010-0000-0000",
                    "receiver_address": "서울시 테스트",
                    "cs_type": "교환",
                    "cs_content": "파손",
                }
            )

            self.assertTrue(prompt["enabled"])
            self.assertEqual(prompt["case_id"], 17)
            self.assertEqual(prompt["recipient_email"], "purchase@example.com")
            self.assertEqual(prompt["payload"]["vendor_type"], "purchase")
            self.assertEqual(prompt["payload"]["vendor_name"], "탑스미넬")
            self.assertIn("테스트 상품", prompt["payload"]["body"])

            headquarters_prompt = app.vendor_cs_mail_prompt(
                {"id": 18, "purchase_vendor": "(주)소일브릿지(본사)"}
            )
            self.assertFalse(headquarters_prompt["enabled"])

    def test_management_cs_creation_returns_vendor_mail_prompt_data(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            app.init_db()
            app.save_vendor_contact("탑스미넬", "purchase@example.com", vendor_type="purchase")
            now = app.now_text()
            connection = app.connect_db()
            try:
                cursor = connection.execute(
                    """
                    INSERT INTO management_records (
                        created_at, source_file, source_sheet, source_row, purchase_vendor, sales_vendor,
                        order_date, ship_date, orderer_name, sender_phone, receiver_name, receiver_phone,
                        product_name, quantity, receiver_address, courier, invoice_number, memo
                    ) VALUES (?, 'source.xlsx', 'Sheet1', 2, '탑스미넬', '쿠팡',
                              '2026-06-18', '2026-06-18', '주문자', '010-1111-2222',
                              '홍길동', '010-0000-0000', '테스트 상품', '1',
                              '서울시 테스트', '롯데택배', '1234567890', '파손')
                    """,
                    (now,),
                )
                connection.commit()
                record_id = int(cursor.lastrowid)
            finally:
                connection.close()

            case_id = app.create_cs_case_from_management(record_id)
            case = app.get_cs_case(case_id)
            prompt = app.vendor_cs_mail_prompt(case)

            self.assertTrue(prompt["enabled"])
            self.assertEqual(prompt["case_id"], case_id)
            self.assertEqual(prompt["recipient_email"], "purchase@example.com")
            self.assertEqual(prompt["payload"]["cs_product"], "테스트 상품")
            self.assertEqual(case["occurred_at"], app.date.today().isoformat())
            self.assertEqual(case["order_date"], "2026-06-18")

    def test_deleting_management_created_cs_case_clears_management_received_state(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            app.init_db()
            now = app.now_text()
            connection = app.connect_db()
            try:
                cursor = connection.execute(
                    """
                    INSERT INTO management_records (
                        created_at, source_file, source_sheet, source_row, purchase_vendor, sales_vendor,
                        order_date, ship_date, orderer_name, sender_phone, receiver_name, receiver_phone,
                        product_name, quantity, receiver_address, courier, invoice_number, memo
                    ) VALUES (?, 'source.xlsx', 'Sheet1', 2, '탑스미넬', '쿠팡',
                              '2026-06-18', '2026-06-18', '주문자', '010-1111-2222',
                              '홍길동', '010-0000-0000', '테스트 상품', '1',
                              '서울시 테스트', '롯데택배', '1234567890', '파손')
                    """,
                    (now,),
                )
                connection.commit()
                record_id = int(cursor.lastrowid)
            finally:
                connection.close()

            case_id = app.create_cs_case_from_management(record_id)
            self.assertTrue(app.get_management_record(record_id)["cs_received_at"])
            self.assertTrue(app.list_management_records(limit=1)[0]["cs_received_at"])

            deleted = app.delete_cs_cases([case_id])

            self.assertEqual(deleted, 1)
            self.assertEqual(app.get_management_record(record_id)["cs_received_at"], "")

    def test_orphan_management_cs_received_state_is_reconciled(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            app.init_db()
            now = app.now_text()
            connection = app.connect_db()
            try:
                cursor = connection.execute(
                    """
                    INSERT INTO management_records (
                        created_at, source_file, source_sheet, source_row, purchase_vendor, sales_vendor,
                        order_date, ship_date, orderer_name, sender_phone, receiver_name, receiver_phone,
                        product_name, quantity, receiver_address, courier, invoice_number, memo, cs_received_at
                    ) VALUES (?, 'source.xlsx', 'Sheet1', 2, '탑스미넬', '쿠팡',
                              '2026-06-18', '2026-06-18', '주문자', '010-1111-2222',
                              '홍길동', '010-0000-0000', '테스트 상품', '1',
                              '서울시 테스트', '롯데택배', '1234567890', '파손', ?)
                    """,
                    (now, now),
                )
                connection.commit()
                record_id = int(cursor.lastrowid)
            finally:
                connection.close()

            rows = app.list_management_records(limit=1)

            self.assertEqual(rows[0]["id"], record_id)
            self.assertEqual(rows[0]["cs_received_at"], "")

    def test_manual_cs_case_save_appears_at_top_of_ledger_results(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            app.init_db()
            connection = app.connect_db()
            try:
                for index in range(25):
                    connection.execute(
                        """
                        INSERT INTO cs_cases (
                            created_at, updated_at, status, vendor_name, product_name,
                            receiver_name, cs_content, occurred_at
                        ) VALUES (?, ?, '접수', ?, ?, ?, ?, '2000-01-01')
                        """,
                        (
                            f"2000-01-01 00:00:{index:02d}",
                            f"2000-01-01 00:00:{index:02d}",
                            f"기존거래처{index}",
                            f"기존상품{index}",
                            f"기존수령인{index}",
                            "기존 CS",
                        ),
                    )
                connection.commit()
            finally:
                connection.close()

            case_id = app.save_cs_case(
                {
                    "vendor_type": "purchase",
                    "vendor_name": "탑스미넬",
                    "recipient_email": "topsminell@naver.com",
                    "cs_product": "테스트 상품",
                    "cs_receiver": "홍길동",
                    "cs_content": "파손 접수",
                }
            )

            rows = app.list_cs_cases(limit=1)

            self.assertEqual(rows[0]["id"], case_id)
            self.assertEqual(rows[0]["occurred_at"], app.date.today().isoformat())
            self.assertEqual(rows[0]["purchase_vendor"], "탑스미넬")

    def test_blank_cs_occurrence_sorts_by_order_date_before_import_time(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            app.init_db()
            connection = app.connect_db()
            try:
                connection.execute(
                    """
                    INSERT INTO cs_cases (
                        created_at, updated_at, status, product_name,
                        receiver_name, occurred_at, order_date, ship_date
                    ) VALUES (
                        '2026-06-23 06:38:34', '2026-06-23 06:38:34', '접수',
                        '과거 업로드 데이터', '홍길동', '', '2022-06-27', '2022-06-27'
                    )
                    """
                )
                current = connection.execute(
                    """
                    INSERT INTO cs_cases (
                        created_at, updated_at, status, product_name,
                        receiver_name, occurred_at, order_date, ship_date
                    ) VALUES (
                        '2026-06-20 09:00:00', '2026-06-20 09:00:00', '접수',
                        '현재 데이터', '김영웅', '2026-06-20', '2026-06-20', '2026-06-20'
                    )
                    """
                )
                connection.commit()
                current_id = int(current.lastrowid)
            finally:
                connection.close()

            rows = app.list_cs_cases(limit=1)

            self.assertEqual(rows[0]["id"], current_id)
            self.assertEqual(rows[0]["product_name"], "현재 데이터")

    def test_cs_case_update_can_edit_order_and_shipping_fields(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            case_id = app.save_cs_case(
                {
                    "vendor_type": "purchase",
                    "vendor_name": "기존매입처",
                    "sales_vendor": "기존매출처",
                    "cs_product": "기존 상품",
                    "cs_receiver": "기존수령자",
                    "cs_phone": "010-0000-0000",
                    "cs_address": "기존 주소",
                    "cs_content": "기존 CS",
                    "original_invoice": "1111111111",
                    "quantity": "1",
                    "courier": "기존택배",
                    "order_date": "2026-06-20",
                    "ship_date": "2026-06-21",
                }
            )

            app.update_cs_case(
                case_id,
                {
                    "sales_vendor": "수정매출처",
                    "purchase_vendor": "수정매입처",
                    "order_date": "2026-06-22",
                    "ship_date": "2026-06-23",
                    "orderer_name": "수정주문자",
                    "orderer_phone": "010-1111-2222",
                    "receiver_name": "수정수령자",
                    "receiver_phone": "010-3333-4444",
                    "product_name": "수정 상품",
                    "quantity": "3",
                    "receiver_address": "수정 주소",
                    "courier": "수정택배",
                    "original_invoice": "9999999999",
                    "completed_at": "2026-06-24",
                },
            )

            updated = app.get_cs_case(case_id)

            self.assertEqual(updated["sales_vendor"], "수정매출처")
            self.assertEqual(updated["purchase_vendor"], "수정매입처")
            self.assertEqual(updated["order_date"], "2026-06-22")
            self.assertEqual(updated["ship_date"], "2026-06-23")
            self.assertEqual(updated["orderer_name"], "수정주문자")
            self.assertEqual(updated["orderer_phone"], "010-1111-2222")
            self.assertEqual(updated["receiver_name"], "수정수령자")
            self.assertEqual(updated["receiver_phone"], "010-3333-4444")
            self.assertEqual(updated["product_name"], "수정 상품")
            self.assertEqual(updated["quantity"], "3")
            self.assertEqual(updated["receiver_address"], "수정 주소")
            self.assertEqual(updated["courier"], "수정택배")
            self.assertEqual(updated["original_invoice"], "9999999999")
            self.assertEqual(updated["completed_at"], "2026-06-24")

    def test_mail_settings_store_bulk_mail_technical_defaults(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))

            app.save_mail_settings(
                "soilbridge@naver.com",
                "application-password",
                bulk_settings={
                    "smtp_port": "587",
                    "smtp_security": "tls",
                    "bulk_batch_size": "35",
                    "bulk_send_interval_seconds": "20",
                    "bulk_batch_pause_minutes": "7",
                    "bulk_test_recipient": "test@example.com",
                },
            )

            public_settings = app.load_mail_settings(include_password=False)
            self.assertEqual(public_settings["naver_email"], "soilbridge@naver.com")
            self.assertTrue(public_settings["has_password"])
            self.assertEqual(public_settings["smtp_host"], "smtp.naver.com")
            self.assertEqual(public_settings["smtp_port"], 587)
            self.assertEqual(public_settings["smtp_security"], "tls")
            self.assertEqual(public_settings["bulk_batch_size"], 35)
            self.assertEqual(public_settings["bulk_send_interval_seconds"], 20)
            self.assertEqual(public_settings["bulk_batch_pause_minutes"], 7)
            self.assertEqual(public_settings["bulk_test_recipient"], "test@example.com")
            self.assertNotIn("naver_password", public_settings)

            private_settings = app.load_mail_settings(include_password=True)
            self.assertEqual(private_settings["naver_password"], "application-password")

    def test_naver_test_mail_uses_saved_account_and_clear_korean_content(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            app.save_mail_settings(
                "soilbridge",
                "application-password",
                bulk_settings={
                    "smtp_port": "587",
                    "smtp_security": "tls",
                    "bulk_test_recipient": "recipient@example.com",
                },
            )

            with patch.object(app, "send_naver_mail") as send_mail:
                app.send_mail_test({})

            send_mail.assert_called_once()
            args = send_mail.call_args.args
            kwargs = send_mail.call_args.kwargs
            self.assertEqual(args[:3], ("soilbridge@naver.com", "application-password", "recipient@example.com"))
            self.assertEqual(args[3], "[소일브릿지] 네이버 메일 SMTP 테스트")
            self.assertIn("소일브릿지 업무자동화 프로그램", args[4])
            self.assertEqual(kwargs["smtp_port"], 587)
            self.assertEqual(kwargs["smtp_security"], "tls")

    def test_naver_mail_message_uses_soillbridge_sender_name(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))

            sent_messages = []

            class FakeSmtp:
                def __init__(self, *args, **kwargs):
                    pass

                def __enter__(self):
                    return self

                def __exit__(self, exc_type, exc, tb):
                    return False

                def starttls(self, context=None):
                    return None

                def login(self, email, password):
                    self.login_args = (email, password)

                def send_message(self, message):
                    sent_messages.append(message)

            with patch.object(app.smtplib, "SMTP", FakeSmtp):
                app.send_naver_mail(
                    "soilbridge@naver.com",
                    "application-password",
                    "recipient@example.com",
                    "제목",
                    "본문",
                    smtp_port=587,
                    smtp_security="tls",
                )

            self.assertEqual(len(sent_messages), 1)
            self.assertIn("(주)소일브릿지", sent_messages[0]["From"])
            self.assertEqual(sent_messages[0]["To"], "recipient@example.com")

    def test_naver_mail_message_includes_cs_attachments(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))

            sent_messages = []

            class FakeSmtp:
                def __init__(self, *args, **kwargs):
                    pass

                def __enter__(self):
                    return self

                def __exit__(self, exc_type, exc, tb):
                    return False

                def login(self, email, password):
                    self.login_args = (email, password)

                def send_message(self, message):
                    sent_messages.append(message)

            with patch.object(app.smtplib, "SMTP_SSL", FakeSmtp):
                app.send_naver_mail(
                    "soilbridge@naver.com",
                    "application-password",
                    "recipient@example.com",
                    "CS 요청",
                    "첨부 확인 부탁드립니다.",
                    attachments=[
                        {
                            "filename": "damage.jpg",
                            "data": b"fake-image",
                            "content_type": "image/jpeg",
                        }
                    ],
                )

            self.assertEqual(len(sent_messages), 1)
            attachments = [
                part
                for part in sent_messages[0].iter_attachments()
                if part.get_filename() == "damage.jpg"
            ]
            self.assertEqual(len(attachments), 1)
            self.assertEqual(attachments[0].get_content_type(), "image/jpeg")

    def test_stock_notice_mail_uses_single_bcc_send(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))
            app.save_mail_settings("soilbridge@naver.com", "application-password")

            with patch.object(app, "send_naver_mail") as send_mail:
                app.send_general_mail({
                    "recipient_email": "",
                    "bcc_emails": ["purchase@example.com", "sales@example.com", "purchase@example.com"],
                    "subject": "입고 및 품절 공지",
                    "body": "공지 내용",
                })

            send_mail.assert_called_once()
            args = send_mail.call_args.args
            kwargs = send_mail.call_args.kwargs
            self.assertEqual(args[:3], ("soilbridge@naver.com", "application-password", "soilbridge@naver.com"))
            self.assertEqual(kwargs["bcc_recipients"], ["purchase@example.com", "sales@example.com"])

    def test_naver_mail_message_sets_bcc_header(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            app = load_app(Path(directory))

            sent_messages = []

            class FakeSmtp:
                def __init__(self, *args, **kwargs):
                    pass

                def __enter__(self):
                    return self

                def __exit__(self, exc_type, exc, tb):
                    return False

                def login(self, email, password):
                    self.login_args = (email, password)

                def send_message(self, message):
                    sent_messages.append(message)

            with patch.object(app.smtplib, "SMTP_SSL", FakeSmtp):
                app.send_naver_mail(
                    "soilbridge@naver.com",
                    "application-password",
                    "soilbridge@naver.com",
                    "입고 및 품절 공지",
                    "공지 내용",
                    bcc_recipients=["purchase@example.com", "sales@example.com"],
                )

            self.assertEqual(len(sent_messages), 1)
            self.assertEqual(sent_messages[0]["To"], "soilbridge@naver.com")
            self.assertIn("purchase@example.com", sent_messages[0]["Bcc"])
            self.assertIn("sales@example.com", sent_messages[0]["Bcc"])


if __name__ == "__main__":
    unittest.main()
