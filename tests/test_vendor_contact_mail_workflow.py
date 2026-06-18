from __future__ import annotations

import importlib
import os
import sys
import tempfile
import unittest
from pathlib import Path

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
            prompt = app.vendor_cs_mail_prompt(app.get_cs_case(case_id))

            self.assertTrue(prompt["enabled"])
            self.assertEqual(prompt["case_id"], case_id)
            self.assertEqual(prompt["recipient_email"], "purchase@example.com")
            self.assertEqual(prompt["payload"]["cs_product"], "테스트 상품")


if __name__ == "__main__":
    unittest.main()
