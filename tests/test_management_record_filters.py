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


class ManagementRecordFilterTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(self.tempdir.cleanup)
        os.environ["WORKHUB_DATA_DIR"] = self.tempdir.name
        sys.modules.pop("workhub_delivery_app", None)
        self.app = importlib.import_module("workhub_delivery_app")
        self.app.init_db()
        self.insert_record("2026-06-29", "화면상단 매출처", "김상단", "상품A")
        self.target_id = self.insert_record("2026-06-01", "월초 숨은 매출처", "이월초", "상품B")
        self.insert_record("2026-07-01", "다음달 매출처", "박다음", "상품C")
        self.insert_record("2026-05-31", "출고만6월 매출처", "최혼입", "상품D", ship_date="2026-06-01")

    def insert_record(self, order_date: str, sales_vendor: str, receiver_name: str, product_name: str, ship_date: str = "") -> int:
        connection = self.app.connect_db()
        try:
            cursor = connection.execute(
                """
                INSERT INTO management_records (
                    created_at, source_file, source_sheet, source_row,
                    purchase_vendor, sales_vendor, transaction_type, ledger_checked,
                    order_date, ship_date, orderer_name, sender_phone, receiver_name,
                    receiver_phone, product_name, quantity, receiver_address, courier,
                    invoice_number, memo
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    self.app.now_text(),
                    f"{sales_vendor}.xlsx",
                    "Sheet1",
                    "1",
                    "매입처",
                    sales_vendor,
                    "판매",
                    "",
                    order_date,
                    ship_date or order_date,
                    receiver_name,
                    "010-0000-0000",
                    receiver_name,
                    "010-1111-1111",
                    product_name,
                    "1",
                    "서울",
                    "택배",
                    f"INV-{order_date}",
                    "",
                ),
            )
            connection.commit()
            return int(cursor.lastrowid)
        finally:
            connection.close()

    def test_filter_options_use_entire_selected_month(self) -> None:
        visible_rows = self.app.list_management_records(limit=1, year="2026", month="06")
        self.assertEqual(len(visible_rows), 1)
        self.assertNotEqual(visible_rows[0]["id"], self.target_id)

        options = self.app.list_management_filter_options("sales_vendor", year="2026", month="06")
        option_by_value = {option["value"]: option for option in options}

        self.assertIn("월초 숨은 매출처", option_by_value)
        self.assertEqual(option_by_value["월초 숨은 매출처"]["record_id"], self.target_id)
        self.assertNotIn("다음달 매출처", option_by_value)
        self.assertNotIn("출고만6월 매출처", option_by_value)

    def test_management_records_apply_column_filters_on_server(self) -> None:
        rows = self.app.list_management_records(
            limit=50000,
            year="2026",
            month="06",
            filters={"sales_vendor": "숨은"},
        )

        self.assertEqual([row["id"] for row in rows], [self.target_id])


if __name__ == "__main__":
    unittest.main()
