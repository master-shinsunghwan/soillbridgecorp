from __future__ import annotations

import importlib
import os
import sys
import tempfile
import unittest
import zipfile
from datetime import date, timedelta
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


class SalesReportUploadTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(self.tempdir.cleanup)
        os.environ["WORKHUB_DATA_DIR"] = self.tempdir.name
        sys.modules.pop("workhub_delivery_app", None)
        self.app = importlib.import_module("workhub_delivery_app")
        self.app.init_db()

    def test_sales_report_upload_helper_accepts_csv_and_records_history(self) -> None:
        saved_path = self.app.save_uploaded_sales_report_file(
            {"file": ("../sales_report.csv", "date,total\n2026-06-19,1000\n".encode("utf-8"))},
            "file",
        )

        saved = self.app.save_sales_report_file(saved_path, self.app.original_uploaded_filename(saved_path.name), "admin")
        files = self.app.list_sales_report_uploads()

        self.assertEqual(len(files), 1)
        self.assertEqual(files[0]["id"], saved["id"])
        self.assertEqual(files[0]["original_name"], "sales_report.csv")
        self.assertEqual(files[0]["uploaded_by"], "admin")
        self.assertGreater(files[0]["size"], 0)

    def test_sales_report_upload_helper_rejects_unknown_file_type(self) -> None:
        with self.assertRaises(ValueError):
            self.app.save_uploaded_sales_report_file({"file": ("sales_report.txt", b"plain")}, "file")

    def test_hermes_sales_report_context_uses_today_only_for_today_request(self) -> None:
        payload = {
            "period": "2026-06",
            "selected_date": "2026-06-29",
            "today_data_uploaded": True,
            "today": {"quantity": 576, "profit_sales_amount": 7288170},
            "month": {"quantity": 9999, "profit_sales_amount": 296117470},
            "seller_total": {"profit_sales_amount": 296117470},
            "supplier_purchase_total": {"purchase_total": 60937748},
            "seller_top": [{"name": "월간 판매사"}],
            "product_top": [{"name": "월간 상품"}],
        }

        context = self.app.hermes_sales_report_context(
            payload,
            self.app.hermes_sales_report_scope("오늘 매출 요약해줘"),
        )

        self.assertEqual(context["scope"], "today")
        self.assertTrue(context["today_only"])
        self.assertEqual(context["today"]["profit_sales_amount"], 7288170)
        self.assertNotIn("month", context)
        self.assertNotIn("seller_total", context)
        self.assertNotIn("top_sellers", context)

    def test_hermes_sales_report_context_keeps_month_for_month_request(self) -> None:
        payload = {
            "period": "2026-06",
            "selected_date": "2026-06-29",
            "today": {"profit_sales_amount": 7288170},
            "month": {"profit_sales_amount": 296117470},
            "seller_total": {"profit_sales_amount": 296117470},
            "supplier_purchase_total": {"purchase_total": 60937748},
            "seller_top": [{"name": "월간 판매사"}],
            "product_top": [{"name": "월간 상품"}],
        }

        context = self.app.hermes_sales_report_context(
            payload,
            self.app.hermes_sales_report_scope("이번 달 누계 매출 알려줘"),
        )

        self.assertEqual(context["scope"], "month")
        self.assertFalse(context["today_only"])
        self.assertEqual(context["month"]["profit_sales_amount"], 296117470)
        self.assertIn("seller_total", context)

    def write_daily_report(self, path: Path) -> None:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Statistics_Sales_2026-06-19"
        sheet.append([
            "일자",
            "소비자가",
            "판매-수량",
            "판매-금액",
            "판매-공급금액",
            "판매-판매배송비",
            "판매-공급배송비",
            "판매-판매합계",
            "판매-공급합계",
            "판매-마진",
            "CS-출고전취소",
            "CS-출고전취소금액(판매)",
            "CS-출고전취소금액(공급)",
            "CS-반품수량(판매)",
            "CS-반품수량(공급)",
            "CS-취소금액(판매)",
            "CS-취소금액(공급)",
            "CS-반품배송비(판매)",
            "CS-반품배송비(공급)",
            "CS-교환배송비(판매)",
            "CS-교환배송비(공급)",
            "CS-추가구매(판매)",
            "CS-추가구매_배송비(판매)",
            "CS-추가구매(공급)",
            "CS-추가구매_배송비(공급)",
            "CS-교환금액(판매)",
            "CS-교환금액_배송비(판매)",
            "CS-교환금액(공급)",
            "CS-교환금액_배송비(공급)",
            "CS-마진",
            "손익-수량 판매사기준",
            "손익-수량 공급사기준",
            "손익-판매금액",
            "손익-공급금액",
            "손익-판매마진",
            "손익-판매배송비",
            "손익-공급배송비",
            "손익-배송비",
            "손익-마진",
            "손익-마진율",
            "거래명세서-매출",
            "거래명세서-매입",
        ])
        sheet.append(["2026-06-19 (금)", 0, 10, 1200, 500, 100, 50, 1300, 550, 750, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 10, 1200, 500, 700, 100, 50, 50, 750, 0.625, 0, 0])
        sheet.append(["2026-06-18 (목)", 0, 8, 1000, 300, 80, 30, 1080, 330, 750, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 8, 8, 1000, 300, 700, 80, 30, 50, 750, 0.6944, 0, 0])
        workbook.save(path)

    def write_seller_report(self, path: Path) -> None:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Statistics_Sales_Seller_2026-06"
        sheet.append([
            "판매사",
            "소비자가",
            "판매-수량",
            "판매-금액",
            "판매-공급금액",
            "판매-판매배송비",
            "판매-공급배송비",
            "판매-판매합계",
            "판매-공급합계",
            "판매-마진",
            "CS-출고전취소",
            "CS-출고전취소금액(판매)",
            "CS-출고전취소금액(공급)",
            "CS-판매사 반품",
            "CS-공급사 반품",
            "CS-금액",
            "CS-공급금액",
            "CS-반품배송비(판매)",
            "CS-반품배송비(공급)",
            "CS-교환배송비(판매)",
            "CS-교환배송비(공급)",
            "CS-추가금액(판매)",
            "CS-추가금액 배송비(판매)",
            "CS-추가금액(공급)",
            "CS-추가금액 배송비(공급)",
            "CS-교환금액(판매)",
            "CS-교환금액 배송비(판매)",
            "CS-교환금액(공급)",
            "CS-교환금액 배송비(공급)",
            "CS-마진",
            "손익-수량 판매사기준",
            "손익-수량 공급사기준",
            "손익-판매금액",
            "손익-공급금액",
            "손익-판매마진",
            "손익-판매배송비",
            "손익-공급배송비",
            "손익-배송비",
            "손익-마진",
            "손익-마진율",
        ])
        sheet.append(["A판매사", 0, 7, 900, 100, 50, 10, 950, 110, 840, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 7, 7, 900, 100, 800, 50, 10, 40, 840, 0.8842])
        sheet.append(["B판매사", 0, 3, 290, 50, 20, 5, 310, 55, 255, 0, 0, 0, 0, 0, -10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -10, 3, 3, 290, 50, 240, 20, 5, 15, 245, 0.7903])
        workbook.save(path)

    def write_product_report(self, path: Path) -> None:
        html = """<!DOCTYPE html>
<html lang="ko"><head><meta http-equiv="Content-Type" content="text/html; charset=euc-kr" /></head>
<body><table border="1">
<tr><td>상품코드</td><td>상품명</td><td>상품메모</td><td>관리메모</td><td>등록일</td><td>소비자가</td><td>판매-수량</td><td>판매-금액</td><td>판매-공급금액</td><td>판매-마진</td><td>CS-출고전취소</td><td>CS-판매사 출고전취소금액</td><td>CS-공급사 출고전취소금액</td><td>CS-판매사 반품</td><td>CS-공급사 반품</td><td>CS-금액</td><td>CS-공급금액</td><td>CS-판매사 추가금액</td><td>CS-공급사 추가금액</td><td>CS-판매사 교환금액</td><td>CS-공급사 교환금액</td><td>CS-마진</td><td>손익-수량 판매사기준</td><td>손익-수량 공급사기준</td><td>손익-판매금액</td><td>손익-공급금액</td><td>손익-마진</td><td>손익-마진율</td><td>거래명세서-매출</td><td>거래명세서-매입</td><td>관리상품코드</td><td>카테고리</td></tr>
<tr><td>P001</td><td>테스트 상품 A</td><td></td><td></td><td>2026-06-01</td><td>0</td><td>5</td><td>700</td><td>100</td><td>600</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>5</td><td>5</td><td>700</td><td>100</td><td>600</td><td>85.71%</td><td>0</td><td>0</td><td></td><td></td></tr>
<tr><td>P002</td><td>테스트 상품 B</td><td></td><td></td><td>2026-06-01</td><td>0</td><td>2</td><td>500</td><td>400</td><td>100</td><td>0</td><td>0</td><td>0</td><td>-1</td><td>-1</td><td>-20</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>-20</td><td>1</td><td>1</td><td>480</td><td>400</td><td>80</td><td>16.67%</td><td>0</td><td>0</td><td></td><td></td></tr>
</table></body></html>"""
        path.write_bytes(html.encode("euc-kr"))

    def write_supplier_report(self, path: Path) -> None:
        headers = [
            "공급사",
            "소비자가",
            "판매-수량",
            "판매-금액",
            "판매-공급금액",
            "판매-판매배송비",
            "판매-공급배송비",
            "판매-판매합계",
            "판매-공급합계",
            "판매-마진",
            "CS-출고전취소",
            "CS-출고전취소금액(판매)",
            "CS-출고전취소금액(공급)",
            "CS-판매사 반품",
            "CS-공급사 반품",
            "CS-금액",
            "CS-공급금액",
            "CS-반품배송비(판매)",
            "CS-반품배송비(공급)",
            "CS-교환배송비(판매)",
            "CS-교환배송비(공급)",
            "CS-추가금액(판매)",
            "CS-추가금액 배송비(판매)",
            "CS-추가금액(공급)",
            "CS-추가금액 배송비(공급)",
            "CS-교환금액(판매)",
            "CS-교환금액 배송비(판매)",
            "CS-교환금액(공급)",
            "CS-교환금액 배송비(공급)",
            "CS-마진",
            "손익-수량 판매사기준",
            "손익-수량 공급사기준",
            "손익-판매금액",
            "손익-공급금액",
            "손익-판매마진",
            "손익-판매배송비",
            "손익-공급배송비",
            "손익-배송비",
            "손익-마진",
            "손익-마진율",
        ]
        rows = [
            ["공급사A", 0, 5, 1000, 700, 100, 50, 1100, 750, 350, 0, 0, 0, 0, 0, -80, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -80, 5, 5, 920, 700, 220, 100, 50, 50, 270, "29.3%"],
            ["공급사B", 0, 2, 500, 900, 0, 0, 500, 900, -400, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 2, 500, 900, -400, 0, 0, 0, -400, "-80%"],
        ]
        table_rows = ["<tr>" + "".join(f"<td>{cell}</td>" for cell in headers) + "</tr>"]
        table_rows.extend("<tr>" + "".join(f"<td>{cell}</td>" for cell in row) + "</tr>" for row in rows)
        html = """<!DOCTYPE html>
<html lang="ko"><head><meta http-equiv="Content-Type" content="text/html; charset=euc-kr" /></head>
<body><table border="1">""" + "".join(table_rows) + "</table></body></html>"
        path.write_bytes(html.encode("euc-kr"))

    def test_sales_report_parser_detects_three_supported_report_types(self) -> None:
        base = Path(self.tempdir.name)
        daily = base / "매출 통계.xlsx"
        seller = base / "매출처별.xlsx"
        product = base / "Statistics_Good_2026-06-19.xls"
        self.write_daily_report(daily)
        self.write_seller_report(seller)
        self.write_product_report(product)

        self.assertEqual(self.app.detect_sales_report_type(daily, daily.name), "daily")
        self.assertEqual(self.app.detect_sales_report_type(seller, seller.name), "seller")
        self.assertEqual(self.app.detect_sales_report_type(product, product.name), "product")

        daily_parsed = self.app.parse_sales_report_file(daily, daily.name)
        seller_parsed = self.app.parse_sales_report_file(seller, seller.name)
        product_parsed = self.app.parse_sales_report_file(product, product.name)

        self.assertEqual(daily_parsed["report_type"], "daily")
        self.assertEqual(daily_parsed["report_date"], "2026-06-19")
        self.assertEqual(daily_parsed["period"], "2026-06")
        self.assertEqual(daily_parsed["rows"][0]["profit_sales_amount"], 1200)
        self.assertEqual(seller_parsed["report_type"], "seller")
        self.assertEqual(seller_parsed["rows"][0]["name"], "A판매사")
        self.assertEqual(product_parsed["report_type"], "product")
        self.assertEqual(product_parsed["rows"][0]["name"], "테스트 상품 A")

    def test_sales_report_type_prefers_product_filename_hint(self) -> None:
        source = Path(self.tempdir.name) / "상품별 매출현황.xls"
        self.write_seller_report(source)

        self.assertEqual(self.app.detect_sales_report_type(source, source.name), "product")

    def test_sales_report_dashboard_combines_daily_seller_and_product_reports(self) -> None:
        base = Path(self.tempdir.name)
        daily = base / "매출 통계.xlsx"
        seller = base / "매출처별.xlsx"
        product = base / "Statistics_Good_2026-06-19.xls"
        self.write_daily_report(daily)
        self.write_seller_report(seller)
        self.write_product_report(product)

        for source in (daily, seller, product):
            self.app.save_sales_report_file(source, source.name, "admin")

        dashboard = self.app.sales_report_dashboard_payload("2026-06", "2026-06-19")

        self.assertEqual(dashboard["selected_date"], "2026-06-19")
        self.assertEqual(dashboard["today"]["profit_sales_amount"], 1190)
        self.assertEqual(dashboard["yesterday"]["profit_sales_amount"], 1000)
        self.assertEqual(dashboard["comparison"]["profit_sales_amount_delta"], 190)
        self.assertEqual(dashboard["comparison"]["profit_sales_amount_delta_rate"], 19.0)
        self.assertEqual(dashboard["month"]["profit_sales_amount"], 2190)
        self.assertEqual(dashboard["seller_total"]["profit_sales_amount"], 1190)
        self.assertEqual(dashboard["consistency"]["difference"], 1000)
        self.assertEqual(dashboard["seller_top"][0]["name"], "A판매사")
        self.assertEqual(dashboard["product_top"][0]["name"], "테스트 상품 A")
        self.assertTrue(dashboard["reviews"])


    def test_sales_report_dashboard_uses_latest_period_without_daily_report(self) -> None:
        base = Path(self.tempdir.name)
        seller = base / "seller.xlsx"
        product = base / "Statistics_Good_2026-06-19.xls"
        self.write_seller_report(seller)
        self.write_product_report(product)

        for source in (seller, product):
            self.app.save_sales_report_file(source, source.name, "admin")

        dashboard = self.app.sales_report_dashboard_payload()

        self.assertEqual(dashboard["period"], "2026-06")
        self.assertEqual(dashboard["selected_date"], "2026-06-19")
        self.assertFalse(dashboard["daily_rows"])
        self.assertEqual(dashboard["seller_top"][0]["profit_sales_amount"], 900)
        self.assertEqual(dashboard["product_top"][0]["profit_sales_amount"], 700)

    def test_sales_report_dashboard_returns_selectable_month_options(self) -> None:
        connection = self.app.connect_db()
        try:
            cursor = connection.execute(
                """
                INSERT INTO sales_report_uploads
                    (stored_name, original_name, report_type, report_date, period, size, uploaded_by, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ("daily-june.xlsx", "daily-june.xlsx", "daily", "2026-06-30", "2026-06", 1, "admin", "2026-06-30"),
            )
            june_file_id = int(cursor.lastrowid)
            cursor = connection.execute(
                """
                INSERT INTO sales_report_uploads
                    (stored_name, original_name, report_type, report_date, period, size, uploaded_by, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ("daily-july.xlsx", "daily-july.xlsx", "daily", "2026-07-01", "2026-07", 1, "admin", "2026-07-01"),
            )
            july_file_id = int(cursor.lastrowid)
            connection.executemany(
                """
                INSERT INTO sales_report_daily_rows
                    (report_date, period, file_id, label, quantity, profit_sales_amount, profit_margin)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    ("2026-06-30", "2026-06", june_file_id, "2026-06-30", 6, 6000, 900),
                    ("2026-07-01", "2026-07", july_file_id, "2026-07-01", 7, 7000, 1000),
                ],
            )
            connection.commit()
        finally:
            connection.close()

        latest_dashboard = self.app.sales_report_dashboard_payload()
        june_dashboard = self.app.sales_report_dashboard_payload("2026-06", "2026-06-30")

        self.assertEqual(latest_dashboard["period"], "2026-07")
        self.assertEqual(latest_dashboard["period_options"], ["2026-07", "2026-06"])
        self.assertEqual(june_dashboard["period"], "2026-06")
        self.assertEqual(june_dashboard["period_options"], ["2026-07", "2026-06"])
        self.assertEqual(june_dashboard["month"]["profit_sales_amount"], 6000)

    def test_sales_report_dashboard_keeps_recent_sales_dates_across_month_boundary(self) -> None:
        connection = self.app.connect_db()
        try:
            cursor = connection.execute(
                """
                INSERT INTO sales_report_uploads
                    (stored_name, original_name, report_type, report_date, period, size, uploaded_by, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ("daily.xlsx", "daily.xlsx", "daily", "2026-07-01", "2026-07", 1, "admin", "2026-07-01"),
            )
            file_id = int(cursor.lastrowid)
            connection.executemany(
                """
                INSERT INTO sales_report_daily_rows
                    (report_date, period, file_id, label, quantity, sales_amount, sales_total,
                     profit_sales_amount, profit_supply_amount, profit_shipping, profit_margin)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    ("2026-06-30", "2026-06", file_id, "2026-06-30", 10, 1000, 1000, 1000, 500, 0, 500),
                    ("2026-07-01", "2026-07", file_id, "2026-07-01", 20, 2000, 2000, 2000, 1000, 0, 1000),
                ],
            )
            connection.commit()
        finally:
            connection.close()

        dashboard = self.app.sales_report_dashboard_payload("2026-07", "2026-07-01")

        self.assertEqual([row["report_date"] for row in dashboard["daily_rows"]], ["2026-07-01"])
        self.assertEqual(
            [row["report_date"] for row in dashboard["recent_daily_rows"]],
            ["2026-07-01", "2026-06-30"],
        )
        self.assertEqual(dashboard["month"]["profit_sales_amount"], 2000)

    def test_daily_dimension_files_use_filename_date_for_dashboard_daily_summary(self) -> None:
        base = Path(self.tempdir.name)
        seller = base / "20260630_seller.xlsx"
        product = base / "20260630_product.xls"
        supplier = base / "20260630_supplier.xls"
        self.write_seller_report(seller)
        self.write_product_report(product)
        self.write_supplier_report(supplier)

        for source in (seller, product, supplier):
            saved = self.app.save_sales_report_file(source, source.name, "admin")
            self.assertEqual(saved["report_date"], "2026-06-30")
            self.assertEqual(saved["period"], "2026-06")

        dashboard = self.app.sales_report_dashboard_payload("2026-06", "2026-06-30")

        self.assertTrue(dashboard["today_data_uploaded"])
        self.assertEqual(dashboard["selected_date"], "2026-06-30")
        self.assertEqual(dashboard["today"]["report_date"], "2026-06-30")
        self.assertEqual(dashboard["today"]["quantity"], 10)
        self.assertEqual(dashboard["today"]["profit_sales_amount"], 1190)
        self.assertEqual(dashboard["month"]["profit_sales_amount"], 1190)
        self.assertEqual(dashboard["daily_rows"][0]["report_date"], "2026-06-30")
        self.assertEqual(dashboard["daily_rows"][0]["profit_sales_amount"], 1190)

    def test_seller_daily_file_extends_existing_daily_month_rows(self) -> None:
        base = Path(self.tempdir.name)
        daily = base / "daily.xlsx"
        seller = base / "20260630_seller.xlsx"
        self.write_daily_report(daily)
        self.write_seller_report(seller)

        self.app.save_sales_report_file(daily, daily.name, "admin")
        self.app.save_sales_report_file(seller, seller.name, "admin")

        dashboard = self.app.sales_report_dashboard_payload("2026-06", "2026-06-30")
        daily_by_date = {row["report_date"]: row for row in dashboard["daily_rows"]}

        self.assertEqual(dashboard["today"]["profit_sales_amount"], 1190)
        self.assertEqual(dashboard["month"]["profit_sales_amount"], 3390)
        self.assertIn("2026-06-19", daily_by_date)
        self.assertIn("2026-06-30", daily_by_date)
        self.assertEqual(daily_by_date["2026-06-30"]["profit_sales_amount"], 1190)

    def test_sales_report_dashboard_excludes_purchase_marked_products(self) -> None:
        connection = self.app.connect_db()
        try:
            cursor = connection.execute(
                """
                INSERT INTO sales_report_uploads
                    (stored_name, original_name, report_type, report_date, period, size, uploaded_by, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ("product.xlsx", "product.xlsx", "product", "2026-06-19", "2026-06", 1, "admin", "2026-06-19"),
            )
            file_id = int(cursor.lastrowid)
            connection.executemany(
                """
                INSERT INTO sales_report_product_rows
                    (period, report_date, file_id, product_code, product_name, quantity, profit_sales_amount, profit_margin)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    ("2026-06", "2026-06-19", file_id, "P001", "정상 상품", 2, 1000, 300),
                    ("2026-06", "2026-06-19", file_id, "P002", "★사입건★ 제외 상품", 5, 9000, -500),
                ],
            )
            connection.commit()
        finally:
            connection.close()

        dashboard = self.app.sales_report_dashboard_payload("2026-06", "2026-06-19")

        self.assertEqual([row["name"] for row in dashboard["product_top"]], ["정상 상품"])
        self.assertEqual(dashboard["product_total"]["quantity"], 2)
        self.assertEqual(dashboard["product_total"]["profit_sales_amount"], 1000)
        self.assertEqual([row["label"] for row in dashboard["monthly_comparison_details"]["product"]], ["정상 상품"])

    def test_sales_report_detail_payloads_are_limited_to_selected_month(self) -> None:
        base = Path(self.tempdir.name)
        daily = base / "매출 통계.xlsx"
        seller = base / "매출처별.xlsx"
        product = base / "Statistics_Good_2026-06-19.xls"
        supplier = base / "Statistics_Sales_Suppler_2026-06-19.xls"
        self.write_daily_report(daily)
        self.write_seller_report(seller)
        self.write_product_report(product)
        self.write_supplier_report(supplier)

        for source in (daily, seller, product, supplier):
            self.app.save_sales_report_file(source, source.name, "admin")

        connection = self.app.connect_db()
        try:
            connection.execute(
                """
                INSERT INTO management_records
                    (created_at, source_file, source_sheet, source_row, sales_vendor, purchase_vendor, order_date, ship_date, product_name, quantity)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ("2026-06-20", "manual", "sheet", 1, "A판매사", "공급사A", "2026-06-19", "2026-06-19", "테스트 상품 A", "3"),
            )
            connection.execute(
                """
                INSERT INTO management_records
                    (created_at, source_file, source_sheet, source_row, sales_vendor, purchase_vendor, order_date, ship_date, product_name, quantity)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ("2026-07-01", "manual", "sheet", 2, "A판매사", "공급사A", "2026-07-01", "2026-07-01", "테스트 상품 A", "99"),
            )
            connection.commit()
        finally:
            connection.close()

        daily_detail = self.app.sales_report_detail_payload("daily", "2026-06-19", "2026-06")
        product_detail = self.app.sales_report_detail_payload("product", "테스트 상품 A", "2026-06")
        seller_detail = self.app.sales_report_detail_payload("seller", "A판매사", "2026-06")
        supplier_detail = self.app.sales_report_detail_payload("supplier", "공급사A", "2026-06")

        self.assertEqual(daily_detail["sections"][0]["rows"][0], ["A판매사", 1, 3, 900, 950, 800])
        self.assertNotIn("일자별 상품 손익", [section["title"] for section in product_detail["sections"]])
        self.assertEqual(product_detail["sections"][0]["rows"][0], ["2026-06-19", 1, 3])
        self.assertEqual(product_detail["sections"][1]["rows"][0], ["A판매사", 1, 3])
        self.assertEqual(daily_detail["sections"][1]["rows"][0], ["테스트 상품 A", "A판매사", 1, 3])
        self.assertEqual(seller_detail["sections"][0]["rows"], [["2026-06-19", 1, 3]])
        self.assertEqual(seller_detail["sections"][1]["rows"], [["2026-06-19", "테스트 상품 A", 1, 3]])
        self.assertEqual(seller_detail["sections"][2]["rows"], [["2026-06-19", 900, 800]])
        self.assertEqual(supplier_detail["sections"][0]["rows"], [["2026-06-19", 1, 3]])
        self.assertEqual(supplier_detail["sections"][1]["rows"], [["2026-06-19", "테스트 상품 A", 1, 3]])
        self.assertEqual(supplier_detail["sections"][2]["rows"], [["2026-06-19", 750, 220]])

    def test_sales_report_margin_check_flags_margin_over_sales(self) -> None:
        connection = self.app.connect_db()
        try:
            cursor = connection.execute(
                """
                INSERT INTO sales_report_uploads
                    (stored_name, original_name, report_type, report_date, period, size, uploaded_by, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ("daily.xlsx", "daily.xlsx", "daily", "2026-06-20", "2026-06", 1, "admin", "2026-06-20"),
            )
            file_id = int(cursor.lastrowid)
            connection.execute(
                """
                INSERT INTO sales_report_daily_rows
                    (report_date, period, file_id, label, quantity, profit_sales_amount, profit_supply_amount, cs_margin, profit_margin)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ("2026-06-20", "2026-06", file_id, "2026-06-20", 1, 100, -20, 0, 120),
            )
            connection.commit()
            check = self.app.sales_report_margin_check(connection, "2026-06")
        finally:
            connection.close()

        self.assertEqual(check["anomaly_count"], 1)
        self.assertIn("손익 공급금액", check["message"])

    def test_date_based_seller_uploads_preserve_previous_days(self) -> None:
        base = Path(self.tempdir.name)
        day23 = base / "seller-2026-06-23.xlsx"
        day24 = base / "seller-2026-06-24.xlsx"
        day23.write_text("placeholder", encoding="utf-8")
        day24.write_text("placeholder", encoding="utf-8")

        original_detect = self.app.detect_sales_report_type
        original_parse = self.app.parse_sales_report_file

        def fake_detect(path: Path, original_name: str = "") -> str:
            return "seller"

        def fake_parse(path: Path, original_name: str = "") -> dict[str, object]:
            report_date = "2026-06-23" if "23" in original_name else "2026-06-24"
            quantity = 1 if report_date == "2026-06-23" else 2
            profit_sales_amount = 100 if report_date == "2026-06-23" else 200
            profit_margin = 80 if report_date == "2026-06-23" else 160
            return {
                "report_type": "seller",
                "report_date": report_date,
                "period": "2026-06",
                "rows": [{"name": "A판매사", "quantity": quantity, "profit_sales_amount": profit_sales_amount, "profit_margin": profit_margin}],
            }

        try:
            self.app.detect_sales_report_type = fake_detect
            self.app.parse_sales_report_file = fake_parse
            self.app.save_sales_report_file(day23, day23.name, "admin")
            self.app.save_sales_report_file(day24, day24.name, "admin")
        finally:
            self.app.detect_sales_report_type = original_detect
            self.app.parse_sales_report_file = original_parse

        connection = self.app.connect_db()
        try:
            rows = connection.execute(
                "SELECT report_date FROM sales_report_seller_rows WHERE period = ? ORDER BY report_date",
                ("2026-06",),
            ).fetchall()
        finally:
            connection.close()

        self.assertEqual([row["report_date"] for row in rows], ["2026-06-23", "2026-06-24"])

        dashboard = self.app.sales_report_dashboard_payload("2026-06", "2026-06-24")
        self.assertEqual(len(dashboard["seller_top"]), 1)
        self.assertEqual(dashboard["seller_top"][0]["name"], "A판매사")
        self.assertEqual(dashboard["seller_top"][0]["quantity"], 3)
        self.assertEqual(dashboard["seller_top"][0]["profit_sales_amount"], 300)

    def test_same_date_dimension_upload_replaces_previous_snapshot(self) -> None:
        connection = self.app.connect_db()
        try:
            cursor = connection.execute(
                """
                INSERT INTO sales_report_uploads
                    (stored_name, original_name, report_type, report_date, period, size, uploaded_by, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ("first.xlsx", "first.xlsx", "seller", "2026-06-24", "2026-06", 1, "admin", "2026-06-24"),
            )
            first_id = int(cursor.lastrowid)
            cursor = connection.execute(
                """
                INSERT INTO sales_report_uploads
                    (stored_name, original_name, report_type, report_date, period, size, uploaded_by, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ("last.xlsx", "last.xlsx", "seller", "2026-06-24", "2026-06", 1, "admin", "2026-06-24"),
            )
            last_id = int(cursor.lastrowid)
            connection.commit()
        finally:
            connection.close()

        self.app.save_sales_report_snapshot(
            first_id,
            {
                "report_type": "seller",
                "report_date": "2026-06-24",
                "period": "2026-06",
                "rows": [{"name": "A판매사", "quantity": 1, "profit_sales_amount": 100, "profit_margin": 80}],
            },
        )
        self.app.save_sales_report_snapshot(
            last_id,
            {
                "report_type": "seller",
                "report_date": "2026-06-24",
                "period": "2026-06",
                "rows": [{"name": "A판매사", "quantity": 9, "profit_sales_amount": 900, "profit_margin": 720}],
            },
        )

        connection = self.app.connect_db()
        try:
            rows = connection.execute(
                """
                SELECT file_id, quantity, profit_sales_amount
                  FROM sales_report_seller_rows
                 WHERE period = ? AND report_date = ? AND seller_name = ?
                """,
                ("2026-06", "2026-06-24", "A판매사"),
            ).fetchall()
        finally:
            connection.close()

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["file_id"], last_id)
        self.assertEqual(rows[0]["quantity"], 9)
        self.assertEqual(rows[0]["profit_sales_amount"], 900)

        connection = self.app.connect_db()
        try:
            upload_ids: dict[str, int] = {}
            for report_type in ("supplier", "product", "daily"):
                cursor = connection.execute(
                    """
                    INSERT INTO sales_report_uploads
                        (stored_name, original_name, report_type, report_date, period, size, uploaded_by, uploaded_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (f"{report_type}.xlsx", f"{report_type}.xlsx", report_type, "2026-06-24", "2026-06", 1, "admin", "2026-06-24"),
                )
                upload_ids[report_type] = int(cursor.lastrowid)
            connection.commit()
        finally:
            connection.close()

        self.app.save_sales_report_snapshot(
            upload_ids["supplier"],
            {
                "report_type": "supplier",
                "report_date": "2026-06-24",
                "period": "2026-06",
                "rows": [{"name": "A공급사", "quantity": 1, "supply_total": 1000}],
            },
        )
        self.app.save_sales_report_snapshot(
            upload_ids["supplier"],
            {
                "report_type": "supplier",
                "report_date": "2026-06-24",
                "period": "2026-06",
                "rows": [{"name": "A공급사", "quantity": 8, "supply_total": 8000}],
            },
        )
        self.app.save_sales_report_snapshot(
            upload_ids["product"],
            {
                "report_type": "product",
                "report_date": "2026-06-24",
                "period": "2026-06",
                "rows": [{"code": "P001", "name": "A상품", "quantity": 1, "profit_sales_amount": 100}],
            },
        )
        self.app.save_sales_report_snapshot(
            upload_ids["product"],
            {
                "report_type": "product",
                "report_date": "2026-06-24",
                "period": "2026-06",
                "rows": [{"code": "P001", "name": "A상품", "quantity": 7, "profit_sales_amount": 700}],
            },
        )
        self.app.save_sales_report_snapshot(
            upload_ids["daily"],
            {
                "report_type": "daily",
                "report_date": "2026-06-24",
                "period": "2026-06",
                "rows": [{"report_date": "2026-06-24", "period": "2026-06", "label": "2026-06-24", "quantity": 1, "profit_sales_amount": 100}],
            },
        )
        self.app.save_sales_report_snapshot(
            upload_ids["daily"],
            {
                "report_type": "daily",
                "report_date": "2026-06-24",
                "period": "2026-06",
                "rows": [{"report_date": "2026-06-24", "period": "2026-06", "label": "2026-06-24", "quantity": 6, "profit_sales_amount": 600}],
            },
        )

        connection = self.app.connect_db()
        try:
            supplier_row = connection.execute(
                "SELECT quantity, supply_total FROM sales_report_supplier_rows WHERE report_date = ? AND supplier_name = ?",
                ("2026-06-24", "A공급사"),
            ).fetchone()
            product_row = connection.execute(
                "SELECT quantity, profit_sales_amount FROM sales_report_product_rows WHERE report_date = ? AND product_name = ?",
                ("2026-06-24", "A상품"),
            ).fetchone()
            daily_row = connection.execute(
                "SELECT quantity, profit_sales_amount FROM sales_report_daily_rows WHERE report_date = ?",
                ("2026-06-24",),
            ).fetchone()
        finally:
            connection.close()

        self.assertEqual(dict(supplier_row), {"quantity": 8, "supply_total": 8000})
        self.assertEqual(dict(product_row), {"quantity": 7, "profit_sales_amount": 700})
        self.assertEqual(dict(daily_row), {"quantity": 6, "profit_sales_amount": 600})

    def test_sales_report_date_from_upload_name_uses_stored_timestamp(self) -> None:
        parsed = self.app.sales_report_date_from_upload_name(
            "20260623072000_925ce8e90482e99a354c_매출처별 매출 합계.xls",
            "2026-06",
        )

        self.assertEqual(parsed, "2026-06-23")

    def test_sales_report_date_from_upload_name_accepts_compact_single_digit_month(self) -> None:
        parsed = self.app.sales_report_date_from_upload_name(
            "2026629 거래처별 매출현황.xlsx",
            "2026-06",
        )

        self.assertEqual(parsed, "2026-06-29")

    def test_sales_report_parser_dates_accept_compact_daily_filename(self) -> None:
        filename = "20260701 supplier sales report.xls"

        self.assertEqual(self.app._sales_report_date(filename), "2026-07-01")
        self.assertEqual(self.app._sales_report_period(filename), "2026-07")

    def test_partner_daily_zip_imports_seller_rows_until_yesterday(self) -> None:
        base = Path(self.tempdir.name)
        archive_path = base / "partner-daily.zip"
        today = date.today()
        yesterday = today - timedelta(days=1)
        with zipfile.ZipFile(archive_path, "w") as archive:
            archive.writestr("A거래처 매출 현황.xlsx", b"a")
            archive.writestr("B거래처 매출 현황.xlsx", b"b")

        original_parse = self.app.parse_sales_report_file

        def fake_parse(path: Path, original_name: str = "") -> dict[str, object]:
            amount = 100 if original_name.startswith("A") else 200
            return {
                "report_type": "daily",
                "report_date": yesterday.isoformat(),
                "period": yesterday.isoformat()[:7],
                "rows": [
                    {
                        "report_date": yesterday.isoformat(),
                        "period": yesterday.isoformat()[:7],
                        "quantity": 1,
                        "profit_sales_amount": amount,
                        "profit_shipping": 10,
                        "profit_margin": amount + 10,
                    },
                    {
                        "report_date": today.isoformat(),
                        "period": today.isoformat()[:7],
                        "quantity": 9,
                        "profit_sales_amount": 999,
                        "profit_shipping": 0,
                        "profit_margin": 999,
                    },
                ],
            }

        try:
            self.app.parse_sales_report_file = fake_parse
            saved = self.app.save_sales_report_file(archive_path, archive_path.name, "admin")
        finally:
            self.app.parse_sales_report_file = original_parse

        self.assertEqual(saved["report_type"], "seller_daily_zip")
        self.assertEqual(saved["row_count"], 2)
        self.assertEqual(saved["source_file_count"], 2)
        self.assertEqual(saved["skipped_current_or_future_rows"], 2)

        connection = self.app.connect_db()
        try:
            rows = connection.execute(
                """
                SELECT report_date, seller_name, quantity, profit_sales_amount, profit_margin, profit_shipping
                  FROM sales_report_seller_rows
                 ORDER BY seller_name
                """
            ).fetchall()
        finally:
            connection.close()

        self.assertEqual([row["report_date"] for row in rows], [yesterday.isoformat(), yesterday.isoformat()])
        self.assertEqual([row["seller_name"] for row in rows], ["A거래처", "B거래처"])
        self.assertEqual([row["profit_sales_amount"] for row in rows], [100, 200])

    def test_dashboard_uses_previous_actual_sales_date_and_month_to_date_sellers(self) -> None:
        connection = self.app.connect_db()
        try:
            cursor = connection.execute(
                """
                INSERT INTO sales_report_uploads
                    (stored_name, original_name, report_type, report_date, period, size, uploaded_by, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ("seller-daily.zip", "seller-daily.zip", "seller_daily_zip", "2026-06-29", "2026-06", 1, "admin", "2026-06-29"),
            )
            file_id = int(cursor.lastrowid)
            cursor = connection.execute(
                """
                INSERT INTO sales_report_uploads
                    (stored_name, original_name, report_type, report_date, period, size, uploaded_by, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ("supplier-daily.xlsx", "supplier-daily.xlsx", "supplier", "2026-06-29", "2026-06", 1, "admin", "2026-06-29"),
            )
            supplier_file_id = int(cursor.lastrowid)
            connection.commit()
        finally:
            connection.close()

        self.app.save_sales_report_snapshot(
            file_id,
            {
                "report_type": "seller_daily_zip",
                "report_date": "2026-06-29",
                "period": "2026-06",
                "rows": [
                    {"name": "A거래처", "report_date": "2026-06-26", "period": "2026-06", "quantity": 1, "profit_sales_amount": 100, "profit_margin": 100},
                    {"name": "A거래처", "report_date": "2026-06-27", "period": "2026-06", "quantity": 0, "profit_sales_amount": 0, "profit_margin": 0},
                    {"name": "B거래처", "report_date": "2026-06-28", "period": "2026-06", "quantity": 0, "profit_sales_amount": 0, "profit_margin": 0},
                    {"name": "C거래처", "report_date": "2026-06-29", "period": "2026-06", "quantity": 3, "profit_sales_amount": 300, "profit_margin": 300},
                ],
            },
        )
        connection = self.app.connect_db()
        try:
            connection.executemany(
                """
                INSERT INTO sales_report_supplier_rows
                    (period, report_date, file_id, supplier_name, quantity, supply_total, profit_supply_amount)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    ("2026-06", "2026-06-26", supplier_file_id, "A매입처", 1, 1000, 1000),
                    ("2026-06", "2026-06-28", supplier_file_id, "A매입처", 0, 0, 0),
                    ("2026-06", "2026-06-29", supplier_file_id, "B매입처", 2, 2000, 2000),
                ],
            )
            connection.commit()
        finally:
            connection.close()

        dashboard = self.app.sales_report_dashboard_payload("2026-06", "2026-06-29")

        self.assertEqual(dashboard["previous_business_date"], "2026-06-26")
        self.assertEqual(dashboard["yesterday"]["profit_sales_amount"], 100)
        self.assertEqual(dashboard["month"]["profit_sales_amount"], 400)
        self.assertEqual(dashboard["seller_total"]["profit_sales_amount"], 400)
        self.assertEqual(dashboard["supplier_purchase_total"]["purchase_total"], 3000)
        seller_by_name = {row["name"]: row for row in dashboard["seller_top"]}
        supplier_by_name = {row["name"]: row for row in dashboard["supplier_purchase_totals"]}
        self.assertEqual(seller_by_name["A거래처"]["profit_sales_amount"], 100)
        self.assertEqual(seller_by_name["C거래처"]["profit_sales_amount"], 300)
        self.assertEqual(supplier_by_name["A매입처"]["purchase_total"], 1000)
        self.assertEqual(supplier_by_name["B매입처"]["purchase_total"], 2000)

    def test_sales_partner_detail_uses_raw_daily_amounts_not_deltas(self) -> None:
        connection = self.app.connect_db()
        try:
            cursor = connection.execute(
                """
                INSERT INTO sales_report_uploads
                    (stored_name, original_name, report_type, report_date, period, size, uploaded_by, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ("seller-daily.zip", "seller-daily.zip", "seller_daily_zip", "2026-06-29", "2026-06", 1, "admin", "2026-06-29"),
            )
            file_id = int(cursor.lastrowid)
            connection.commit()
        finally:
            connection.close()

        self.app.save_sales_report_snapshot(
            file_id,
            {
                "report_type": "seller_daily_zip",
                "report_date": "2026-06-29",
                "period": "2026-06",
                "rows": [
                    {"name": "A거래처", "report_date": "2026-06-04", "period": "2026-06", "quantity": 4, "profit_sales_amount": 2680000, "profit_margin": 2658500},
                    {"name": "A거래처", "report_date": "2026-06-05", "period": "2026-06", "quantity": -1, "profit_sales_amount": -225000, "profit_margin": -203500},
                    {"name": "A거래처", "report_date": "2026-06-06", "period": "2026-06", "quantity": -2, "profit_sales_amount": -2455000, "profit_margin": -2455000},
                ],
            },
        )

        detail = self.app.sales_report_detail_payload("seller", "A거래처", "2026-06")
        amount_rows = detail["sections"][2]["rows"]

        self.assertEqual(detail["metrics"][2]["value"], 0)
        self.assertEqual(
            amount_rows,
            [
                ["2026-06-04", 2680000, 2658500],
                ["2026-06-05", -225000, -203500],
                ["2026-06-06", -2455000, -2455000],
            ],
        )

    def test_sales_dimension_reports_without_date_inherit_latest_daily_context(self) -> None:
        base = Path(self.tempdir.name)
        daily = base / "daily.xlsx"
        product = base / "product.xlsx"
        supplier = base / "supplier.xlsx"
        self.write_daily_report(daily)
        product.write_text("product placeholder", encoding="utf-8")
        supplier.write_text("supplier placeholder", encoding="utf-8")

        self.app.save_sales_report_file(daily, daily.name, "admin")

        original_detect = self.app.detect_sales_report_type
        original_parse = self.app.parse_sales_report_file

        def fake_detect(path: Path, original_name: str = "") -> str:
            return "supplier" if "supplier" in original_name else "product"

        def fake_parse(path: Path, original_name: str = "") -> dict[str, object]:
            if "supplier" in original_name:
                return {
                    "report_type": "supplier",
                    "report_date": "",
                    "period": "",
                    "rows": [{"name": "테스트 공급사", "quantity": 3, "supply_total": 1500}],
                }
            return {
                "report_type": "product",
                "report_date": "",
                "period": "",
                "rows": [{"code": "P001", "name": "테스트 상품", "quantity": 2, "profit_sales_amount": 1000}],
            }

        try:
            self.app.detect_sales_report_type = fake_detect
            self.app.parse_sales_report_file = fake_parse
            product_saved = self.app.save_sales_report_file(product, product.name, "admin")
            supplier_saved = self.app.save_sales_report_file(supplier, supplier.name, "admin")
        finally:
            self.app.detect_sales_report_type = original_detect
            self.app.parse_sales_report_file = original_parse

        self.assertEqual(product_saved["report_date"], "2026-06-19")
        self.assertEqual(product_saved["period"], "2026-06")
        self.assertEqual(supplier_saved["report_date"], "2026-06-19")
        self.assertEqual(supplier_saved["period"], "2026-06")

        connection = self.app.connect_db()
        try:
            product_row = connection.execute(
                "SELECT report_date, period FROM sales_report_product_rows WHERE product_name = ?",
                ("테스트 상품",),
            ).fetchone()
            supplier_row = connection.execute(
                "SELECT report_date, period FROM sales_report_supplier_rows WHERE supplier_name = ?",
                ("테스트 공급사",),
            ).fetchone()
        finally:
            connection.close()

        self.assertEqual(dict(product_row), {"report_date": "2026-06-19", "period": "2026-06"})
        self.assertEqual(dict(supplier_row), {"report_date": "2026-06-19", "period": "2026-06"})

    def test_sales_supplier_report_feeds_purchase_total_panel(self) -> None:
        base = Path(self.tempdir.name)
        supplier = base / "Statistics_Sales_Suppler_2026-06-19.xls"
        self.write_supplier_report(supplier)

        self.assertEqual(self.app.detect_sales_report_type(supplier, supplier.name), "supplier")
        parsed = self.app.parse_sales_report_file(supplier, supplier.name)
        self.assertEqual(parsed["report_type"], "supplier")
        self.assertEqual(parsed["period"], "2026-06")
        self.assertEqual(parsed["report_date"], "2026-06-19")
        self.assertEqual(parsed["rows"][0]["name"], "공급사A")
        self.assertEqual(parsed["rows"][0]["cs_amount"], -80)

        self.app.save_sales_report_file(supplier, supplier.name, "admin")
        dashboard = self.app.sales_report_dashboard_payload("2026-06", "2026-06-19")
        purchase_totals = dashboard["supplier_purchase_totals"]
        purchase_total = dashboard["supplier_purchase_total"]

        self.assertEqual(purchase_total["quantity"], 7)
        self.assertEqual(purchase_total["purchase_total"], 1600)
        self.assertEqual(purchase_totals[0]["name"], "공급사B")
        self.assertEqual(purchase_totals[0]["purchase_total"], 900)
        self.assertEqual(purchase_totals[1]["name"], "공급사A")
        self.assertEqual(purchase_totals[1]["purchase_total"], 700)


    def test_supplier_purchase_total_uses_monthly_profit_supply_amount(self) -> None:
        base = Path(self.tempdir.name)
        first = base / "supplier-20260624.xlsx"
        second = base / "supplier-20260630.xlsx"
        first.write_text("first supplier placeholder", encoding="utf-8")
        second.write_text("second supplier placeholder", encoding="utf-8")

        original_detect = self.app.detect_sales_report_type
        original_parse = self.app.parse_sales_report_file

        def fake_detect(path: Path, original_name: str = "") -> str:
            return "supplier"

        def fake_parse(path: Path, original_name: str = "") -> dict[str, object]:
            if "20260624" in original_name:
                return {
                    "report_type": "supplier",
                    "report_date": "2026-06-24",
                    "period": "2026-06",
                    "rows": [
                        {"name": "A공급사", "quantity": 10, "supply_total": 1200, "profit_supply_amount": 1000},
                        {"name": "B공급사", "quantity": 5, "supply_total": 700, "profit_supply_amount": 600},
                    ],
                }
            return {
                "report_type": "supplier",
                "report_date": "2026-06-30",
                "period": "2026-06",
                "rows": [
                    {"name": "A공급사", "quantity": 3, "supply_total": 450, "profit_supply_amount": 400},
                    {"name": "C공급사", "quantity": 2, "supply_total": 350, "profit_supply_amount": 300},
                ],
            }

        try:
            self.app.detect_sales_report_type = fake_detect
            self.app.parse_sales_report_file = fake_parse
            self.app.save_sales_report_file(first, first.name, "admin")
            self.app.save_sales_report_file(second, second.name, "admin")
        finally:
            self.app.detect_sales_report_type = original_detect
            self.app.parse_sales_report_file = original_parse

        dashboard = self.app.sales_report_dashboard_payload("2026-06", "2026-06-30")
        purchase_by_name = {row["name"]: row for row in dashboard["supplier_purchase_totals"]}

        self.assertEqual(dashboard["supplier_purchase_total"]["purchase_total"], 2300)
        self.assertEqual(dashboard["supplier_purchase_total"]["quantity"], 20)
        self.assertEqual(purchase_by_name["A공급사"]["purchase_total"], 1400)
        self.assertEqual(purchase_by_name["B공급사"]["purchase_total"], 600)
        self.assertEqual(purchase_by_name["C공급사"]["purchase_total"], 300)

    def test_sales_report_nas_scan_imports_and_archives_new_files_once(self) -> None:
        base = Path(self.tempdir.name)
        nas_dir = base / "nas-sales"
        nas_dir.mkdir()
        source = nas_dir / "Statistics_Sales_Suppler_2026-06-19.xls"
        self.write_supplier_report(source)

        settings = self.app.save_sales_automation_settings({
            "nas_enabled": True,
            "nas_import_dir": str(nas_dir),
            "nas_processed_dir": "processed",
            "nas_error_dir": "error",
            "nas_scan_interval_minutes": 10,
        })
        first_result = self.app.scan_sales_report_nas_folder(settings, uploaded_by="auto-nas")

        self.assertEqual(first_result["imported_count"], 1)
        self.assertEqual(first_result["error_count"], 0)
        self.assertFalse(source.exists())
        self.assertTrue((nas_dir / "processed" / "Statistics_Sales_Suppler_2026-06-19.xls").exists())

        files = self.app.list_sales_report_uploads()
        self.assertEqual(len(files), 1)
        self.assertEqual(files[0]["uploaded_by"], "auto-nas")

        duplicate = nas_dir / "Statistics_Sales_Suppler_2026-06-19.xls"
        self.write_supplier_report(duplicate)
        second_result = self.app.scan_sales_report_nas_folder(settings, uploaded_by="auto-nas")

        self.assertEqual(second_result["imported_count"], 0)
        self.assertEqual(second_result["duplicate_count"], 1)
        self.assertEqual(len(self.app.list_sales_report_uploads()), 1)


if __name__ == "__main__":
    unittest.main()
