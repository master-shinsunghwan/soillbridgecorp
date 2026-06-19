from __future__ import annotations

import importlib
import os
import sys
import tempfile
import unittest
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
        self.assertEqual(dashboard["today"]["profit_sales_amount"], 1200)
        self.assertEqual(dashboard["yesterday"]["profit_sales_amount"], 1000)
        self.assertEqual(dashboard["comparison"]["profit_sales_amount_delta"], 200)
        self.assertEqual(dashboard["comparison"]["profit_sales_amount_delta_rate"], 20.0)
        self.assertEqual(dashboard["month"]["profit_sales_amount"], 2200)
        self.assertEqual(dashboard["seller_total"]["profit_sales_amount"], 1190)
        self.assertEqual(dashboard["consistency"]["difference"], 1010)
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

        self.assertEqual(purchase_totals[0]["name"], "공급사B")
        self.assertEqual(purchase_totals[0]["purchase_total"], 900)
        self.assertEqual(purchase_totals[1]["name"], "공급사A")
        self.assertEqual(purchase_totals[1]["purchase_total"], 750)


if __name__ == "__main__":
    unittest.main()
