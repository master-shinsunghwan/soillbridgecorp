from __future__ import annotations

import importlib
import os
import sys
import tempfile
from datetime import datetime
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook, load_workbook


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


SUPPLY_HEADERS = [
    "순서",
    "매입거래처",
    "매출거래처",
    "거래구분",
    "장부입력확인",
    "주문일",
    "출고일",
    "주문자",
    "발신자연락처",
    "수령자",
    "수령자연락처",
    "제 품 명",
    "수량",
    "상 세 주 소",
    "택배사**",
    "배송번호",
    "특이(요청)사항",
    "주문상품고유번호",
    "상품코드",
    "주문번호",
    "고객선택옵션",
]


def create_supply_workbook(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Supply"
    worksheet.append(SUPPLY_HEADERS)
    worksheet.append(
        [
            1,
            "(주)소일브릿지(본사)",
            "토스",
            "",
            "",
            "2026-06-15 09:30:46",
            "",
            "김주문",
            "010-1111-2222",
            "이수령",
            "010-3333-4444",
            "[테스트] 감식초 1000ml",
            2,
            "서울시 테스트구 테스트로 1",
            "롯데택배",
            "1234567890",
            "문 앞",
            "111385084",
            "SOSO27788588",
            "WSO260615-000000021",
            "기본",
        ]
    )
    workbook.save(path)


def load_app(tmp_path: Path):
    os.environ["WORKHUB_DATA_DIR"] = str(tmp_path)
    sys.modules.pop("workhub_delivery_app", None)
    return importlib.import_module("workhub_delivery_app")


def test_delivery_summary_accepts_supply_format(tmp_path: Path) -> None:
    from delivery_text_summary import summarize_workbook

    workbook_path = tmp_path / "supply.xlsx"
    create_supply_workbook(workbook_path)

    text, sheet_names = summarize_workbook(workbook_path, sort_mode="first")

    assert sheet_names == ["Supply"]
    assert "[테스트] 감식초 1000ml - 2개 (1건)" in text


def test_management_import_accepts_supply_header_on_first_row(tmp_path: Path) -> None:
    app = load_app(tmp_path)
    workbook_path = tmp_path / "supply.xlsx"
    create_supply_workbook(workbook_path)

    inserted, skipped = app.import_management_workbook(workbook_path)
    rows = app.list_management_records(limit=None)

    assert (inserted, skipped) == (1, 0)
    assert rows[0]["receiver_name"] == "이수령"
    assert rows[0]["invoice_number"] == "1234567890"
    assert rows[0]["order_item_id"] == "111385084"
    assert rows[0]["product_code"] == "SOSO27788588"
    assert rows[0]["order_number"] == "WSO260615-000000021"
    assert rows[0]["customer_option"] == "기본"


def test_management_template_export_uses_supply_columns_and_filename(tmp_path: Path) -> None:
    app = load_app(tmp_path)
    rows = [
        {
            "purchase_vendor": "(주)소일브릿지(본사)",
            "sales_vendor": "토스",
            "order_date": "2026-06-15",
            "receiver_name": "이수령",
            "receiver_phone": "010-3333-4444",
            "product_name": "[테스트] 감식초 1000ml",
            "quantity": "2",
            "receiver_address": "서울시 테스트구 테스트로 1",
            "courier": "롯데택배",
            "invoice_number": "1234567890",
            "memo": "문 앞",
            "order_item_id": "111385084",
            "product_code": "SOSO27788588",
            "order_number": "WSO260615-000000021",
            "customer_option": "기본",
        }
    ]

    exported = load_workbook(BytesIO(app.management_workbook_bytes_from_template(rows)), data_only=True)
    worksheet = exported.worksheets[0]

    assert [worksheet.cell(1, column).value for column in range(1, 22)] == SUPPLY_HEADERS
    assert worksheet.cell(2, 1).value == "1"
    assert worksheet.cell(2, 18).value == "111385084"
    assert worksheet.cell(2, 20).value == "WSO260615-000000021"
    assert (
        app.management_template_filename_stem(rows, {}, datetime(2026, 6, 15))
        == "통합관리대장 양식 2026년 20260615"
    )


def main() -> None:
    with tempfile.TemporaryDirectory() as directory:
        test_delivery_summary_accepts_supply_format(Path(directory) / "summary")
    with tempfile.TemporaryDirectory() as directory:
        test_management_import_accepts_supply_header_on_first_row(Path(directory) / "import")
    with tempfile.TemporaryDirectory() as directory:
        test_management_template_export_uses_supply_columns_and_filename(Path(directory) / "export")


if __name__ == "__main__":
    main()
