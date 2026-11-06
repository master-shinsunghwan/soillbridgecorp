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


class ReturnLabelParsingTests(unittest.TestCase):
    def setUp(self) -> None:
        sys.modules.pop("workhub_cs_return_ocr", None)
        self.ocr = importlib.import_module("workhub_cs_return_ocr")

    def test_prefers_tracking_number_near_heading(self) -> None:
        text = """
        운송장번호 8448-1987-5631
        보내는분 배*옥
        상품명 벨트리 페이셜 티슈 100매 10개
        반품 6012-4989-4354
        """

        parsed = self.ocr.extract_return_label_fields(text)

        self.assertEqual(parsed["tracking_number"], "8448-1987-5631")
        self.assertIn("배*옥", parsed["masked_names"])

    def test_matches_masked_name_and_product(self) -> None:
        parsed = self.ocr.extract_return_label_fields(
            "운송장번호 8448-2028-2810 보내는분 김*란 노르디쿡 무쇠팬 28cm 뒤집개"
        )
        cases = [
            {"id": 10, "receiver_name": "김영란", "product_name": "노르디쿡 IH 무쇠팬 28cm + 뒤집개", "receiver_phone": "010-1111-2222"},
            {"id": 11, "receiver_name": "김수란", "product_name": "다른 상품", "receiver_phone": "010-3333-4444"},
        ]

        result = self.ocr.match_return_label(parsed, cases)

        self.assertEqual(result["match_status"], "matched")
        self.assertEqual(result["case_id"], 10)

    def test_does_not_auto_match_tied_candidates(self) -> None:
        parsed = self.ocr.extract_return_label_fields(
            "운송장번호 8448-1111-2222 보내는분 김*란 노르디쿡 무쇠팬"
        )
        cases = [
            {"id": 20, "receiver_name": "김영란", "product_name": "노르디쿡 무쇠팬"},
            {"id": 21, "receiver_name": "김수란", "product_name": "노르디쿡 무쇠팬"},
        ]

        result = self.ocr.match_return_label(parsed, cases)

        self.assertEqual(result["match_status"], "review")
        self.assertEqual(result["case_id"], 0)
        self.assertEqual([item["id"] for item in result["candidates"]], [20, 21])

    def test_formats_editable_kakao_text(self) -> None:
        rows = [
            {
                "receiver_name": "배신옥",
                "product_name": "벨트리 페이셜 티슈 100매",
                "quantity": "10",
                "condition": "양품",
                "tracking_number": "8448-1987-5631",
                "selected": True,
            },
            {"receiver_name": "제외", "selected": False},
        ]

        text = self.ocr.format_kakao_return_text(rows, "2026-11-06")

        self.assertIn("★11월 6일(금) 택배 회수송장★", text)
        self.assertIn("배신옥 / 벨트리 페이셜 티슈 100매 - 10개", text)
        self.assertIn("회수송장: 8448-1987-5631", text)
        self.assertNotIn("제외", text)


class ReturnLabelUiContractTests(unittest.TestCase):
    def test_cs_ledger_contains_return_label_review_controls(self) -> None:
        source = (SCRIPTS / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('id="ledgerReturnLabelOcr"', source)
        self.assertIn('id="returnLabelOcrModal"', source)
        self.assertIn('id="returnLabelKakaoText"', source)
        self.assertIn('/api/cs-return-label-analyze', source)
        self.assertIn('/api/cs-return-label-apply', source)


class ReturnLabelPersistenceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        os.environ["WORKHUB_DATA_DIR"] = self.temp_dir.name
        sys.modules.pop("workhub_delivery_app", None)
        self.app = importlib.import_module("workhub_delivery_app")
        self.app.init_db()

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def create_case(self, name: str = "김영란", product: str = "노르디쿡 무쇠팬") -> int:
        return self.app.save_cs_case(
            {
                "cs_receiver": name,
                "cs_product": product,
                "cs_type": "불량반품",
                "cs_content": "코팅불량",
                "occurred_at": "2026-11-06",
            },
            status="회수지시(11/6)",
        )

    def test_applies_invoice_and_links_original_image_to_existing_case(self) -> None:
        case_id = self.create_case()
        attachment = self.app.save_cs_return_label_attachment("label.jpg", b"jpeg-data", "image/jpeg")

        result = self.app.apply_cs_return_label_rows(
            [{
                "action": "match",
                "case_id": case_id,
                "tracking_number": "8448-2028-2810",
                "attachment_token": attachment["token"],
            }]
        )

        self.assertEqual(result["updated"], 1)
        connection = self.app.connect_db()
        try:
            case = connection.execute("SELECT return_invoice FROM cs_cases WHERE id = ?", (case_id,)).fetchone()
            linked = connection.execute(
                "SELECT case_id FROM cs_case_attachments WHERE token = ?", (attachment["token"],)
            ).fetchone()
        finally:
            connection.close()
        self.assertEqual(case["return_invoice"], "8448-2028-2810")
        self.assertEqual(linked["case_id"], case_id)

    def test_creates_explicit_unregistered_case(self) -> None:
        attachment = self.app.save_cs_return_label_attachment("new.png", b"png-data", "image/png")

        result = self.app.apply_cs_return_label_rows(
            [{
                "action": "new",
                "tracking_number": "8448-2000-6446",
                "receiver_name": "국민상회 별하람점",
                "product_name": "에센 대용량 스텐 찜솥 32CM",
                "quantity": "1",
                "condition": "바닥 구멍뚫림",
                "occurred_at": "2026-11-06",
                "attachment_token": attachment["token"],
            }]
        )

        self.assertEqual(result["created"], 1)
        rows = self.app.list_cs_cases(query="국민상회 별하람점", limit=10)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["return_invoice"], "8448-2000-6446")

    def test_rejects_duplicate_return_invoice(self) -> None:
        first = self.create_case("김영란")
        second = self.create_case("이영숙")
        self.app.update_cs_case(first, {"return_invoice": "8448-2028-2810"})

        with self.assertRaisesRegex(ValueError, "이미 등록"):
            self.app.apply_cs_return_label_rows(
                [{"action": "match", "case_id": second, "tracking_number": "8448-2028-2810"}]
            )

    def test_unrecognized_photo_offers_recent_cs_cases_for_manual_selection(self) -> None:
        case_id = self.create_case("배신옥", "벨트리 페이셜 티슈 100매")

        result = self.app.analyze_cs_return_label_uploads(
            [("label.jpg", b"photo-data")],
            ocr_func=lambda _data: {
                "tracking_number": "8448-1920-7540",
                "masked_names": [],
                "quantity": "",
                "condition": "",
                "raw_text": "",
            },
        )

        self.assertEqual(result["rows"][0]["match_status"], "unregistered")
        self.assertEqual(result["rows"][0]["action"], "skip")
        self.assertIn(case_id, [item["id"] for item in result["rows"][0]["candidates"]])


if __name__ == "__main__":
    unittest.main()
