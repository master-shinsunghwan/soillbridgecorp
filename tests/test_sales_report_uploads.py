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


if __name__ == "__main__":
    unittest.main()
