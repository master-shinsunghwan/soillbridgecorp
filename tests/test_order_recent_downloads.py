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


class OrderRecentDownloadTests(unittest.TestCase):
    def load_app(self):
        for module_name in ("workhub_delivery_app", "workhub_crm"):
            sys.modules.pop(module_name, None)
        tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(tempdir.cleanup)
        os.environ["WORKHUB_DATA_DIR"] = tempdir.name
        return importlib.import_module("workhub_delivery_app")

    def test_order_download_history_keeps_latest_ten_files(self) -> None:
        app = self.load_app()
        app.DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

        for index in range(12):
            output_path = app.DOWNLOAD_DIR / f"order-output-{index}.xlsx"
            output_path.write_bytes(f"file-{index}".encode("utf-8"))
            app.register_order_download(output_path, f"작업 {index}")

        downloads = app.list_order_downloads()

        self.assertEqual(len(downloads), 10)
        self.assertEqual(downloads[0]["workflow"], "작업 11")
        self.assertEqual(downloads[-1]["workflow"], "작업 2")
        self.assertNotIn("작업 1", {item["workflow"] for item in downloads})
        self.assertTrue(app.order_download_path(downloads[0]["id"]).exists())
        with self.assertRaises(FileNotFoundError):
            app.order_download_path("missing")


if __name__ == "__main__":
    unittest.main()
