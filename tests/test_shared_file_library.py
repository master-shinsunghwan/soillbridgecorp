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


class SharedFileLibraryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(self.tempdir.cleanup)
        os.environ["WORKHUB_DATA_DIR"] = self.tempdir.name
        sys.modules.pop("workhub_delivery_app", None)
        self.app = importlib.import_module("workhub_delivery_app")
        self.app.init_db()

    def test_save_list_download_and_delete_shared_file(self) -> None:
        source = Path(self.tempdir.name) / "sample.txt"
        source.write_text("workhub helper", encoding="utf-8")

        saved = self.app.save_shared_file(source, "../업무 안내.txt", "admin")
        files = self.app.list_shared_files()

        self.assertEqual(len(files), 1)
        self.assertEqual(files[0]["id"], saved["id"])
        self.assertEqual(files[0]["original_name"], "업무 안내.txt")
        self.assertEqual(files[0]["uploaded_by"], "admin")
        self.assertGreater(files[0]["size"], 0)

        path, metadata = self.app.shared_file_download_info(saved["id"])
        self.assertTrue(path.is_file())
        self.assertEqual(path.read_text(encoding="utf-8"), "workhub helper")
        self.assertEqual(metadata["original_name"], "업무 안내.txt")

        self.app.delete_shared_file(saved["id"])
        self.assertEqual(self.app.list_shared_files(), [])
        self.assertFalse(path.exists())

    def test_shared_file_upload_helper_accepts_general_work_files(self) -> None:
        saved_path = self.app.save_uploaded_shared_file(
            {"file": ("업무 안내.txt", "plain text".encode("utf-8"))},
            "file",
        )

        self.assertTrue(saved_path.is_file())
        self.assertEqual(saved_path.read_text(encoding="utf-8"), "plain text")
        self.assertTrue(saved_path.name.endswith("_업무 안내.txt"))


if __name__ == "__main__":
    unittest.main()
