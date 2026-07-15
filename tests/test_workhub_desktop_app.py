from __future__ import annotations

import importlib.util
import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "scripts" / "workhub_vps_desktop_app.py"


def load_desktop_module():
    spec = importlib.util.spec_from_file_location("workhub_vps_desktop_app", MODULE_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class WorkhubDesktopAppTests(unittest.TestCase):
    def test_default_url_uses_vps_domain(self) -> None:
        module = load_desktop_module()

        self.assertEqual(module.DEFAULT_APP_URL, "https://workhub.soilbridgecorp.cloud/")
        with mock.patch.dict(os.environ, {"WORKHUB_DESKTOP_URL": ""}, clear=False):
            self.assertEqual(module.resolve_app_url(None), module.DEFAULT_APP_URL)

    def test_url_can_be_overridden_for_staging(self) -> None:
        module = load_desktop_module()

        with mock.patch.dict(os.environ, {"WORKHUB_DESKTOP_URL": "https://example.test/workhub"}, clear=False):
            self.assertEqual(module.resolve_app_url(None), "https://example.test/workhub")
        self.assertEqual(module.resolve_app_url("http://127.0.0.1:8792/"), "http://127.0.0.1:8792/")

    def test_invalid_url_is_rejected(self) -> None:
        module = load_desktop_module()

        with self.assertRaises(ValueError):
            module.resolve_app_url("file:///C:/workhub.html")

    def test_desktop_download_bridge_saves_files_to_configured_folder(self) -> None:
        module = load_desktop_module()

        with tempfile.TemporaryDirectory() as tempdir:
            with mock.patch.dict(os.environ, {"WORKHUB_DESKTOP_DOWNLOAD_DIR": tempdir}, clear=False):
                api = module.WorkhubDesktopApi("https://workhub.soilbridgecorp.cloud/")
                result = api.saveDownload("report:/bad?.xlsx", "V09SS0hVQg==")

                self.assertTrue(result["ok"])
                saved_path = Path(str(result["path"]))
                self.assertTrue(saved_path.exists())
                self.assertEqual(saved_path.read_bytes(), b"WORKHUB")
                self.assertNotIn(":", saved_path.name)
                self.assertNotIn("?", saved_path.name)

    def test_desktop_download_bridge_saves_chunked_files(self) -> None:
        module = load_desktop_module()

        with tempfile.TemporaryDirectory() as tempdir:
            with mock.patch.dict(os.environ, {"WORKHUB_DESKTOP_DOWNLOAD_DIR": tempdir}, clear=False):
                api = module.WorkhubDesktopApi("https://workhub.soilbridgecorp.cloud/")
                started = api.beginDownload("chunked.xlsx")
                self.assertTrue(started["ok"])
                download_id = str(started["id"])

                self.assertTrue(api.appendDownloadChunk(download_id, "V09SSw==")["ok"])
                self.assertTrue(api.appendDownloadChunk(download_id, "SFVC")["ok"])
                finished = api.finishDownload(download_id)

                self.assertTrue(finished["ok"])
                saved_path = Path(str(finished["path"]))
                self.assertTrue(saved_path.exists())
                self.assertEqual(saved_path.read_bytes(), b"WORKHUB")

    def test_launcher_has_no_browser_fallback(self) -> None:
        source = MODULE_PATH.read_text(encoding="utf-8")

        self.assertNotIn("import webbrowser", source)
        self.assertNotIn("webbrowser.open", source)
        self.assertIn("webview.create_window", source)
        self.assertIn("private_mode=False", source)
        self.assertIn("def saveDownload", source)
        self.assertIn("def beginDownload", source)

    def test_launcher_enables_webview_downloads(self) -> None:
        source = MODULE_PATH.read_text(encoding="utf-8")

        setting = 'webview.settings["ALLOW_DOWNLOADS"] = True'
        self.assertIn(setting, source)
        self.assertLess(source.index(setting), source.index("webview.create_window"))

    def test_build_script_packages_windowed_exe(self) -> None:
        build_script = (ROOT / "build_workhub_desktop_app.ps1").read_text(encoding="utf-8")

        self.assertIn("--onefile", build_script)
        self.assertIn("--windowed", build_script)
        self.assertIn("workhub_vps_desktop_app.py", build_script)
        self.assertIn("SoilbridgeWorkhub.exe", build_script)


if __name__ == "__main__":
    unittest.main()
