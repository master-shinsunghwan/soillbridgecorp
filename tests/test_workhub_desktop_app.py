from __future__ import annotations

import hashlib
import importlib.util
import io
import json
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "scripts" / "workhub_vps_desktop_app.py"


class FakeUrlResponse(io.BytesIO):
    def __init__(self, data: bytes, url: str):
        super().__init__(data)
        self._url = url

    def geturl(self) -> str:
        return self._url

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()
        return False


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

    def test_desktop_update_versions_are_compared_numerically(self) -> None:
        module = load_desktop_module()

        self.assertTrue(module.is_newer_version("1.1.1", "1.1.0"))
        self.assertTrue(module.is_newer_version("1.10.0", "1.9.9"))
        self.assertFalse(module.is_newer_version("1.1.0", "1.1.0"))
        self.assertFalse(module.is_newer_version("1.0.9", "1.1.0"))
        with self.assertRaises(ValueError):
            module.version_parts("latest")

    def test_desktop_update_is_downloaded_verified_and_staged(self) -> None:
        module = load_desktop_module()
        executable = b"MZ" + (b"WORKHUB-UPDATE" * 32)
        digest = hashlib.sha256(executable).hexdigest()
        manifest = {
            "version": "1.1.1",
            "url": "https://github.com/example/releases/download/desktop-v1.1.1/SoilbridgeWorkhub.exe",
            "sha256": digest,
            "size": len(executable),
            "published_at": "2026-07-16T00:00:00Z",
            "notes": "test update",
        }

        with tempfile.TemporaryDirectory() as tempdir:
            root = Path(tempdir)
            module.DESKTOP_APP_DIR = root
            module.DESKTOP_UPDATE_DIR = root / "Updates"
            module.PENDING_UPDATE_PATH = module.DESKTOP_UPDATE_DIR / "pending_update.json"
            responses = [
                FakeUrlResponse(json.dumps(manifest).encode("utf-8"), "https://raw.githubusercontent.com/example/manifest.json"),
                FakeUrlResponse(executable, "https://release-assets.githubusercontent.com/example/asset"),
            ]
            with mock.patch.object(module, "urlopen", side_effect=responses):
                result = module.stage_available_desktop_update("https://raw.githubusercontent.com/example/manifest.json")

            self.assertTrue(result["ok"])
            self.assertTrue(result["staged"])
            staged_path = Path(str(result["path"]))
            self.assertEqual(staged_path.read_bytes(), executable)
            self.assertEqual(module.read_pending_desktop_update()["sha256"], digest)

    def test_desktop_update_rejects_tampered_executable(self) -> None:
        module = load_desktop_module()
        expected = b"MZ-valid"
        tampered = b"MZ-faked"
        manifest = {
            "version": "1.1.1",
            "url": "https://github.com/example/SoilbridgeWorkhub.exe",
            "sha256": hashlib.sha256(expected).hexdigest(),
            "size": len(expected),
        }

        with tempfile.TemporaryDirectory() as tempdir:
            root = Path(tempdir)
            module.DESKTOP_APP_DIR = root
            module.DESKTOP_UPDATE_DIR = root / "Updates"
            module.PENDING_UPDATE_PATH = module.DESKTOP_UPDATE_DIR / "pending_update.json"
            response = FakeUrlResponse(tampered, "https://release-assets.githubusercontent.com/example/asset")
            with mock.patch.object(module, "urlopen", return_value=response):
                with self.assertRaisesRegex(ValueError, "검증값"):
                    module.download_desktop_update(module.validate_desktop_update_manifest(manifest))

            self.assertFalse(list(module.DESKTOP_UPDATE_DIR.glob("*.part")))

    def test_pending_update_launches_hidden_verified_replacement(self) -> None:
        module = load_desktop_module()
        update_bytes = b"MZ" + (b"AUTO-UPDATE" * 16)
        digest = hashlib.sha256(update_bytes).hexdigest()

        with tempfile.TemporaryDirectory() as tempdir:
            root = Path(tempdir)
            update_dir = root / "Updates"
            update_dir.mkdir()
            staged_path = update_dir / "SoilbridgeWorkhub-1.1.1.exe"
            staged_path.write_bytes(update_bytes)
            target_path = root / "SoilbridgeWorkhub.exe"
            target_path.write_bytes(b"old")
            module.DESKTOP_APP_DIR = root
            module.DESKTOP_UPDATE_DIR = update_dir
            module.PENDING_UPDATE_PATH = update_dir / "pending_update.json"
            module.write_json_atomic(
                module.PENDING_UPDATE_PATH,
                {
                    "version": "1.1.1",
                    "url": "https://github.com/example/SoilbridgeWorkhub.exe",
                    "sha256": digest,
                    "size": len(update_bytes),
                    "path": str(staged_path),
                },
            )

            with mock.patch.dict(os.environ, {"WORKHUB_DESKTOP_DISABLE_UPDATES": ""}, clear=False):
                with mock.patch.object(module.sys, "frozen", True, create=True):
                    with mock.patch.object(module.sys, "executable", str(target_path)):
                        with mock.patch.object(module.subprocess, "Popen") as popen:
                            self.assertTrue(module.launch_pending_desktop_update())

            popen.assert_called_once()
            script_path = next(update_dir.glob("apply_update_*.ps1"))
            script = script_path.read_text(encoding="utf-8-sig")
            self.assertIn("Wait-Process -Id $ParentPid", script)
            self.assertIn("Get-FileHash -Algorithm SHA256", script)
            self.assertIn("$Target.update-backup", script)
            self.assertIn("Start-Process -FilePath $Target", script)

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

    def test_custom_download_folder_is_persisted_and_used(self) -> None:
        module = load_desktop_module()

        with tempfile.TemporaryDirectory() as tempdir:
            root = Path(tempdir)
            module.DESKTOP_SETTINGS_PATH = root / "settings.json"
            module.DEFAULT_DOWNLOAD_DIR = root / "Downloads"
            custom_dir = root / "TeamExports"
            with mock.patch.dict(os.environ, {"WORKHUB_DESKTOP_DOWNLOAD_DIR": ""}, clear=False):
                module.set_desktop_download_dir(custom_dir)
                api = module.WorkhubDesktopApi("https://workhub.soilbridgecorp.cloud/")
                result = api.saveDownload("management.xlsx", "V09SS0hVQg==")

            self.assertTrue(result["ok"])
            self.assertEqual(Path(str(result["path"])).parent, custom_dir)
            self.assertEqual(module.read_desktop_settings()["download_dir"], str(custom_dir))
            self.assertEqual(module.desktop_download_settings()["source"], "custom")

    def test_download_folder_picker_persists_selected_folder(self) -> None:
        module = load_desktop_module()

        with tempfile.TemporaryDirectory() as tempdir:
            root = Path(tempdir)
            custom_dir = root / "Selected"
            custom_dir.mkdir()
            module.DESKTOP_SETTINGS_PATH = root / "settings.json"
            module.DEFAULT_DOWNLOAD_DIR = root / "Downloads"
            module._window = mock.Mock()
            module._window.create_file_dialog.return_value = (str(custom_dir),)
            fake_webview = types.SimpleNamespace(FOLDER_DIALOG=20)
            with mock.patch.dict(os.environ, {"WORKHUB_DESKTOP_DOWNLOAD_DIR": ""}, clear=False):
                with mock.patch.dict(sys.modules, {"webview": fake_webview}):
                    result = module.WorkhubDesktopApi(module.DEFAULT_APP_URL).chooseDownloadFolder()

            self.assertTrue(result["ok"])
            self.assertTrue(result["customized"])
            self.assertEqual(result["path"], str(custom_dir))
            module._window.create_file_dialog.assert_called_once_with(
                fake_webview.FOLDER_DIALOG,
                directory=str(module.DEFAULT_DOWNLOAD_DIR),
            )

    def test_reset_download_folder_restores_default(self) -> None:
        module = load_desktop_module()

        with tempfile.TemporaryDirectory() as tempdir:
            root = Path(tempdir)
            module.DESKTOP_SETTINGS_PATH = root / "settings.json"
            module.DEFAULT_DOWNLOAD_DIR = root / "Downloads"
            with mock.patch.dict(os.environ, {"WORKHUB_DESKTOP_DOWNLOAD_DIR": ""}, clear=False):
                module.set_desktop_download_dir(root / "Custom")
                result = module.WorkhubDesktopApi(module.DEFAULT_APP_URL).resetDownloadFolder()

            self.assertTrue(result["ok"])
            self.assertFalse(result["customized"])
            self.assertEqual(result["path"], str(module.DEFAULT_DOWNLOAD_DIR))
            self.assertNotIn("download_dir", module.read_desktop_settings())

    def test_launcher_has_no_browser_fallback(self) -> None:
        source = MODULE_PATH.read_text(encoding="utf-8")

        self.assertNotIn("import webbrowser", source)
        self.assertNotIn("webbrowser.open", source)
        self.assertIn("webview.create_window", source)
        self.assertIn("private_mode=False", source)
        self.assertIn("def saveDownload", source)
        self.assertIn("def beginDownload", source)
        self.assertIn("def chooseDownloadFolder", source)
        self.assertIn("def resetDownloadFolder", source)

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
        self.assertIn("Copy-PowerShellScriptWithBom", build_script)
        self.assertIn("Text.UTF8Encoding($true)", build_script)

    def test_desktop_update_publisher_builds_release_and_manifest(self) -> None:
        publish_script = (ROOT / "publish_workhub_desktop_update.ps1").read_text(encoding="utf-8")

        self.assertIn("build_workhub_desktop_app.ps1", publish_script)
        self.assertIn("gh release create", publish_script)
        self.assertIn("gh release upload", publish_script)
        self.assertIn("Get-FileHash -Algorithm SHA256", publish_script)
        self.assertIn("static\\desktop_update.json", publish_script)

    def test_installer_avoids_fragile_vbs_quote_generation(self) -> None:
        install_script = (ROOT / "install_workhub_desktop_app.ps1").read_text(encoding="utf-8")

        self.assertNotIn("$StartupCommand", install_script)
        self.assertNotIn("CreateObject(\"WScript.Shell\")", install_script)
        self.assertIn("New-WorkhubDesktopShortcut", install_script)

    def test_installer_stops_running_app_before_replacing_exe(self) -> None:
        install_script = (ROOT / "install_workhub_desktop_app.ps1").read_text(encoding="utf-8")

        stop_call = "Stop-RunningWorkhub"
        copy_call = "Copy-Item -LiteralPath $ExeSource -Destination $ExeTarget -Force"
        self.assertIn('Get-Process -Name "SoilbridgeWorkhub"', install_script)
        self.assertIn("Stop-Process -Force", install_script)
        self.assertIn("Start-Sleep -Milliseconds 300", install_script)
        self.assertLess(install_script.rindex(stop_call), install_script.index(copy_call))


if __name__ == "__main__":
    unittest.main()
