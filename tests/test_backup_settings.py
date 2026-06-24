from __future__ import annotations

import importlib
import os
import sys
import tempfile
import zipfile
from pathlib import Path
from types import SimpleNamespace


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


def load_app(tmp_path: Path):
    os.environ["WORKHUB_DATA_DIR"] = str(tmp_path)
    sys.modules.pop("workhub_delivery_app", None)
    return importlib.import_module("workhub_delivery_app")


def test_backup_settings_control_auto_and_default_backup_dir(tmp_path: Path) -> None:
    app = load_app(tmp_path)
    configured_dir = tmp_path / "google-drive-style-backups"

    settings = app.save_backup_settings(
        {
            "backup_dir": str(configured_dir),
            "auto_enabled": False,
            "auto_hour": 2,
            "retention_days": 30,
        }
    )
    backup = app.create_workhub_backup("manual")

    assert settings["backup_dir"] == str(configured_dir)
    assert settings["auto_enabled"] is False
    assert settings["auto_hour"] == 2
    assert settings["retention_days"] == 30
    assert (configured_dir / backup["name"]).exists()
    assert app.list_backup_files()[0]["name"] == backup["name"]


def test_create_backup_can_use_one_time_target_dir_without_changing_settings(tmp_path: Path) -> None:
    app = load_app(tmp_path)
    configured_dir = tmp_path / "configured"
    selected_dir = tmp_path / "selected-once"
    app.save_backup_settings({"backup_dir": str(configured_dir)})

    backup = app.create_workhub_backup("selected", backup_dir=selected_dir)

    assert (selected_dir / backup["name"]).exists()
    assert not (configured_dir / backup["name"]).exists()
    assert app.load_backup_settings()["backup_dir"] == str(configured_dir)


def test_create_backup_includes_runtime_work_files_for_full_restore(tmp_path: Path) -> None:
    app = load_app(tmp_path)
    configured_dir = tmp_path / "backups"
    app.save_backup_settings({"backup_dir": str(configured_dir)})
    (app.UPLOAD_DIR / "cs").mkdir(parents=True, exist_ok=True)
    (app.UPLOAD_DIR / "cs" / "photo.png").write_bytes(b"image")
    app.SHARED_FILE_DIR.mkdir(parents=True, exist_ok=True)
    (app.SHARED_FILE_DIR / "guide.xlsx").write_bytes(b"shared")
    app.SALES_REPORT_DIR.mkdir(parents=True, exist_ok=True)
    (app.SALES_REPORT_DIR / "sales.xlsx").write_bytes(b"sales")

    backup = app.create_workhub_backup("manual")

    with zipfile.ZipFile(configured_dir / backup["name"]) as archive:
        names = set(archive.namelist())
    assert "config/workhub.db" in names
    assert "output/uploads/cs/photo.png" in names
    assert "shared_files/guide.xlsx" in names
    assert "sales_reports/sales.xlsx" in names


def test_offline_backup_can_skip_external_upload(tmp_path: Path) -> None:
    app = load_app(tmp_path)
    configured_dir = tmp_path / "backups"
    app.save_backup_settings(
        {
            "backup_dir": str(configured_dir),
            "external_enabled": True,
            "rclone_remote": "workhub-gdrive",
            "rclone_path": "WorkhubBackups",
        }
    )

    backup = app.create_workhub_backup("offline-download", upload_external=False)

    assert (configured_dir / backup["name"]).exists()
    assert backup["external_backup"]["status"] == "disabled"


def test_rclone_external_backup_upload_builds_google_drive_copy_command(tmp_path: Path) -> None:
    app = load_app(tmp_path)
    backup_path = tmp_path / "workhub_backup_20260619_120000.zip"
    backup_path.parent.mkdir(parents=True, exist_ok=True)
    backup_path.write_bytes(b"backup")
    calls = []

    def fake_runner(command, **kwargs):
        calls.append((command, kwargs))
        return SimpleNamespace(returncode=0, stdout="copied", stderr="")

    settings = app.save_backup_settings(
        {
            "external_enabled": True,
            "external_type": "rclone",
            "rclone_remote": "gdrive",
            "rclone_path": "Soillbridge/Workhub_Backup",
            "rclone_executable": "rclone",
        }
    )

    result = app.upload_backup_to_external_storage(backup_path, settings, runner=fake_runner)

    assert result["status"] == "success"
    assert result["target"] == "gdrive:Soillbridge/Workhub_Backup"
    assert calls[0][0] == ["rclone", "copy", str(backup_path), "gdrive:Soillbridge/Workhub_Backup"]


def test_external_backup_upload_failure_is_reported_without_deleting_local_backup(tmp_path: Path) -> None:
    app = load_app(tmp_path)
    backup_path = tmp_path / "workhub_backup_20260619_120000.zip"
    backup_path.parent.mkdir(parents=True, exist_ok=True)
    backup_path.write_bytes(b"backup")

    def fake_runner(command, **kwargs):
        return SimpleNamespace(returncode=1, stdout="", stderr="remote not found")

    settings = app.save_backup_settings(
        {
            "external_enabled": True,
            "external_type": "rclone",
            "rclone_remote": "gdrive",
            "rclone_path": "Soillbridge/Workhub_Backup",
        }
    )

    result = app.upload_backup_to_external_storage(backup_path, settings, runner=fake_runner)

    assert result["status"] == "fail"
    assert "remote not found" in result["message"]
    assert backup_path.exists()


def main() -> None:
    with tempfile.TemporaryDirectory() as directory:
        test_backup_settings_control_auto_and_default_backup_dir(Path(directory) / "settings")
    with tempfile.TemporaryDirectory() as directory:
        test_create_backup_can_use_one_time_target_dir_without_changing_settings(Path(directory) / "selected")
    with tempfile.TemporaryDirectory() as directory:
        test_create_backup_includes_runtime_work_files_for_full_restore(Path(directory) / "full")
    with tempfile.TemporaryDirectory() as directory:
        test_offline_backup_can_skip_external_upload(Path(directory) / "offline")
    with tempfile.TemporaryDirectory() as directory:
        test_rclone_external_backup_upload_builds_google_drive_copy_command(Path(directory) / "rclone")
    with tempfile.TemporaryDirectory() as directory:
        test_external_backup_upload_failure_is_reported_without_deleting_local_backup(Path(directory) / "rclone-fail")


if __name__ == "__main__":
    main()
