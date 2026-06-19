from __future__ import annotations

import importlib
import os
import sys
import tempfile
from pathlib import Path


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


def main() -> None:
    with tempfile.TemporaryDirectory() as directory:
        test_backup_settings_control_auto_and_default_backup_dir(Path(directory) / "settings")
    with tempfile.TemporaryDirectory() as directory:
        test_create_backup_can_use_one_time_target_dir_without_changing_settings(Path(directory) / "selected")


if __name__ == "__main__":
    main()
