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


class VpsDeploymentTests(unittest.TestCase):
    def read(self, relative: str) -> str:
        return (ROOT / relative).read_text(encoding="utf-8")

    def test_vps_deployment_artifacts_are_present_and_point_to_runtime_paths(self) -> None:
        expected_files = [
            ".env.example",
            "README_DEPLOY.md",
            "DEBUG_GUIDE.md",
            "deploy/systemd/workhub.service",
            "deploy/nginx/workhub.conf",
            "deploy/scripts/backup.sh",
            "deploy/scripts/update.sh",
            "deploy/scripts/rollback.sh",
        ]
        for relative in expected_files:
            self.assertTrue((ROOT / relative).exists(), f"{relative} is missing")

        env_example = self.read(".env.example")
        self.assertIn("WORKHUB_ENV=production", env_example)
        self.assertIn("WORKHUB_HOST=127.0.0.1", env_example)
        self.assertIn("WORKHUB_PORT=8770", env_example)
        self.assertIn("WORKHUB_DATA_DIR=/opt/workhub/data", env_example)
        self.assertIn("WORKHUB_BACKUP_DIR=/opt/workhub/backups", env_example)
        self.assertIn("WORKHUB_INITIAL_ADMIN_USERNAME=", env_example)
        self.assertIn("WORKHUB_INITIAL_ADMIN_PASSWORD=", env_example)

        service = self.read("deploy/systemd/workhub.service")
        self.assertIn("EnvironmentFile=/opt/workhub/.env", service)
        self.assertIn("WorkingDirectory=/opt/soillbridgecorp", service)
        self.assertIn("scripts/workhub_delivery_app.py", service)
        self.assertNotIn("_workhub_zip_inspect", service)

        nginx = self.read("deploy/nginx/workhub.conf")
        self.assertIn("server_name erp.soilbridgecorp.cloud", nginx)
        self.assertIn("return 301 https://$host$request_uri", nginx)
        self.assertIn("proxy_pass http://127.0.0.1:8770", nginx)
        self.assertIn("client_max_body_size 50M", nginx)

        for relative in ("deploy/scripts/backup.sh", "deploy/scripts/update.sh", "deploy/scripts/rollback.sh"):
            script = self.read(relative)
            self.assertTrue(script.startswith("#!/usr/bin/env bash"))
            self.assertIn("set -euo pipefail", script)

        dockerfile = self.read("Dockerfile")
        self.assertIn("apt-get install -y --no-install-recommends ca-certificates rclone", dockerfile)

    def test_env_example_is_not_ignored_but_real_env_files_are_ignored(self) -> None:
        gitignore = self.read(".gitignore")
        self.assertIn(".env", gitignore)
        self.assertIn(".env.*", gitignore)
        self.assertIn("!.env.example", gitignore)

    def test_default_password_literals_are_not_embedded_in_source(self) -> None:
        source = self.read("scripts/workhub_delivery_app.py")
        self.assertNotIn("admin1234", source)
        self.assertNotIn("user1234", source)

    def test_production_initial_admin_is_created_from_environment_only(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            previous_env = {
                key: os.environ.get(key)
                for key in (
                    "WORKHUB_DATA_DIR",
                    "WORKHUB_ENV",
                    "WORKHUB_INITIAL_ADMIN_USERNAME",
                    "WORKHUB_INITIAL_ADMIN_NAME",
                    "WORKHUB_INITIAL_ADMIN_PASSWORD",
                )
            }
            try:
                os.environ["WORKHUB_DATA_DIR"] = directory
                os.environ["WORKHUB_ENV"] = "production"
                os.environ["WORKHUB_INITIAL_ADMIN_USERNAME"] = "soilhq"
                os.environ["WORKHUB_INITIAL_ADMIN_NAME"] = "Soillbridge HQ"
                os.environ["WORKHUB_INITIAL_ADMIN_PASSWORD"] = "VpsLaunch!2026"
                sys.modules.pop("workhub_delivery_app", None)
                app = importlib.import_module("workhub_delivery_app")

                app.init_db()

                self.assertEqual(app.authenticate_user("soilhq", "VpsLaunch!2026")["role"], "admin")
                self.assertIsNone(app.authenticate_user("admin", "VpsLaunch!2026"))
                connection = app.connect_db()
                try:
                    usernames = {
                        row["username"]
                        for row in connection.execute("SELECT username FROM users").fetchall()
                    }
                finally:
                    connection.close()
                self.assertEqual(usernames, {"soilhq"})
            finally:
                for key, value in previous_env.items():
                    if value is None:
                        os.environ.pop(key, None)
                    else:
                        os.environ[key] = value
                sys.modules.pop("workhub_delivery_app", None)


    def test_session_cookie_uses_vps_security_environment(self) -> None:
        source = self.read("scripts/workhub_delivery_app.py")

        self.assertIn("WORKHUB_COOKIE_SECURE", source)
        self.assertIn("WORKHUB_COOKIE_SAMESITE", source)
        self.assertIn("WORKHUB_COOKIE_SECURE\n            or isinstance", source)
        self.assertIn("SameSite={WORKHUB_COOKIE_SAMESITE}", source)


if __name__ == "__main__":
    unittest.main()
