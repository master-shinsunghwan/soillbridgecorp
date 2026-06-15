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


class WorkhubAppFeatureParityTests(unittest.TestCase):
    def load_app(self):
        for module_name in ("workhub_delivery_app", "workhub_crm"):
            sys.modules.pop(module_name, None)
        tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(tempdir.cleanup)
        os.environ["WORKHUB_DATA_DIR"] = tempdir.name
        return importlib.import_module("workhub_delivery_app")

    def test_root_app_exposes_crm_feature_set_from_packaged_app(self) -> None:
        crm = importlib.import_module("workhub_crm")
        app = self.load_app()

        permission_keys = {key for key, _, _ in app.PERMISSION_DEFINITIONS}

        self.assertTrue(hasattr(crm, "crm_dashboard_payload"))
        self.assertIn("crm_view", permission_keys)
        self.assertIn("crm_manage", permission_keys)
        self.assertIn("crm_message_manage", permission_keys)
        self.assertIn("crm_view", app.DEFAULT_ROLE_PERMISSIONS["user"])
        self.assertIn("sub_admin", app.DEFAULT_ROLE_PERMISSIONS)
        self.assertIn('data-open="crm"', app.render_app_html({"display_name": "???", "role": "admin", "permissions": app.ALL_PERMISSIONS}))
