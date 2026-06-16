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

    def test_order_workspace_cards_open_existing_modals(self) -> None:
        html_source = (SCRIPTS / "workhub_delivery_app.py").read_text(encoding="utf-8")

        for mode in ("delivery", "invoice", "lotte", "vehicle"):
            self.assertIn(f"{mode}: {{", html_source)

        self.assertIn("ORDER_MODAL_MODES", html_source)
        self.assertIn("ORDER_MODAL_TITLES", html_source)
        self.assertIn("function openOrderModal(mode)", html_source)
        self.assertIn('data-open="order"', html_source)
        self.assertIn('data-order-card="delivery"', html_source)
        self.assertIn('role="button" tabindex="0"', html_source)
        self.assertIn("onclick=\"openOrderModal('delivery')\"", html_source)
        self.assertIn("onclick=\"event.stopPropagation();openOrderModal('delivery');\"", html_source)
        self.assertIn("setActiveNav(\"order\")", html_source)
        self.assertIn("setPageTitle(ORDER_MODAL_TITLES[currentOrderMode]", html_source)
        self.assertIn('sidebar.addEventListener("click"', html_source)
        self.assertIn('event.target.closest("[data-open]")', html_source)
        self.assertIn('event.target.closest("[data-order-card]")', html_source)
        self.assertIn('document.addEventListener("keydown"', html_source)
        self.assertIn("event.stopImmediatePropagation()", html_source)

    def test_delivery_modal_title_matches_menu_label(self) -> None:
        html_source = (SCRIPTS / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('data-open="delivery" onclick="event.stopPropagation();openOrderModal(\'delivery\');">실행</button>', html_source)
        self.assertIn('modalTitle.textContent = "개별 택배건 정리";', html_source)

    def test_order_workspace_has_right_side_execution_cards(self) -> None:
        html_source = (SCRIPTS / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('id="orderWorkspace"', html_source)
        self.assertIn('id="orderWorkspaceTitle"', html_source)
        self.assertIn('id="orderWorkspaceCards"', html_source)
        self.assertIn("order-exec-card", html_source)
        self.assertIn("ORDER_WORKFLOWS", html_source)
        self.assertIn("function showOrderWorkspace(mode)", html_source)
        self.assertIn('orderWorkspace.classList.toggle("active"', html_source)


if __name__ == "__main__":
    unittest.main()
