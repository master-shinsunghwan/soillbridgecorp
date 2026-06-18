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

        for mode in ("delivery", "invoice", "lotte", "salesVendor", "vehicle"):
            self.assertIn(f"{mode}: {{", html_source)

        self.assertIn("ORDER_MODAL_MODES", html_source)
        self.assertIn("ORDER_MODAL_TITLES", html_source)
        self.assertIn("function openOrderModal(mode)", html_source)
        self.assertIn('data-open="order"', html_source)
        self.assertIn('data-order-card="delivery"', html_source)
        self.assertIn('data-order-execute="delivery"', html_source)
        self.assertIn('data-order-card="salesVendor"', html_source)
        self.assertIn('data-order-execute="salesVendor"', html_source)
        self.assertIn('/api/sales-vendor-summary', html_source)
        self.assertIn('data-icon="📦"', html_source)
        self.assertIn('data-icon="🔎"', html_source)
        self.assertIn('data-icon="▦"', html_source)
        self.assertIn('data-icon="🚚"', html_source)
        for mode in ("delivery", "invoice", "lotte", "salesVendor", "vehicle"):
            self.assertNotIn(f'data-order-card="{mode}" role="button" tabindex="0"', html_source)
        self.assertNotIn("onclick=\"openOrderModal('delivery')\"", html_source)
        self.assertNotIn("onclick=\"event.stopPropagation();openOrderModal('delivery');\"", html_source)
        self.assertIn("setActiveNav(\"order\")", html_source)
        self.assertIn("setPageTitle(ORDER_MODAL_TITLES[currentOrderMode]", html_source)
        self.assertIn('sidebar.addEventListener("click"', html_source)
        self.assertIn('event.target.closest("[data-open]")', html_source)
        self.assertIn('document.addEventListener("keydown"', html_source)
        self.assertIn("event.stopImmediatePropagation()", html_source)

    def test_delivery_modal_title_matches_menu_label(self) -> None:
        html_source = (SCRIPTS / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('data-order-execute="delivery">실행</button>', html_source)
        self.assertIn('modalTitle.textContent = "개별 택배건 정리";', html_source)

    def test_order_workspace_has_right_side_execution_cards(self) -> None:
        html_source = (SCRIPTS / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('id="orderWorkspace"', html_source)
        self.assertIn('id="orderWorkspaceTitle"', html_source)
        self.assertIn('id="orderWorkspaceCards"', html_source)
        self.assertIn('id="orderRecentDownloads"', html_source)
        self.assertIn('id="orderDownloadList"', html_source)
        self.assertIn('/api/order-downloads', html_source)
        self.assertIn('/api/order-download?id=', html_source)
        self.assertIn("order-exec-card", html_source)
        self.assertIn("ORDER_WORKFLOWS", html_source)
        self.assertIn("function showOrderWorkspace(mode)", html_source)
        self.assertIn("function loadOrderDownloads()", html_source)
        self.assertIn("function renderOrderDownloads(downloads)", html_source)
        self.assertIn('orderWorkspace.classList.toggle("active"', html_source)

    def test_upload_dialog_does_not_use_daisyui_modal_class(self) -> None:
        for app_file in (
            ROOT / "scripts" / "workhub_delivery_app.py",
            ROOT / "_workhub_zip_inspect" / "scripts" / "workhub_delivery_app.py",
        ):
            html_source = app_file.read_text(encoding="utf-8")

            self.assertIn('class="workhub-modal" role="dialog"', html_source)
            self.assertIn('modal.querySelector(".workhub-modal")', html_source)
            self.assertNotIn('class="modal" role="dialog"', html_source)
            self.assertNotIn(".modal-backdrop.open .modal *", html_source)
            self.assertNotIn(".workhub-modal-backdrop.open .workhub-modal *", html_source)

    def test_naver_mail_defaults_are_managed_from_admin_workspace(self) -> None:
        for app_file in (
            ROOT / "scripts" / "workhub_delivery_app.py",
            ROOT / "_workhub_zip_inspect" / "scripts" / "workhub_delivery_app.py",
        ):
            html_source = app_file.read_text(encoding="utf-8")

            self.assertIn('id="adminNaverEmailInput"', html_source)
            self.assertIn('id="adminNaverPasswordInput"', html_source)
            self.assertIn('id="adminSaveMailCredentials"', html_source)
            self.assertIn('id="adminMailSettingsSave"', html_source)
            self.assertIn("function loadAdminMailSettings()", html_source)
            self.assertIn("function saveAdminMailSettings()", html_source)
            self.assertIn('"/api/mail-settings"', html_source)
            self.assertIn("메일 기본정보 저장", html_source)

            cs_fields_start = html_source.index('class="cs-fields" id="csFields"')
            vendor_select = html_source.index('id="vendorContactSelect"', cs_fields_start)
            cs_mail_account_slice = html_source[cs_fields_start:vendor_select]
            self.assertNotIn('id="naverEmailInput"', cs_mail_account_slice)
            self.assertNotIn('id="naverPasswordInput"', cs_mail_account_slice)
            self.assertNotIn('id="saveMailCredentials"', cs_mail_account_slice)

    def test_vendor_contact_upload_is_managed_from_admin_workspace(self) -> None:
        for app_file in (
            ROOT / "scripts" / "workhub_delivery_app.py",
            ROOT / "_workhub_zip_inspect" / "scripts" / "workhub_delivery_app.py",
        ):
            html_source = app_file.read_text(encoding="utf-8")

            self.assertIn("업체 메일 주소록", html_source)
            self.assertIn('id="vendorContactsFileInput"', html_source)
            self.assertIn('id="vendorContactsDropMain"', html_source)
            self.assertIn("function uploadVendorContactsWorkbook()", html_source)

            admin_start = html_source.index('id="userAdminWorkspace"')
            admin_end = html_source.index('id="userAdminBody"', admin_start)
            admin_slice = html_source[admin_start:admin_end]
            self.assertIn('id="vendorContactsFileInput"', admin_slice)
            self.assertIn('id="vendorContactsDropMain"', admin_slice)

            cs_fields_start = html_source.index('class="cs-fields" id="csFields"')
            recipient_field = html_source.index('id="recipientEmailInput"', cs_fields_start)
            cs_contact_slice = html_source[cs_fields_start:recipient_field]
            self.assertNotIn('id="vendorContactsFileInput"', cs_contact_slice)
            self.assertNotIn('id="vendorContactsDropMain"', cs_contact_slice)

    def test_excel_downloads_keep_object_url_until_browser_starts_download(self) -> None:
        for app_file in (
            ROOT / "scripts" / "workhub_delivery_app.py",
            ROOT / "_workhub_zip_inspect" / "scripts" / "workhub_delivery_app.py",
        ):
            html_source = app_file.read_text(encoding="utf-8")

            self.assertIn("function downloadWorkbookResponse(response, fallbackName)", html_source)
            self.assertIn("window.setTimeout(() => URL.revokeObjectURL(url), 1000)", html_source)
            self.assertIn('await downloadWorkbookResponse(response, "차량인수증.xlsx")', html_source)
            self.assertIn(
                'await downloadWorkbookResponse(\n            response,\n            currentMode === "invoice" ? "송장번호_추출.xlsx" : "롯데택배_발주서.xlsx"\n          )',
                html_source,
            )


if __name__ == "__main__":
    unittest.main()
