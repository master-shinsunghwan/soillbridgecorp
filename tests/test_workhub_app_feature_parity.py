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

    def test_dashboard_entry_shows_notice_import_schedule_and_calendar(self) -> None:
        for app_file in (
            ROOT / "scripts" / "workhub_delivery_app.py",
            ROOT / "_workhub_zip_inspect" / "scripts" / "workhub_delivery_app.py",
        ):
            html_source = app_file.read_text(encoding="utf-8")

            self.assertIn('id="sidebarNoticePreview"', html_source)
            self.assertIn('id="dashboardImportScheduleCard"', html_source)
            self.assertIn('data-company-panel="calendar"', html_source)
            self.assertIn('class="company-panel active dashboard-calendar-panel" data-company-panel="calendar"', html_source)
            notice_index = html_source.index('id="sidebarNoticePreview"')
            import_index = html_source.index('id="dashboardImportScheduleCard"')
            calendar_index = html_source.index('data-company-panel="calendar"')
            self.assertLess(notice_index, import_index)
            self.assertLess(import_index, calendar_index)
            self.assertIn("수입제품 입고 일정", html_source)
            self.assertIn('id="dashboardImportScheduleSummary"', html_source)
            self.assertIn('id="dashboardImportScheduleBody"', html_source)
            self.assertIn('class="company-card dashboard-calendar-card"', html_source)
            self.assertIn('id="dashboardSalesPanel"', html_source)
            self.assertIn("매출 현황", html_source)
            sales_panel_start = html_source.index('id="dashboardSalesPanel"')
            sales_panel_end = html_source.index('</aside>', sales_panel_start)
            sales_panel_html = html_source[sales_panel_start:sales_panel_end]
            self.assertIn("발주모아 매출 데이터 연결 대기 중", sales_panel_html)
            self.assertNotIn("선택한 날짜", sales_panel_html)
            self.assertNotIn("이번 달 요약", sales_panel_html)
            self.assertIn(".dashboard-calendar-panel", html_source)
            self.assertIn("grid-template-columns: minmax(0, 1fr) minmax(0, 1fr)", html_source)
            self.assertIn("function renderDashboardImportSchedule()", html_source)
            self.assertIn("renderDashboardImportSchedule();", html_source)
            self.assertIn('companyActiveTab === "notice" && panel.dataset.companyPanel === "calendar"', html_source)
            self.assertIn("loadDashboardEntryData().catch", html_source)

    def test_sidebar_navigation_uses_hierarchical_text_colors(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        for color_token in (
            "--sidebar-text-primary",
            "--sidebar-text-secondary",
            "--sidebar-text-muted",
            "--sidebar-text-subtle",
            "--sidebar-accent",
        ):
            self.assertIn(color_token, html_source)
        self.assertIn(".nav-item .nav-label span", html_source)
        self.assertIn(".nav-item svg", html_source)
        self.assertIn(".nav-subitem.active", html_source)
        self.assertIn("color: var(--sidebar-text-primary)", html_source)
        self.assertIn("color: var(--sidebar-text-secondary)", html_source)
        self.assertIn("color: var(--sidebar-text-muted)", html_source)

    def test_company_portal_click_always_opens_submenu_after_workspace_switch(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn("const companyGroup = document.querySelector(\"#companyNavGroup\")", html_source)
        self.assertIn("companyGroup?.classList.add(\"open\")", html_source)
        self.assertNotIn("wasCompanyGroupOpen", html_source)
        self.assertNotIn('companyGroup?.classList.toggle("open"', html_source)
        self.assertNotIn('document.querySelector("#companyNavGroup").classList.toggle("open");', html_source)

    def test_daily_ledger_uploads_live_under_each_ledger_sidebar_group(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        management_group_start = html_source.index('id="managementNavGroup"')
        ledger_group_start = html_source.index('id="ledgerNavGroup"')
        crm_group_start = html_source.index('id="crmNavGroup"')
        admin_group_start = html_source.index('id="adminNavGroup"')
        admin_group_end = html_source.index('id="leaveWorkspace"', admin_group_start)
        management_group = html_source[management_group_start:ledger_group_start]
        ledger_group = html_source[ledger_group_start:crm_group_start]
        admin_group = html_source[admin_group_start:admin_group_end]

        self.assertIn('data-open="management"', management_group)
        self.assertIn('id="managementImportOpen"', management_group)
        self.assertIn('data-management-import-mode="daily"', management_group)
        self.assertIn("통합관리대장 일일 추가 업로드", management_group)

        self.assertIn('data-open="ledger"', ledger_group)
        self.assertIn('id="ledgerImportOpen"', ledger_group)
        self.assertIn('data-ledger-import-mode="daily"', ledger_group)
        self.assertIn("CS처리대장 일일 추가 업로드", ledger_group)

        self.assertIn('data-management-import-mode="replace"', admin_group)
        self.assertIn('data-ledger-import-mode="replace"', admin_group)
        self.assertNotIn('data-management-import-mode="daily"', admin_group)
        self.assertNotIn('data-ledger-import-mode="daily"', admin_group)
        self.assertIn('#managementNavToggle', html_source)
        self.assertIn('#ledgerNavToggle', html_source)

    def test_cs_request_from_ledger_menu_supports_mail_attachments(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        ledger_group_start = html_source.index('id="ledgerNavGroup"')
        crm_group_start = html_source.index('id="crmNavGroup"')
        ledger_group = html_source[ledger_group_start:crm_group_start]

        self.assertIn('data-mail-popup="cs"', ledger_group)
        self.assertIn("CS처리 요청", ledger_group)
        self.assertIn('id="csAttachmentInput"', html_source)
        self.assertIn('id="csAttachmentSummary"', html_source)
        self.assertIn('accept="image/*,video/*"', html_source)
        self.assertIn("appendCsMailPayload", html_source)
        self.assertIn("collect_mail_attachments", html_source)
        self.assertIn("attachments=attachments", html_source)
        self.assertIn('document.querySelectorAll("[data-mail-popup]")', html_source)

    def test_crm_task_board_uses_collapsible_advanced_filters_from_branch(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('id="crmTaskAdvancedToggle"', html_source)
        self.assertIn('id="crmAdvancedFilters" hidden', html_source)
        self.assertIn('aria-controls="crmAdvancedFilters"', html_source)
        self.assertIn(".crm-advanced-filters", html_source)
        self.assertIn("crmTaskAdvancedToggle?.addEventListener", html_source)
        self.assertIn('crmTaskAdvancedToggle.textContent = open ? "필터 닫기" : "고급 필터"', html_source)
        self.assertIn('data-crm-nav-tab="dashboard">업무 현황</button>', html_source)
        self.assertIn('data-crm-tab="accounts">직원 현황</button>', html_source)
        self.assertIn('data-crm-tab="messages">연동 로그</button>', html_source)

    def test_leave_workflow_has_multi_step_approval_cancel_and_accrual_ui(self) -> None:
        for app_file in (
            ROOT / "scripts" / "workhub_delivery_app.py",
            ROOT / "_workhub_zip_inspect" / "scripts" / "workhub_delivery_app.py",
        ):
            html_source = app_file.read_text(encoding="utf-8")

            for permission in (
                "leave_approve_team",
                "leave_approve_director",
                "leave_approve_ceo",
                "leave_director_override",
            ):
                self.assertIn(permission, html_source)
            for table_name in ("company_holidays", "leave_notifications"):
                self.assertIn(table_name, html_source)
            for endpoint in ("/api/leave-cancel", "/api/leave-accrual", "/api/company-holiday"):
                self.assertIn(endpoint, html_source)
            for function_name in (
                "save_company_holiday",
                "apply_annual_leave_accrual",
                "cancel_leave_request",
                "list_leave_notifications",
                "actor_can_override_leave",
            ):
                self.assertIn(f"def {function_name}", html_source)
            self.assertIn('id="leaveReservedDays"', html_source)
            self.assertIn('id="leaveNotificationList"', html_source)
            self.assertIn('id="leaveAccrualApply"', html_source)
            self.assertIn('id="leaveHolidayDateInput"', html_source)
            self.assertIn('id="leaveHolidaySave"', html_source)
            self.assertIn('data-leave-comment="${row.id}"', html_source)
            self.assertIn('data-leave-cancel="${row.id}"', html_source)
            self.assertIn('data-leave-decision="override"', html_source)

    def test_delivery_modal_title_matches_menu_label(self) -> None:
        html_source = (SCRIPTS / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('data-order-execute="delivery">실행</button>', html_source)
        self.assertIn('modalTitle.textContent = "개별 택배건 정리";', html_source)

    def test_delivery_summary_supports_safe_number_confirmation(self) -> None:
        for app_file in (
            ROOT / "scripts" / "workhub_delivery_app.py",
            ROOT / "_workhub_zip_inspect" / "scripts" / "workhub_delivery_app.py",
        ):
            html_source = app_file.read_text(encoding="utf-8")

            self.assertIn("build_summary_payload", html_source)
            self.assertIn("safeNumberConfirmMessage", html_source)
            self.assertIn("safe_number_candidates", html_source)
            self.assertIn("approved_text", html_source)
            self.assertIn("안심번호 합포 후보", html_source)
            self.assertIn('id="safeNumberPackageDialog"', html_source)
            self.assertIn("function requestSafeNumberPackageApproval", html_source)
            self.assertIn('id="safeNumberPackageApprove"', html_source)
            self.assertIn('id="safeNumberPackageReject"', html_source)
            self.assertNotIn("window.confirm(safeNumberConfirmMessage", html_source)

    def test_ledger_imports_are_split_into_daily_and_replace_modes(self) -> None:
        html_source = (SCRIPTS / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn("일일 추가 업로드", html_source)
        self.assertIn("전체 데이터 교체 업로드", html_source)
        self.assertIn("/api/management-import-preview", html_source)
        self.assertIn("/api/cs-cases-import-preview", html_source)
        self.assertIn("function requestImportWarningApproval", html_source)
        self.assertIn('id="importWarningDialog"', html_source)
        self.assertIn('id="importCorrectionDialog"', html_source)
        self.assertIn("function requestImportCorrectionApproval", html_source)
        self.assertIn("data-correction-field", html_source)
        self.assertIn('formData.append("corrections"', html_source)
        self.assertIn('mode", mode', html_source)
        self.assertIn('mode === "replace"', html_source)
        self.assertIn("preview_management_import", html_source)
        self.assertIn("preview_cs_cases_import", html_source)

    def test_backup_workspace_supports_auto_and_selected_backup_settings(self) -> None:
        html_source = (SCRIPTS / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('id="backupAutoEnabled"', html_source)
        self.assertIn('id="backupAutoHour"', html_source)
        self.assertIn('id="backupRetentionDays"', html_source)
        self.assertIn('id="backupDirInput"', html_source)
        self.assertIn('id="backupSettingsSave"', html_source)
        self.assertIn('id="backupCreateSelected"', html_source)
        self.assertIn('id="backupExternalEnabled"', html_source)
        self.assertIn('id="backupRcloneRemote"', html_source)
        self.assertIn('id="backupRclonePath"', html_source)
        self.assertIn('id="backupRcloneExecutable"', html_source)
        self.assertIn('id="backupExternalStatus"', html_source)
        self.assertIn("/api/backup-settings", html_source)
        self.assertIn("function saveBackupSettings", html_source)
        self.assertIn("function createBackupAtSelectedPath", html_source)
        self.assertIn("Google Drive 업로드 상태", html_source)
        self.assertIn("rclone_remote", html_source)
        self.assertIn("load_backup_settings", html_source)
        self.assertIn("save_backup_settings", html_source)

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

    def test_bulk_mail_technical_settings_are_managed_from_admin_workspace(self) -> None:
        for app_file in (
            ROOT / "scripts" / "workhub_delivery_app.py",
            ROOT / "_workhub_zip_inspect" / "scripts" / "workhub_delivery_app.py",
        ):
            html_source = app_file.read_text(encoding="utf-8")

            self.assertIn('id="adminSmtpPort"', html_source)
            self.assertIn('id="adminSmtpSecurity"', html_source)
            self.assertIn('id="adminBulkBatchSize"', html_source)
            self.assertIn('id="adminBulkSendInterval"', html_source)
            self.assertIn('id="adminBulkBatchPause"', html_source)
            self.assertIn('id="adminBulkTestRecipient"', html_source)
            self.assertIn('id="adminMailTechnicalSave"', html_source)
            self.assertIn('id="adminMailTestSend"', html_source)
            self.assertIn("function saveAdminMailTechnicalSettings()", html_source)
            self.assertIn("function sendAdminMailTestMessage()", html_source)
            self.assertIn('"/api/mail-test"', html_source)

    def test_naver_mail_integration_admin_ui_uses_clear_operating_text(self) -> None:
        for app_file in (
            ROOT / "scripts" / "workhub_delivery_app.py",
            ROOT / "_workhub_zip_inspect" / "scripts" / "workhub_delivery_app.py",
        ):
            html_source = app_file.read_text(encoding="utf-8")

            self.assertIn("네이버 메일 연동", html_source)
            self.assertIn("네이버 메일 아이디", html_source)
            self.assertIn("네이버 메일 비밀번호", html_source)
            self.assertIn("연동 테스트 메일 발송", html_source)
            self.assertIn("저장된 네이버 메일 계정으로 1건만 발송합니다.", html_source)
            self.assertIn("테스트 메일을 발송했습니다.", html_source)

    def test_mail_settings_redirects_cleanly_when_login_session_expires(self) -> None:
        for app_file in (
            ROOT / "scripts" / "workhub_delivery_app.py",
            ROOT / "_workhub_zip_inspect" / "scripts" / "workhub_delivery_app.py",
        ):
            html_source = app_file.read_text(encoding="utf-8")

            self.assertIn("function handleLoginRequiredResponse", html_source)
            self.assertIn("response.status === 401", html_source)
            self.assertIn("로그인이 만료되었습니다. 다시 로그인해주세요.", html_source)
            self.assertIn('window.location.href = "/login"', html_source)
            self.assertIn('credentials: "same-origin"', html_source)

    def test_distribution_mail_tree_only_opens_writable_mail_flows(self) -> None:
        for app_file in (
            ROOT / "scripts" / "workhub_delivery_app.py",
            ROOT / "_workhub_zip_inspect" / "scripts" / "workhub_delivery_app.py",
        ):
            html_source = app_file.read_text(encoding="utf-8")

            self.assertNotIn('data-mail-popup="supplier"', html_source)
            self.assertNotIn('data-mail-popup="seller"', html_source)
            self.assertIn('data-mail-popup="cs"', html_source)
            self.assertIn('data-mail-popup="stock"', html_source)
            self.assertIn("입고 및 품절 공지", html_source)
            self.assertIn('openModal(type === "stock" ? "mail-stock" : "cs")', html_source)
            self.assertIn('class="stock-notice-fields" id="stockNoticeFields"', html_source)
            self.assertIn('id="stockNoticeDateInput"', html_source)
            self.assertNotIn('id="stockManagerNameInput"', html_source)
            self.assertNotIn('id="stockManagerPhoneInput"', html_source)
            self.assertNotIn('id="stockSenderEmailInput"', html_source)
            self.assertIn("function defaultStockContactInfo()", html_source)
            self.assertIn("cachedMailSettings.naver_email ||", html_source)
            self.assertIn('id="stockVendorPickerButton"', html_source)
            self.assertIn('id="stockVendorTree"', html_source)
            self.assertIn('id="stockSelectedVendorLabel"', html_source)
            self.assertIn('id="stockRecipientEmailInput" type="hidden"', html_source)
            self.assertIn('id="stockVendorNameInput" type="hidden"', html_source)
            self.assertNotIn('for="stockRecipientEmailInput">받는 업체 메일', html_source)
            self.assertNotIn('for="stockVendorNameInput">업체명', html_source)
            self.assertIn("function renderStockVendorContacts()", html_source)
            self.assertIn("renderStockVendorContacts();", html_source)
            self.assertIn("async function loadVendorContacts()", html_source)
            self.assertIn("function applySelectedVendor()", html_source)
            self.assertIn('vendorContactSelect.addEventListener("change", applySelectedVendor)', html_source)
            self.assertIn("function applySelectedStockVendor()", html_source)
            self.assertIn("stockRecipientEmailInput.value = selected.email;", html_source)
            self.assertIn('stockVendorPickerButton?.addEventListener("click"', html_source)
            self.assertIn('stockVendorTree?.addEventListener("click"', html_source)
            self.assertIn("function defaultStockNoticeBody()", html_source)
            self.assertIn("function collectStockNoticePayload()", html_source)
            self.assertIn('stockNoticeFields.style.display = "block"', html_source)
            self.assertIn('"/api/mail-send"', html_source)
            self.assertIn('send_general_mail(payload)', html_source)
            self.assertEqual(1, html_source.count("function defaultStockNoticeBody()"))
            self.assertIn('} else if (mode === "cs") {', html_source)
            self.assertIn('} else if (mode === "mail-stock") {', html_source)
            self.assertNotIn('mode === "cs" || mode === "mail-stock"', html_source)
            self.assertNotIn('currentMode === "cs" || currentMode === "mail-stock"', html_source)
            self.assertNotIn("제품 입고 및 품절 안내드립니다", html_source)
            self.assertNotIn("안내 내용", html_source)
            self.assertNotIn("확인 후 관련 일정", html_source)
            self.assertNotIn("csBodyInput.value = defaultStockNoticeBody();", html_source)

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

    def test_tools_sidebar_is_shared_file_library_only(self) -> None:
        for app_file in (
            ROOT / "scripts" / "workhub_delivery_app.py",
            ROOT / "_workhub_zip_inspect" / "scripts" / "workhub_delivery_app.py",
        ):
            html_source = app_file.read_text(encoding="utf-8")

            self.assertIn('data-open="fileLibrary"', html_source)
            self.assertIn('id="fileLibraryWorkspace"', html_source)
            self.assertIn('id="sharedFileInput"', html_source)
            self.assertIn('id="sharedFileBody"', html_source)
            self.assertIn('"/api/shared-files"', html_source)
            self.assertIn('"/api/shared-file-upload"', html_source)
            self.assertIn('"/api/shared-file-delete"', html_source)
            self.assertIn('/api/shared-file-download?id=', html_source)
            self.assertIn("function loadSharedFiles()", html_source)
            self.assertIn("function uploadSharedFile()", html_source)
            self.assertIn("function downloadSharedFile", html_source)
            self.assertNotIn('data-open="invoice"><i data-lucide="file-spreadsheet"></i>', html_source)
            self.assertNotIn('data-open="vehicle"><i data-lucide="truck"></i>', html_source)

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
