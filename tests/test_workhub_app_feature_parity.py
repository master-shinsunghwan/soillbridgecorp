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
            self.assertIn("매출현황 및 관리 데이터 연결 대기 중", sales_panel_html)
            self.assertNotIn("선택한 날짜", sales_panel_html)
            self.assertNotIn("이번 달 요약", sales_panel_html)
            self.assertIn(".dashboard-calendar-panel", html_source)
            self.assertIn("grid-template-columns: minmax(0, 1fr) minmax(0, 1fr)", html_source)
            self.assertIn("function renderDashboardImportSchedule()", html_source)
            self.assertIn("renderDashboardImportSchedule();", html_source)
            self.assertIn('companyActiveTab === "notice" && panel.dataset.companyPanel === "calendar"', html_source)
            self.assertIn("loadDashboardEntryData().catch", html_source)

    def test_portal_notice_records_are_stored_in_server_db(self) -> None:
        app = self.load_app()
        app.init_db()

        saved = app.save_portal_notice(
            {
                "date": "2026-06-26",
                "title": "업무 포털 업데이트 안내",
                "owner": "관리자",
                "body": "공지사항 서버 저장 테스트",
            },
            {"username": "admin", "display_name": "Admin"},
        )
        rows = app.list_portal_notices()

        self.assertEqual(rows[0]["id"], saved["id"])
        self.assertEqual(rows[0]["date"], "2026-06-26")
        self.assertEqual(rows[0]["title"], "업무 포털 업데이트 안내")
        self.assertEqual(rows[0]["body"], "공지사항 서버 저장 테스트")
        self.assertEqual(app.delete_portal_notice(int(saved["id"])), 1)
        self.assertEqual(app.list_portal_notices(), [])

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
        self.assertIn('.nav-item[data-nav-tone="home"]', html_source)
        self.assertIn('.nav-item[data-nav-tone="management"]', html_source)
        self.assertIn('.nav-item[data-nav-tone="cs"]', html_source)
        self.assertIn('.nav-item[data-nav-tone="admin"]', html_source)
        self.assertIn('data-nav-tone="home"', html_source)
        self.assertIn('data-nav-tone="management"', html_source)
        self.assertIn('data-nav-tone="cs"', html_source)
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

    def test_ledger_sidebar_groups_open_directly_without_upload_subtrees(self) -> None:
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
        self.assertNotIn('id="managementImportOpen"', management_group)
        self.assertNotIn('data-management-import-mode="daily"', management_group)
        self.assertNotIn("통합관리대장 업로드", management_group)
        self.assertNotIn('class="nav-submenu"', management_group)

        self.assertIn('data-open="ledger"', ledger_group)
        self.assertNotIn('id="ledgerImportOpen"', ledger_group)
        self.assertNotIn('data-ledger-import-mode="daily"', ledger_group)
        self.assertNotIn("CS처리대장 업로드", ledger_group)
        self.assertNotIn('data-mail-popup="cs"', ledger_group)
        self.assertNotIn("CS처리 요청", ledger_group)
        self.assertNotIn('class="nav-submenu"', ledger_group)

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

        self.assertNotIn('data-mail-popup="cs"', ledger_group)
        self.assertNotIn("CS처리 요청", ledger_group)
        self.assertIn('id="csAttachmentInput"', html_source)
        self.assertIn('id="csAttachmentDropzone"', html_source)
        self.assertIn('id="csAttachmentChoose"', html_source)
        self.assertIn('id="csAttachmentDropMain"', html_source)
        self.assertIn('id="csAttachmentSummary"', html_source)
        self.assertIn('accept="image/*,video/*,.pdf,.xlsx,.xls,.doc,.docx,.zip"', html_source)
        self.assertIn("setupDropzone(csAttachmentDropzone", html_source)
        self.assertIn('csAttachmentChoose?.addEventListener("click"', html_source)
        self.assertIn("파일 선택", html_source)
        self.assertIn("드래그", html_source)
        self.assertIn('id="sendCsMailButton"', html_source)
        self.assertIn("async function sendCurrentCsMail()", html_source)
        self.assertIn('sendCsMailButton?.addEventListener("click"', html_source)
        self.assertIn("appendCsMailPayload", html_source)
        self.assertIn("collect_mail_attachments", html_source)
        self.assertIn("attachments=attachments", html_source)

    def test_ledger_cs_popup_and_cell_edit_activation_are_workspace_safe(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn("function mountCsFieldsInLedgerWorkspace()", html_source)
        self.assertIn("ledgerWorkspaceMount.appendChild(csFields)", html_source)
        self.assertIn("#ledgerWorkspace .cs-fields.ledger-cs-popup", html_source)
        self.assertIn("position: fixed;", html_source)
        self.assertIn("transform: translate(-50%, -50%);", html_source)
        self.assertIn("grid-template-columns: repeat(2, minmax(0, 1fr));", html_source)
        self.assertIn("text-field cs-wide", html_source)
        self.assertIn("restoreCsFieldsToModal();", html_source)

        self.assertIn('tabindex="0"', html_source)
        self.assertIn("function selectEditableCell(scope, cell)", html_source)
        self.assertIn('managementBody.addEventListener("dblclick"', html_source)
        self.assertIn('ledgerBody.addEventListener("dblclick"', html_source)
        self.assertIn("function handleEditableCellNavigation(scope, event)", html_source)
        self.assertIn('data-ledger-edit-row', html_source)
        management_click_start = html_source.index('managementBody.addEventListener("click"')
        management_click_end = html_source.index('managementBody.addEventListener("dblclick"', management_click_start)
        ledger_click_start = html_source.index('ledgerBody.addEventListener("click"')
        ledger_click_end = html_source.index('ledgerBody.addEventListener("dblclick"', ledger_click_start)
        self.assertNotIn('openCellEditor("management"', html_source[management_click_start:management_click_end])
        self.assertIn('const editButton = event.target.closest("[data-ledger-edit-row]");', html_source[ledger_click_start:ledger_click_end])
        self.assertIn('if (editableCell) openCellEditor("ledger", editableCell);', html_source[ledger_click_start:ledger_click_end])
        ledger_cell_click = html_source[
            html_source.index('const editableCell = event.target.closest(".editable-cell[data-field]");', ledger_click_start):ledger_click_end
        ]
        self.assertNotIn('openCellEditor("ledger"', ledger_cell_click)

    def test_sheet_like_tables_support_drag_selection_and_numeric_summary(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn("sheetRangeSelection", html_source)
        self.assertIn("sheet-selection-summary", html_source)
        self.assertIn("function applySheetRangeSelection(scope, anchor, current)", html_source)
        self.assertIn("function updateSheetSelectionSummary(cells)", html_source)
        self.assertIn("function numericValueFromSheetCell(cell)", html_source)
        self.assertIn("합계 ${formatSheetNumber(sum)}", html_source)
        self.assertIn("평균 ${formatSheetNumber(average)}", html_source)
        self.assertIn('managementBody.addEventListener("mousedown"', html_source)
        self.assertIn('ledgerBody.addEventListener("mousedown"', html_source)
        self.assertIn('managementBody.addEventListener("mousemove"', html_source)
        self.assertIn('ledgerBody.addEventListener("mousemove"', html_source)
        self.assertIn('document.addEventListener("mouseup", finishSheetRangeSelection)', html_source)

    def test_crm_daily_work_logs_are_available_from_work_management(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")
        crm_source = (ROOT / "scripts" / "workhub_crm.py").read_text(encoding="utf-8")

        self.assertIn('data-crm-nav-tab="daily"', html_source)
        self.assertIn('id="crmTabDaily"', html_source)
        self.assertIn('id="crmPanelDaily"', html_source)
        self.assertIn('id="crmDailyLogForm"', html_source)
        self.assertIn('id="crmDailyLogBody"', html_source)
        self.assertIn("crm-daily-widget", html_source)
        self.assertIn("crm-daily-log-list", html_source)
        self.assertIn("crm-daily-log-card", html_source)
        self.assertIn("grid-template-columns: repeat(auto-fit, minmax(min(100%, 300px), 1fr));", html_source)
        self.assertIn("overflow-wrap: anywhere;", html_source)
        self.assertIn("word-break: keep-all;", html_source)
        self.assertIn("오늘 한 일", html_source)
        self.assertIn("내일 할 일", html_source)
        self.assertNotIn('id="crmDailyFormSummary"', html_source)
        self.assertNotIn('id="crmDailyDraftButton"', html_source)
        self.assertIn("/api/crm-daily-logs", html_source)
        self.assertIn("/api/crm-daily-log-save", html_source)
        self.assertIn("function renderCrmDailyLogs", html_source)
        self.assertIn("function saveCrmDailyLogForm", html_source)
        self.assertIn("crm_daily_logs", crm_source)
        self.assertIn("completed_today_tasks", crm_source)
        self.assertIn("open_task_samples", crm_source)
        self.assertIn("def list_crm_daily_logs", crm_source)
        self.assertIn("def save_crm_daily_log", crm_source)

    def test_ledger_completed_rows_ignore_whitespace_variants_for_yellow_highlight(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn("function isOverallCompletedStatus(statusValue)", html_source)
        self.assertIn('status.includes("전체") && status.includes("처리") && status.includes("완료")', html_source)
        self.assertIn("if (isOverallCompletedStatus(statusValue)) return true;", html_source)
        self.assertIn("const type = normalizedLedgerText(typeValue);", html_source)
        self.assertIn("const status = normalizedLedgerText(statusValue);", html_source)
        self.assertIn("function isLedgerCompletedCase(csCase)", html_source)
        self.assertIn('if (status.includes("완료")) return true;', html_source)
        self.assertIn("if (String(csCase.completed_at || '').trim()) return true;", html_source)
        self.assertIn(".ledger-table tbody tr.completed-cs td", html_source)
        self.assertIn("background: #fff8d8 !important;", html_source)
        self.assertIn('field: "completed_at", label: "완료일", value: csCase.completed_at', html_source)
        self.assertIn("const completedAt = fieldValue(row.querySelector('[data-field=\"completed_at\"]'));", html_source)
        self.assertIn('row.classList.add("completed-cs")', html_source)
        self.assertIn('row.classList.toggle("completed-cs", isCompletedByValues(csType, status) || Boolean(completedAt));', html_source)

    def test_management_duplicate_rows_override_even_row_striping(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn(".ledger-table tbody tr:nth-child(even) td", html_source)
        self.assertIn(".ledger-table tbody tr.management-duplicate td", html_source)
        self.assertIn("background-color: var(--duplicate-row-color, #eef6ff);", html_source)
        self.assertIn('row.classList.add("management-duplicate")', html_source)

    def test_hermes_workspace_menu_and_configurable_agent_bridge_exist(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('id="hermesNavGroup"', html_source)
        self.assertIn('data-open="hermes"', html_source)
        self.assertIn("AI 업무채팅", html_source)
        self.assertIn("자동화 요청", html_source)
        self.assertIn("작업내역", html_source)
        self.assertIn("관리자 설정", html_source)
        self.assertIn('id="hermesWorkspace"', html_source)
        self.assertIn('id="hermesBaseUrl"', html_source)
        self.assertIn('id="hermesConnectionTest"', html_source)
        self.assertIn("/api/hermes-status", html_source)
        self.assertIn("/api/hermes-chat", html_source)
        self.assertIn("/api/hermes-automation", html_source)
        self.assertIn('data-hermes-chat-mode="auto"', html_source)
        self.assertIn('data-hermes-chat-mode="automation"', html_source)
        self.assertIn('data-hermes-chat-mode="general"', html_source)
        self.assertIn('data-hermes-chat-mode="search"', html_source)
        self.assertIn('data-hermes-chat-mode="image"', html_source)
        self.assertIn("function setHermesChatMode(mode)", html_source)
        self.assertIn("body: JSON.stringify({ message, mode: requestedMode })", html_source)
        self.assertIn("save_hermes_text_result", html_source)
        self.assertIn("generated_text_file", html_source)
        self.assertIn("hermes-result-link", html_source)
        self.assertIn("HERMES_SETTINGS_PATH", html_source)
        self.assertIn("hermes_request", html_source)
        self.assertIn("hermes_use", html_source)
        self.assertIn("hermes_automation", html_source)
        self.assertIn("hermes_admin", html_source)
        self.assertIn("hermes-mark", html_source)
        self.assertIn("/static/hermes-icon.png", html_source)
        self.assertIn('data-hermes-preset="vps"', html_source)
        self.assertIn('data-hermes-preset="local"', html_source)
        self.assertIn("HERMES_PROFILE", html_source)
        self.assertIn("describe_hermes_connection_error", html_source)
        self.assertIn('api_key.lower().startswith(("basic ", "bearer "))', html_source)

    def test_hermes_chat_history_can_be_summarized_and_filtered(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('id="hermesSummaryCreate"', html_source)
        self.assertIn('id="hermesHistoryFilter"', html_source)
        self.assertIn('id="hermesSummaryList"', html_source)
        self.assertIn("/api/hermes-summary", html_source)
        self.assertIn("summarize_hermes_chat_history", html_source)
        self.assertIn("renderHermesSummaries", html_source)
        self.assertIn("createHermesSummary", html_source)
        self.assertIn('kind: "summary"', html_source)
        self.assertEqual(html_source.count("function renderHermesHistory(items = [])"), 1)

    def test_hermes_chat_uses_half_width_answer_and_quick_actions(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn("hermes-chat-layout", html_source)
        self.assertIn("hermes-side-grid", html_source)
        self.assertIn("hermes-quick-actions", html_source)
        self.assertIn('data-hermes-quick="summary"', html_source)
        self.assertIn('data-hermes-quick="history"', html_source)
        self.assertIn('data-hermes-quick="automation"', html_source)
        self.assertIn('data-hermes-quick="settings"', html_source)
        self.assertIn('setHermesTab("history")', html_source)
        self.assertIn('setHermesTab("automation")', html_source)
        self.assertIn('setHermesTab("settings")', html_source)

    def test_dashboard_sales_amounts_show_full_numbers_not_million_suffix(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        start = html_source.index("function formatSalesMillion")
        end = html_source.index("function formatSalesPercent", start)
        formatter_slice = html_source[start:end]
        self.assertIn('Math.abs(number).toLocaleString("ko-KR")', formatter_slice)
        self.assertIn('}원`', formatter_slice)
        self.assertNotIn("백만", formatter_slice)

    def test_admin_navigation_and_admin_pages_can_scroll(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn("height: 100vh;", html_source)
        self.assertIn("max-height: 100vh;", html_source)
        self.assertIn("overscroll-behavior: contain;", html_source)
        self.assertIn("#userAdminWorkspace.active", html_source)
        self.assertIn("#backupWorkspace.active", html_source)
        self.assertIn("#systemUpdateWorkspace.active", html_source)
        self.assertIn("main.workspace-scroll-mode", html_source)
        self.assertIn("main:has(#userAdminWorkspace.active)", html_source)
        self.assertIn("main:has(#backupWorkspace.active)", html_source)
        self.assertIn("main:has(#systemUpdateWorkspace.active)", html_source)
        self.assertIn('"workspace-scroll-mode"', html_source)
        self.assertIn("showUserAdmin || showBackup || showSystemUpdate", html_source)
        self.assertIn("#userAdminPermissions.permission-grid", html_source)
        self.assertIn("max-height: min(40vh, 380px);", html_source)
        self.assertIn("#userAdminWorkspace .admin-table-wrap", html_source)
        self.assertIn("max-height: min(44vh, 460px);", html_source)
        self.assertIn("position: sticky;", html_source)

    def test_user_admin_can_delete_edit_and_review_deleted_history(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn("/api/users-delete", html_source)
        self.assertIn("delete_user_account", html_source)
        self.assertIn("deleted_user_accounts", html_source)
        self.assertIn("list_deleted_user_accounts", html_source)
        self.assertIn('id="userAdminDeletedBody"', html_source)
        self.assertIn('data-user-edit="${user.id}"', html_source)
        self.assertIn('data-user-delete="${user.id}"', html_source)
        self.assertIn("비밀번호는 보안상 확인할 수 없으며", html_source)
        self.assertNotIn("password_hash TEXT", html_source[html_source.index("CREATE TABLE IF NOT EXISTS deleted_user_accounts"):html_source.index("CREATE TABLE IF NOT EXISTS shared_files")])

    def test_ledger_arrow_keys_move_between_cells_before_editing(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn("function handleEditableCellNavigation(scope, event)", html_source)
        self.assertIn("function moveEditableCell(scope, cell, rowDelta, colDelta)", html_source)
        self.assertIn("ArrowUp", html_source)
        self.assertIn("ArrowDown", html_source)
        self.assertIn("ArrowLeft", html_source)
        self.assertIn("ArrowRight", html_source)
        self.assertIn("event.preventDefault();", html_source)
        self.assertIn('handleEditableCellNavigation("management", event);', html_source)
        self.assertIn('handleEditableCellNavigation("ledger", event);', html_source)
        self.assertIn('event.key === "F2"', html_source)

    def test_cell_editor_apply_saves_and_refreshes_visible_ledgers(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn("async function saveEditedCellRow(scope, row)", html_source)
        self.assertIn("function scheduleCellEditorAutoApply(scope)", html_source)
        self.assertIn("function commitCellEditorOnChange(scope)", html_source)
        self.assertIn("control.addEventListener(\"input\", () => scheduleCellEditorAutoApply(scope));", html_source)
        self.assertIn("control.addEventListener(\"change\", () => commitCellEditorOnChange(scope));", html_source)
        self.assertIn("control.addEventListener(\"blur\", () => commitCellEditorOnChange(scope));", html_source)
        apply_start = html_source.index("async function applyCellEditor(scope")
        apply_end = html_source.index("function dirtyRows", apply_start)
        apply_block = html_source[apply_start:apply_end]
        self.assertIn("await saveEditedCellRow(scope, row);", apply_block)
        self.assertIn('notice.textContent = scope === "management"', apply_block)

        save_start = html_source.index("async function saveEditedCellRow(scope, row)")
        save_end = html_source.index("function dirtyRows", save_start)
        save_block = html_source[save_start:save_end]
        self.assertIn("await saveManagementPayload(payload);", save_block)
        self.assertIn("updateManagementRecordCache(payload);", save_block)
        self.assertIn("applyManagementFilters();", save_block)
        self.assertIn("await saveLedgerPayload(payload);", save_block)
        self.assertIn("updateLedgerCaseCache(payload);", save_block)
        self.assertIn("applyLedgerFilters();", save_block)

    def test_ledger_rows_expose_visible_edit_button_and_order_field_updates(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('data-ledger-edit-row', html_source)
        self.assertIn('openCellEditor("ledger", editableCell);', html_source)
        self.assertIn('sales_vendor: fieldValue(row.querySelector(\'[data-field="sales_vendor"]\'))', html_source)
        self.assertIn('purchase_vendor: fieldValue(row.querySelector(\'[data-field="purchase_vendor"]\'))', html_source)
        self.assertIn('order_date: fieldValue(row.querySelector(\'[data-field="order_date"]\'))', html_source)
        self.assertIn('receiver_address: fieldValue(row.querySelector(\'[data-field="receiver_address"]\'))', html_source)
        self.assertIn('original_invoice: fieldValue(row.querySelector(\'[data-field="original_invoice"]\'))', html_source)
        self.assertIn("CS_CASE_UPDATE_FIELDS", html_source)
        self.assertIn('"receiver_address"', html_source)
        self.assertIn('"original_invoice"', html_source)

    def test_checked_rows_can_receive_selected_cell_value_in_bulk(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('id="managementBulkApply"', html_source)
        self.assertIn('id="ledgerBulkApply"', html_source)
        self.assertIn("async function applySelectedCellToCheckedRows(scope)", html_source)
        self.assertIn("selectedRows(body, rowSelector)", html_source)
        self.assertIn("const fieldSelector = scope === \"management\"", html_source)
        self.assertIn("setEditableCellValue(targetCell, value);", html_source)
        self.assertIn("await saveCurrentWorkspaceRows({ mode: scope, selectedOnly: true });", html_source)
        self.assertIn('managementBulkApply.addEventListener("click"', html_source)
        self.assertIn('ledgerBulkApply.addEventListener("click"', html_source)

    def test_management_manual_add_uses_popup_workflow(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('id="managementAddManual"', html_source)
        self.assertIn('id="managementManualFields"', html_source)
        self.assertIn('id="managementManualClose"', html_source)
        self.assertIn('id="saveManagementManual"', html_source)
        self.assertIn("function openManagementManualPopup()", html_source)
        self.assertIn("function closeManagementManualPopup()", html_source)
        self.assertIn("async function saveManagementManualRecord()", html_source)
        self.assertIn('id="manualManagementReceiverAddress" data-management-manual-field="receiver_address" rows="2"', html_source)
        self.assertIn('id="manualManagementMemo" data-management-manual-field="memo" rows="2"', html_source)
        self.assertIn(".management-manual-fields #manualManagementReceiverAddress", html_source)
        self.assertIn(".management-manual-fields #manualManagementMemo", html_source)
        self.assertIn("height: 52px;", html_source)
        self.assertIn('fetch("/api/management-record-create"', html_source)
        self.assertIn('managementAddManual.addEventListener("click"', html_source)

    def test_workhub_uses_in_app_dialogs_instead_of_native_browser_dialogs(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn("function requestAppConfirm", html_source)
        self.assertNotIn("window.alert(", html_source)
        self.assertNotIn("window.confirm(", html_source)
        self.assertNotIn("if (!confirm(", html_source)
        self.assertNotIn("alert(error.message)", html_source)

    def test_excel_style_filter_reset_controls_exist_for_ledgers(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('id="ledgerFilterResetAll"', html_source)
        self.assertIn("function clearActiveLedgerFilter()", html_source)
        self.assertIn("function clearAllLedgerFilters()", html_source)
        self.assertIn("function clearAllManagementFilters()", html_source)
        self.assertIn("delete ledgerFilters[activeLedgerFilterField];", html_source)
        self.assertIn("delete managementFilters[activeManagementFilterField];", html_source)
        self.assertIn("Object.keys(ledgerFilters).forEach", html_source)
        self.assertIn("Object.keys(managementFilters).forEach", html_source)
        self.assertIn('ledgerFilterResetAll.addEventListener("click"', html_source)

    def test_ledger_filter_options_follow_other_active_filters(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn("function matchesLedgerFiltersExcept(csCase, excludedField = \"\")", html_source)
        self.assertIn("function matchesManagementFiltersExcept(record, excludedField = \"\")", html_source)
        self.assertIn(".filter((csCase) => matchesLedgerFiltersExcept(csCase, field))", html_source)
        self.assertIn(".filter((record) => matchesManagementFiltersExcept(record, field))", html_source)
        self.assertIn("if (field === excludedField) return true;", html_source)

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
            self.assertIn('id="leaveAdminUserSearch"', html_source)
            self.assertIn('id="leaveAdminUserList"', html_source)
            self.assertIn('data-leave-tab="mine">내 연차</button>', html_source)
            self.assertIn('data-leave-tab="request">신청하기</button>', html_source)
            self.assertIn('data-leave-tab="approvals">승인 관리</button>', html_source)
            self.assertIn('data-leave-tab="admin">직원 기준</button>', html_source)
            self.assertIn('data-leave-tab="holidays">공휴일 설정</button>', html_source)
            self.assertIn('data-leave-tab="history">이력</button>', html_source)
            self.assertIn('id="leaveTabHolidays"', html_source)
            self.assertIn('id="leaveTabHistory"', html_source)
            self.assertIn("main:has(#leaveWorkspace.active)", html_source)
            self.assertIn("#leaveWorkspace.active", html_source)
            self.assertIn("can_view_staff", html_source)
            self.assertIn("function renderLeaveAdminUserList", html_source)
            self.assertIn("function selectLeaveAdminUser", html_source)
            self.assertIn("data-leave-admin-user", html_source)
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

        self.assertIn("통합관리대장 업로드", html_source)
        self.assertIn("CS처리대장 업로드", html_source)
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

    def test_ledger_imports_show_progress_dialog(self) -> None:
        html_source = (SCRIPTS / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('id="importProgressDialog"', html_source)
        self.assertIn('id="importProgressStage"', html_source)
        self.assertIn('id="importProgressSteps"', html_source)
        self.assertIn("function showImportProgress", html_source)
        self.assertIn("function updateImportProgress", html_source)
        self.assertIn("function finishImportProgress", html_source)
        self.assertIn('updateImportProgress("preview"', html_source)
        self.assertIn('updateImportProgress("saving"', html_source)
        self.assertIn('updateImportProgress("done"', html_source)

    def test_backup_workspace_supports_auto_and_selected_backup_settings(self) -> None:
        html_source = (SCRIPTS / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('id="backupAutoEnabled"', html_source)
        self.assertIn('id="backupAutoHour"', html_source)
        self.assertIn('id="backupRetentionDays"', html_source)
        self.assertIn('id="backupDirInput"', html_source)
        self.assertIn('id="backupSettingsSave"', html_source)
        self.assertIn('id="backupCreateSelected"', html_source)
        self.assertIn('id="backupOfflineDownload"', html_source)
        self.assertIn('id="backupExternalEnabled"', html_source)
        self.assertIn('id="backupRcloneRemote"', html_source)
        self.assertIn('id="backupRclonePath"', html_source)
        self.assertIn('id="backupRcloneExecutable"', html_source)
        self.assertIn('id="backupExternalStatus"', html_source)
        self.assertIn("/api/backup-settings", html_source)
        self.assertIn("/api/backup-offline-download", html_source)
        self.assertIn("function saveBackupSettings", html_source)
        self.assertIn("function createBackupAtSelectedPath", html_source)
        self.assertIn("function downloadOfflineBackup", html_source)
        self.assertIn("Google Drive 업로드 상태", html_source)
        self.assertIn("rclone_remote", html_source)
        self.assertIn("load_backup_settings", html_source)
        self.assertIn("save_backup_settings", html_source)
        self.assertIn("BACKUP_DATA_DIRECTORIES", html_source)
        self.assertIn("write_backup_directory", html_source)
        self.assertIn("restore_backup_directory", html_source)
        self.assertIn('"backup_scope"', html_source)

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
            self.assertIn('id="stockVendorTypeFilter"', html_source)
            self.assertIn('id="stockVendorSearchInput"', html_source)
            self.assertIn('id="stockVendorPickerButton"', html_source)
            self.assertIn('id="stockVendorTree"', html_source)
            self.assertIn('id="stockSelectedVendorLabel"', html_source)
            self.assertIn('id="stockRecipientEmailInput" type="hidden"', html_source)
            self.assertIn('id="stockVendorNameInput" type="hidden"', html_source)
            self.assertNotIn('for="stockRecipientEmailInput">받는 업체 메일', html_source)
            self.assertNotIn('for="stockVendorNameInput">업체명', html_source)
            self.assertIn("function renderStockVendorContacts()", html_source)
            self.assertIn("let selectedStockVendors = [];", html_source)
            self.assertIn("data-stock-vendor-select-all", html_source)
            self.assertIn("data-stock-vendor-checkbox", html_source)
            self.assertIn("function stockVendorMatchesSearch(contact", html_source)
            self.assertIn("function visibleStockVendorContactsForPicker()", html_source)
            self.assertIn("function toggleAllStockVendorsForType(checked)", html_source)
            self.assertIn("async function sendCurrentStockNoticeMail()", html_source)
            self.assertIn("renderStockVendorContacts();", html_source)
            self.assertIn("async function loadVendorContacts()", html_source)
            self.assertIn("function applySelectedVendor()", html_source)
            self.assertIn('vendorContactSelect.addEventListener("change", applySelectedVendor)', html_source)
            self.assertIn("function applySelectedStockVendor()", html_source)
            self.assertIn('stockRecipientEmailInput.value = first?.email || "";', html_source)
            self.assertIn('stockVendorPickerButton?.addEventListener("click"', html_source)
            self.assertIn('stockVendorSearchInput?.addEventListener("input"', html_source)
            self.assertIn('stockVendorTree?.addEventListener("change"', html_source)
            self.assertIn("function defaultStockNoticeBody()", html_source)
            self.assertIn("function collectStockNoticePayload(vendor = null)", html_source)
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

    def test_sales_report_upload_has_dedicated_sales_workspace_mode(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('id="salesReportFileInput"', html_source)
        self.assertIn('id="salesReportUploadCard"', html_source)
        self.assertIn('id="salesReportNavGroup"', html_source)
        self.assertIn('id="salesReportNavToggle"', html_source)
        self.assertIn('"sales_report_manage"', html_source)
        self.assertIn("__SALES_REPORT_NAV__", html_source)
        self.assertIn('accept=".xlsx,.xlsm,.xls,.csv,.zip"', html_source)
        self.assertIn('data-open="salesReport">매출표 업로드</button>', html_source)
        self.assertIn('mode === "salesReport"', html_source)
        self.assertIn("sales-report-only", html_source)
        self.assertIn("#userAdminWorkspace.sales-report-only > .workspace-head", html_source)
        self.assertIn('!can("sales_report_manage")', html_source)
        self.assertIn('"/api/sales-report-upload"', html_source)
        self.assertIn('"/api/sales-report-uploads"', html_source)
        self.assertIn("function uploadSalesReportWorkbook()", html_source)
        self.assertIn("function openSalesReportUploadPicker()", html_source)
        self.assertIn("function loadSalesReportUploads()", html_source)
        self.assertIn('document.querySelector("#salesReportNavToggle")?.addEventListener("click"', html_source)
        self.assertIn('showWorkspace("salesReport")', html_source)
        self.assertIn('button.closest("#salesReportNavGroup")', html_source)
        self.assertNotIn('id="salesReportDropMain"', html_source)
        self.assertNotIn('id="salesReportChooseFile"', html_source)
        self.assertIn('id="salesReportManualUpload"', html_source)
        self.assertIn('salesReportManualUpload?.addEventListener("click", openSalesReportUploadPicker);', html_source)
        self.assertIn('id="salesReportUploadMessage"', html_source)
        self.assertIn('id="salesReportRecentList"', html_source)
        self.assertNotIn('label[for=\'salesReportFileInput\']', html_source)

        admin_nav_start = html_source.index("ADMIN_TOOLS_NAV_HTML")
        admin_nav_end = html_source.index("SALES_REPORT_NAV_HTML", admin_nav_start)
        admin_nav_slice = html_source[admin_nav_start:admin_nav_end]
        self.assertNotIn('data-admin-focus="salesReport"', admin_nav_slice)
        self.assertNotIn('data-open="salesReport"', admin_nav_slice)

        admin_start = html_source.index('id="userAdminWorkspace"')
        admin_end = html_source.index('id="userAdminBody"', admin_start)
        admin_slice = html_source[admin_start:admin_end]
        self.assertIn('id="salesReportFileInput"', admin_slice)
        self.assertIn('>매출현황</div>', admin_slice)
        self.assertIn('id="salesReportManualUpload"', admin_slice)
        self.assertNotIn('id="salesNasImportDir"', admin_slice)
        self.assertNotIn('id="salesNasScanNow"', admin_slice)
        self.assertNotIn("NAS 자동 업로드", admin_slice)

    def test_sales_report_permission_shows_menu_for_non_admin_staff(self) -> None:
        app = self.load_app()

        html_source = app.render_app_html({
            "username": "sales-staff",
            "display_name": "매출담당",
            "role": "user",
            "permissions": ["sales_report_manage"],
        })

        self.assertIn('id="salesReportNavGroup"', html_source)
        self.assertIn('data-open="salesReport">매출표 업로드</button>', html_source)
        self.assertIn('id="userAdminWorkspace"', html_source)
        self.assertIn('id="salesReportUploadCard"', html_source)
        self.assertNotIn('id="adminToolsNavGroup"', html_source)
        self.assertIn('mode === "userAdmin" && (!userAdminWorkspace || currentUser.role !== "admin")', html_source)

    def test_sales_report_dashboard_layout_uses_three_report_types(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn('"/api/sales-report-dashboard"', html_source)
        self.assertIn("sales_report_dashboard_payload", html_source)
        self.assertIn("function loadSalesReportDashboard()", html_source)
        self.assertIn("function renderSalesReportDashboard(data)", html_source)
        self.assertIn("function formatSalesNumber(value)", html_source)
        self.assertIn("function formatSalesPercent(value)", html_source)
        self.assertIn('id="salesReportKpiGrid"', html_source)
        self.assertIn('id="salesReportDailyBody"', html_source)
        self.assertIn('id="salesReportSellerBody"', html_source)
        self.assertIn('id="salesReportProductBody"', html_source)
        self.assertIn('id="salesReportReviewBody"', html_source)
        self.assertIn("날짜별", html_source)
        self.assertIn("상품별", html_source)
        self.assertIn("매출처별", html_source)
        self.assertIn("매입처별 총합계 금액", html_source)
        self.assertNotIn('["파일 검증"', html_source)

    def test_daily_sales_detail_popup_uses_full_height_with_compact_density(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertIn(".sales-detail-popup.compact", html_source)
        self.assertIn(".sales-detail-popup.expanded", html_source)
        self.assertIn(".sales-detail-popup.compact .sales-detail-table-wrap", html_source)
        self.assertIn(".sales-detail-popup.expanded .sales-detail-table-wrap", html_source)
        self.assertIn(".sales-detail-sections", html_source)
        self.assertIn(".sales-detail-popup.compact .sales-detail-sections", html_source)
        self.assertIn(".sales-detail-popup.expanded .sales-detail-sections", html_source)
        self.assertIn('salesDetailBody.classList.toggle("has-note", Boolean(data.note));', html_source)
        self.assertIn('salesDetailBody.innerHTML = `${metricHtml}${noteHtml}<div class="sales-detail-sections">${sectionsHtml}</div>`;', html_source)
        self.assertIn("height: calc(100vh - 40px);", html_source)
        self.assertIn("grid-template-rows: auto auto minmax(0, 1fr);", html_source)
        self.assertIn("min-height: 0;", html_source)
        self.assertIn("height: 100% !important;", html_source)
        self.assertIn('const salesDetailPanel = document.querySelector("#salesDetailPopup .sales-detail-popup");', html_source)
        self.assertIn("function setSalesDetailPopupMode(kind)", html_source)
        self.assertIn('salesDetailPanel.classList.toggle("compact", compact);', html_source)
        self.assertIn('salesDetailPanel.classList.toggle("expanded", !compact);', html_source)
        self.assertIn("setSalesDetailPopupMode(data.kind);", html_source)
        self.assertIn("setSalesDetailPopupMode(kind);", html_source)

    def test_lucide_imports_do_not_request_missing_arrow_right_export(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")

        self.assertNotIn("Search, ArrowRight, LogOut", html_source)
        self.assertNotIn('"arrow-right": ArrowRight', html_source)
        self.assertIn('"arrow-right": ChevronRight', html_source)

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

    def test_work_management_tree_links_to_shared_file_library(self) -> None:
        html_source = (ROOT / "scripts" / "workhub_delivery_app.py").read_text(encoding="utf-8")
        crm_group_start = html_source.index('id="crmNavGroup"')
        crm_group_end = html_source.index("__HERMES_NAV__", crm_group_start)
        crm_group = html_source[crm_group_start:crm_group_end]

        self.assertIn('data-open="fileLibrary"', crm_group)
        self.assertIn("업무 파일", crm_group)
        self.assertIn('id="fileLibraryWorkspace"', html_source)
        self.assertIn('id="sharedFileInput"', html_source)
        self.assertIn('id="sharedFileBody"', html_source)

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
