# Sales Report Dashboard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the first usable Workhub sales dashboard from the three approved 얼마에요 export file types.

**Architecture:** Extend the existing `scripts/workhub_delivery_app.py` sales upload flow. Uploaded files are detected as daily, product, or seller reports, parsed into SQLite tables, and served through a new dashboard API consumed by the existing `salesReport` workspace.

**Tech Stack:** Python stdlib, SQLite, `openpyxl`, `pandas.read_html` for HTML-table `.xls`, existing vanilla HTML/CSS/JS in `scripts/workhub_delivery_app.py`, `unittest`.

---

### Task 1: Parser And Storage

**Files:**
- Modify: `scripts/workhub_delivery_app.py`
- Test: `tests/test_sales_report_uploads.py`

- [ ] **Step 1: Write failing parser tests**

Add tests that create three small sample files: a daily `.xlsx`, a product HTML-table `.xls`, and a seller `.xlsx`. Assert that `detect_sales_report_type()`, `parse_sales_report_file()`, `save_sales_report_file()`, and `sales_report_dashboard_payload()` return expected daily totals, seller totals, product TOP rows, and consistency difference.

- [ ] **Step 2: Run parser tests to verify RED**

Run:

```powershell
C:\Users\ssh19\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m unittest tests.test_sales_report_uploads
```

Expected: fail because parser/dashboard functions do not exist.

- [ ] **Step 3: Implement parser and DB tables**

Add these functions to `scripts/workhub_delivery_app.py`:

- `detect_sales_report_type(path, original_name="")`
- `parse_sales_report_file(path, original_name="")`
- `save_sales_report_snapshot(file_id, parsed)`
- `sales_report_dashboard_payload(period="", report_date="")`

Add SQLite tables:

- `sales_report_daily_rows`
- `sales_report_product_rows`
- `sales_report_seller_rows`

Update `save_sales_report_file()` to detect, parse, and store recognized sales files while keeping upload history.

- [ ] **Step 4: Run parser tests to verify GREEN**

Run:

```powershell
C:\Users\ssh19\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m unittest tests.test_sales_report_uploads
```

Expected: pass.

### Task 2: API And UI

**Files:**
- Modify: `scripts/workhub_delivery_app.py`
- Test: `tests/test_workhub_app_feature_parity.py`

- [ ] **Step 1: Write failing UI/API parity test**

Assert that the app contains:

- `/api/sales-report-dashboard`
- `loadSalesReportDashboard`
- `renderSalesReportDashboard`
- `salesReportKpiGrid`
- `salesReportDailyBody`
- `salesReportSellerBody`
- `salesReportProductBody`
- `salesReportReviewBody`

- [ ] **Step 2: Run UI/API test to verify RED**

Run:

```powershell
C:\Users\ssh19\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m unittest tests.test_workhub_app_feature_parity.WorkhubAppFeatureParityTests.test_sales_report_dashboard_layout_uses_three_report_types
```

Expected: fail because the dashboard API and UI elements do not exist.

- [ ] **Step 3: Implement dashboard API and UI**

Add `GET /api/sales-report-dashboard`, protected by `sales_report_manage`.

Replace the sales report-only workspace body with:

- Upload card with three recognized file types.
- KPI grid.
- Daily trend table.
- Seller TOP table.
- Product TOP table.
- Review-needed table.

Add frontend functions:

- `loadSalesReportDashboard()`
- `renderSalesReportDashboard(data)`
- `formatSalesNumber(value)`
- `formatSalesPercent(value)`

Update upload success to reload both recent files and dashboard data.

- [ ] **Step 4: Run UI/API test to verify GREEN**

Run the same parity test and ensure it passes.

### Task 3: Verification And Persistence

**Files:**
- Modify: code and tests from Tasks 1-2

- [ ] **Step 1: Run full automated verification**

Run:

```powershell
C:\Users\ssh19\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m py_compile scripts\workhub_delivery_app.py
C:\Users\ssh19\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m unittest discover -s tests
git diff --check
```

- [ ] **Step 2: Browser verification**

Restart Workhub on `http://127.0.0.1:8770/`, open the sales report workspace, and verify the KPI/table layout appears without console errors.

- [ ] **Step 3: Commit and push**

Run:

```powershell
git add scripts\workhub_delivery_app.py tests\test_sales_report_uploads.py tests\test_workhub_app_feature_parity.py docs\superpowers\plans\2026-06-19-sales-report-dashboard.md
git commit -m "Build sales report dashboard"
git push origin main
```
