# Sales Report Dashboard Design

## Product Context

Workhub needs a sales status screen under `매출현황 및 관리`. The screen should turn uploaded 얼마에요 sales exports into a daily operating view for Soillbridge: today versus yesterday, month-to-date totals, seller rankings, product rankings, margin warnings, CS impact, and file consistency checks.

The current app is a single Python HTTP application centered on `scripts/workhub_delivery_app.py`. A sales report upload screen already exists and is protected by the `sales_report_manage` permission.

## Source Files

The dashboard will support three upload file types.

1. `매출 통계.xlsx`
   - Purpose: date-by-date sales trend.
   - Key fields: `일자`, `판매-수량`, `판매-금액`, `판매-판매합계`, `손익-판매금액`, `손익-마진`, `손익-마진율`.
   - Used for: today sales, yesterday sales, day-over-day change, monthly cumulative trend.

2. `Statistics_Good_YYYY-MM-DD.xls`
   - Purpose: product-level sales statistics.
   - The sample `.xls` is actually an HTML table export encoded as Korean text.
   - Key fields: `상품코드`, `상품명`, `판매-수량`, `판매-금액`, `손익-판매금액`, `손익-마진`, CS fields.
   - Used for: product TOP list, product-level CS impact, product detail review.

3. `매출처별.xlsx`
   - Purpose: seller/customer-level monthly sales statistics.
   - Key fields: `판매사`, `판매-수량`, `판매-금액`, `손익-판매금액`, `손익-마진`, CS fields.
   - Used for: seller TOP list, seller-level CS impact, seller detail review.

## Upload Model

The structure should support two modes.

### Full Upload

Full upload seeds or rebuilds the sales dashboard baseline.

- Admin-only.
- Used when starting the system or replacing the whole month/history.
- Accepts all available file types for a period.
- Can clear and rebuild the sales report tables for the selected period after confirmation.
- Must show a warning before replacing existing data.

### 기준일 Upload

Daily operation uses 기준일 uploads.

- The user uploads the latest files for a 기준일.
- Files are stored as snapshots by `report_date` and `report_type`.
- If the same `report_date + report_type` already exists, show a duplicate warning and allow cancel or replace.
- The date-by-date file may contain the whole month-to-date range. That is acceptable; the app should upsert all rows included in that file.
- Product and seller files are also treated as 기준일 snapshots, even when they represent month-to-date totals as of that date.

This means the user's understanding is correct: upload the full dataset first, then continue with 기준일 data. The app should not assume every daily file contains only one day's delta.

## Dashboard Layout

Use the approved A-style operational layout.

Top KPI cards:

- Today 손익 매출.
- Yesterday 손익 매출.
- Day-over-day amount and percentage.
- Month-to-date 손익 매출.
- Seller-file total.
- File consistency check.

Main body:

- Left: date-by-date sales trend table.
- Right: seller/customer TOP table.
- Lower left: product TOP table.
- Lower right: review-needed table.

Filters:

- Period.
- 기준일.
- Seller/product search.
- Review status.

Actions:

- 날짜별 업로드.
- 상품별 업로드.
- 매출처별 업로드.
- Excel download.
- Refresh.

## File Consistency Check

The dashboard should compare totals across file types.

Sample finding:

- Date/product monthly 손익 매출: `133,210,410`.
- Seller monthly 손익 매출: `133,197,960`.
- Difference: `12,450`.

When the difference is non-zero, the dashboard should show a warning card and list the compared totals. This is a safety feature, not a hard failure.

## Permissions

- Viewing the sales dashboard and uploading sales report files uses `sales_report_manage`.
- Full upload or period rebuild should require admin role plus `sales_report_manage`.
- 기준일 upload can use `sales_report_manage`.

## Error Handling

- Unknown file type: reject with a clear message listing supported file types.
- Missing required columns: show the missing column names before saving.
- Duplicate 기준일/report type: warn before replacement.
- File parse failure: preserve the uploaded original only if useful for debugging, but do not create dashboard rows.
- Total mismatch: save the data but show a review warning.

## Testing

Add tests for:

- File type detection for the three samples.
- Parsing `.xlsx` and HTML-table `.xls`.
- Full upload versus 기준일 upload permissions.
- Duplicate detection.
- Dashboard API returns today/yesterday/month totals.
- File consistency warning appears when totals differ.
- Sidebar navigation keeps `매출현황 및 관리` separated from admin settings.

## Non-Goals

- Do not build a public analytics product.
- Do not add charts in the first implementation unless the table layout is already stable.
- Do not require 매출처별 data to calculate today/yesterday, because that comes from the date-by-date file.
- Do not assume daily uploads are delta-only files.
