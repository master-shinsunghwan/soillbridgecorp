# Workhub Project Handoff - 2026-07-15

## Project

- Repository: `https://github.com/master-shinsunghwan/soillbridgecorp.git`
- Local working folder: `C:\Users\ssh19\OneDrive\Documents\Codex\soillbridgecorp`
- Main app file: `scripts\workhub_delivery_app.py`
- Desktop wrapper file: `scripts\workhub_vps_desktop_app.py`
- VPS URL: `https://workhub.soilbridgecorp.cloud/`
- Current branch: `main`
- Latest checked commit at handoff:
  - `e255378 Rename import schedule due date label`

## Current Git Status

At handoff time, local `main` is aligned with `origin/main` and there are no uncommitted changes.

Recent relevant commits:

- `e255378` - Changed import schedule label from `입고예정일` to `컨테이너 하역 예정일`.
- `34999a6` - Added chunked desktop Excel download support.
- `10b05e5` - Added initial desktop app Excel download bridge.
- `4cd8ca5` - Persisted import cost report downloads.
- `451fa87` - Added import cost handoff for completed containers.

## Latest Desktop Installer

Use this ZIP when installing the local Windows desktop app:

```text
C:\Users\ssh19\OneDrive\바탕 화면\SoilbridgeWorkhub_Desktop_20260715_082929.zip
```

Backup copy in repository output folder:

```text
C:\Users\ssh19\OneDrive\Documents\Codex\soillbridgecorp\output\SoilbridgeWorkhub_Desktop_20260715_082929.zip
```

Install steps:

1. Extract the ZIP.
2. Run `Install.cmd`.
3. Use the desktop shortcut for `(주)소일브릿지 업무자동화`.

Important:

- Existing old EXE files do not automatically receive wrapper fixes.
- The Excel download fix is inside the desktop wrapper, so PCs with the old EXE must reinstall using the latest ZIP above.
- Excel files should save to the Windows Downloads folder when using the latest desktop app.

## Local Run

Recommended local run command:

```powershell
cd "C:\Users\ssh19\OneDrive\Documents\Codex\soillbridgecorp"
python scripts\workhub_delivery_app.py 8781
```

Open:

```text
http://127.0.0.1:8781/
```

If port `8781` is busy, use another port such as `8782` or `8783`.

## VPS Deploy Flow

Do not deploy automatically unless requested by the user.

Normal deploy expectation:

1. Verify locally first.
2. Commit and push to GitHub.
3. Pull/rebuild/restart on VPS.
4. Verify `https://workhub.soilbridgecorp.cloud/`.

Typical local checks:

```powershell
python -m py_compile scripts\workhub_delivery_app.py scripts\workhub_vps_desktop_app.py
python -m unittest tests.test_workhub_desktop_app tests.test_workhub_app_feature_parity
```

For focused sales/import-cost work, also run related tests if edited:

```powershell
python -m unittest tests.test_sales_report_uploads
python -m unittest tests.test_import_costs
```

## Important Product Areas

### 통합관리대장 / CS처리대장

- Main ledger and CS ledger are both handled in `scripts\workhub_delivery_app.py`.
- Recent requirements:
  - Double-click or `F2` should activate cell editing.
  - Arrow keys should move between cells, not scroll the page.
  - Filter options should behave like Excel cascading filters.
  - Color-filter support was requested for 통합관리대장.
  - CS completed rows should always show the completed highlight color.

### 매출현황 및 관리

- Daily upload direction:
  - `YYYYMMDD 매입처별 매출현황.xls`
  - `YYYYMMDD 상품별 매출현황.xls`
  - `YYYYMMDD 채널별 매출현황.xls`
- Date should be parsed from the filename prefix when uploaded later than the actual sales date.
- Recent 7-day chart should show the latest 7 sales dates, even across month boundaries.
- July view should use July monthly totals when July data exists.

### 수입 원가 계산

This is the current high-priority debugging area.

Known user concerns:

- Program-read values and manually checked values differed.
- `관세` was previously being misread as `-61원`.
- `D/O 운임`, `통관수수료`, `수입부가세`, `수수료 부가세` must be visible and included correctly.
- Existing saved calculations must be recalculated/redisplayed accurately when source data or formulas are fixed.
- Saved data list should be shown as one horizontal, easy-to-scan row per record.
- Sort saved data by `수입원장 기준 날짜`.
- Add row number.
- Always show `우리가 관리하는 품명`.
- Long saved lists must scroll.
- Report Excel output had download issues in the installed desktop app; latest wrapper ZIP above contains the download fix, but user confirmation on a real PC flow is still needed.

### 수출입 업무 및 화물 입출고 관리

- User requested a delete button for schedules because import/export or cargo schedules can be cancelled.
- Label change completed:
  - `입고예정일` -> `컨테이너 하역 예정일`

## Desktop Download Fix Summary

The desktop app download issue was handled in two layers:

- Server/page side:
  - `downloadWorkbookResponse`
  - `desktopDownloadBridgeApi`
  - chunked `saveBlobThroughDesktopBridge`
- Desktop wrapper side:
  - `beginDownload`
  - `appendDownloadChunk`
  - `finishDownload`
  - `cancelDownload`

Files:

```text
scripts\workhub_delivery_app.py
scripts\workhub_vps_desktop_app.py
tests\test_workhub_desktop_app.py
tests\test_workhub_app_feature_parity.py
```

Key warning:

- If a PC still runs an older installed EXE, Excel downloads may still fail.
- Reinstall using `SoilbridgeWorkhub_Desktop_20260715_082929.zip`.

## Secrets / Credentials

Do not place passwords, API keys, tokens, Google Drive rclone tokens, or VPS private keys in this handoff file.

Use existing local/VPS secure files or ask the project owner for credentials when needed.

## Suggested Next Debugging Order

1. Reproduce current 수입 원가 계산 discrepancy with the same HBL shown by the user.
2. Inspect parser output before UI formatting.
3. Verify each fee bucket:
   - D/O 운임
   - 통관수수료
   - 관세
   - 수입부가세
   - 수수료 부가세
   - 기타 비용
4. Confirm whether tax checkboxes include/exclude values only from product unit cost, not from source totals.
5. Recalculate existing saved records after formula fix.
6. Confirm report Excel download in:
   - Normal Chrome/browser
   - Latest installed desktop EXE
7. Commit, push, and deploy only after local verification.

## Quick Start For New Codex Task

Paste this into the new project/task:

```text
Workhub 작업을 이어서 진행합니다.
저장소: C:\Users\ssh19\OneDrive\Documents\Codex\soillbridgecorp
주요 파일: scripts\workhub_delivery_app.py
인계 파일: WORKHUB_HANDOFF_20260715.md
우선순위: 수입 원가 계산 디버깅, 저장 데이터 UI 정리, 설치형 EXE 엑셀 다운로드 확인
먼저 git status, 최근 커밋, 로컬 실행 상태를 확인하고 진행해주세요.
VPS 배포는 로컬 테스트 후 승인받고 진행해주세요.
```
