# Workhub Current Status

## Local Host

- URL: http://127.0.0.1:8776/login
- Host: 127.0.0.1
- Port: 8776

## Workspace

- Local folder: `C:\Users\ssh19\OneDrive\Documents\Codex\soillbridgecorp`
- GitHub: https://github.com/master-shinsunghwan/soillbridgecorp
- Current commit: `136515d Add supply format management template support`

## Run Locally

```powershell
$env:WORKHUB_PORT = "8776"
& "C:\Users\ssh19\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" scripts\workhub_delivery_app.py
```

Open:

```text
http://127.0.0.1:8776/login
```

## Latest Work

- Added `통합관리대장 양식` download option.
- Saved the supplied `Supply` Excel format as the management ledger export template.
- Removed real row data from the saved template and kept only the blank format/style.
- Added support for the same format in individual delivery text summary.
- Preserved extra columns: `주문상품고유번호`, `상품코드`, `주문번호`, `고객선택옵션`.
