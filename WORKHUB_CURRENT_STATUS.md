# Workhub Current Status

## Local Host

- URL: `http://127.0.0.1:8770/`
- Login URL: `http://127.0.0.1:8770/login`
- Port: `8770`

## Workspace

- Local folder: `C:\Users\ssh19\OneDrive\Documents\Codex\soillbridgecorp`
- GitHub: `https://github.com/master-shinsunghwan/soillbridgecorp`
- Branch: `main`
- App verification commit: `826c3b8 Fix Workhub lucide icon import`

## Run Locally

```powershell
.\업무허브 실행.cmd
```

또는:

```powershell
python -m pip install -r requirements.txt
python scripts\workhub_delivery_app.py 8770
```

## Login

- Admin: `admin / admin1234`

## Latest Verified Work

- Restored and stabilized the other-PC Workhub transfer.
- Fixed the login-after-screen freeze caused by a missing lucide `ArrowRight` export.
- Updated the sales KPI card to show `매입처별 총합계 금액`.
- Verified login and sidebar navigation in the browser.
- Full test suite passed: `55 tests OK`.

## Handoff

- Handoff note: `WORKHUB_HANDOFF_20260622.md`
- Transfer package: include `config/workhub.db`; for other-PC checks, use the portable ZIP that also includes `runtime/python`.
- Exclude `.git`, `node_modules`, caches, and sensitive mail/token files.
