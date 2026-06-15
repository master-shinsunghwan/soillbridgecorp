# Workhub Project Guidance

## Persistence Workflow

- Use this Git checkout as the working folder: `C:\Users\ssh19\OneDrive\Documents\Codex\soillbridgecorp`.
- After meaningful code changes, save in both places:
  1. Local/OneDrive: keep files in this checkout.
  2. GitHub: run verification, commit, and push to `origin/main` unless the user asks for a branch or PR.
- Before committing UI/build changes, run:
  - `npm run build:css`
  - `python tests/test_daisyui_foundation.py`
  - `python -m py_compile scripts/workhub_delivery_app.py`
- Do not commit runtime data, local DB files, `node_modules`, or NAS/local test data.
