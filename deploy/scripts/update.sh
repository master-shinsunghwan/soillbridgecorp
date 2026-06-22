#!/usr/bin/env bash
set -euo pipefail

APP_DIR="${WORKHUB_APP_DIR:-/opt/soillbridgecorp}"
BACKUP_SCRIPT="${WORKHUB_BACKUP_SCRIPT:-${APP_DIR}/deploy/scripts/backup.sh}"

cd "$APP_DIR"

if [[ -x "$BACKUP_SCRIPT" ]]; then
  "$BACKUP_SCRIPT"
fi

git pull --ff-only

if [[ ! -d ".venv" ]]; then
  python3 -m venv .venv
fi

source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

if [[ -f "package-lock.json" ]] && command -v npm >/dev/null 2>&1; then
  npm ci
  npm run build:css
  npm run build
fi

sudo systemctl restart workhub
sudo systemctl status workhub --no-pager
