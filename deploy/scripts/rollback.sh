#!/usr/bin/env bash
set -euo pipefail

if [[ $# -ne 1 ]]; then
  echo "Usage: $0 <commit_hash>" >&2
  exit 2
fi

TARGET_COMMIT="$1"
APP_DIR="${WORKHUB_APP_DIR:-/opt/soillbridgecorp}"
BACKUP_SCRIPT="${WORKHUB_BACKUP_SCRIPT:-${APP_DIR}/deploy/scripts/backup.sh}"

cd "$APP_DIR"

if [[ -x "$BACKUP_SCRIPT" ]]; then
  "$BACKUP_SCRIPT"
fi

git fetch --all --prune
git reset --hard "$TARGET_COMMIT"

source .venv/bin/activate
pip install -r requirements.txt

if [[ -f "package-lock.json" ]] && command -v npm >/dev/null 2>&1; then
  npm ci
  npm run build:css
  npm run build
fi

sudo systemctl restart workhub
sudo journalctl -u workhub -n 100 --no-pager
