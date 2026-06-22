#!/usr/bin/env bash
set -euo pipefail

ENV_FILE="${WORKHUB_ENV_FILE:-/opt/workhub/.env}"
if [[ -f "$ENV_FILE" ]]; then
  set -a
  # shellcheck disable=SC1090
  source "$ENV_FILE"
  set +a
fi

DATA_DIR="${WORKHUB_DATA_DIR:-/opt/workhub/data}"
BACKUP_DIR="${WORKHUB_BACKUP_DIR:-/opt/workhub/backups}"
RETENTION_DAYS="${WORKHUB_BACKUP_RETENTION_DAYS:-90}"
TIMESTAMP="$(date +%Y-%m-%d_%H%M%S)"
BACKUP_NAME="workhub_backup_${TIMESTAMP}.tar.gz"
DB_FILE="${DATA_DIR}/config/workhub.db"
TMP_DIR="$(mktemp -d)"

cleanup() {
  rm -rf "$TMP_DIR"
}
trap cleanup EXIT

mkdir -p "$BACKUP_DIR"

if [[ -f "$DB_FILE" ]] && command -v sqlite3 >/dev/null 2>&1; then
  mkdir -p "$TMP_DIR/config"
  sqlite3 "$DB_FILE" ".backup '${TMP_DIR}/config/workhub.db'"
  tar --exclude="./config/workhub.db" -czf "${BACKUP_DIR}/${BACKUP_NAME}" -C "$DATA_DIR" . -C "$TMP_DIR" config/workhub.db
else
  tar -czf "${BACKUP_DIR}/${BACKUP_NAME}" -C "$DATA_DIR" .
fi

find "$BACKUP_DIR" -type f -name "workhub_backup_*.tar.gz" -mtime +"$RETENTION_DAYS" -delete

echo "${BACKUP_DIR}/${BACKUP_NAME}"
