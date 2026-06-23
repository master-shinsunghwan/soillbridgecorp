#!/usr/bin/env bash
set -euo pipefail

BRIDGE_SOURCE="${BRIDGE_SOURCE:-/opt/soilbridgecorp/deploy/scripts/hermes_workhub_bridge.py}"
TOKEN_FILE="${TOKEN_FILE:-/opt/company-erp/workhub-hermes-bridge.token}"
CONTAINER_PATTERN="${CONTAINER_PATTERN:-^hermes-agent}"
PORT="${HERMES_BRIDGE_PORT:-4871}"

if [ ! -s "$TOKEN_FILE" ]; then
  umask 077
  openssl rand -hex 32 > "$TOKEN_FILE"
fi

while true; do
  HERMES_CONTAINER="$(docker ps --format '{{.Names}}' | grep -E "$CONTAINER_PATTERN" | head -1 || true)"
  if [ -n "$HERMES_CONTAINER" ]; then
    break
  fi
  sleep 5
done

docker cp "$BRIDGE_SOURCE" "$HERMES_CONTAINER:/tmp/workhub-hermes-bridge.py" >/dev/null

exec docker exec \
  -e "WORKHUB_HERMES_BRIDGE_TOKEN=$(cat "$TOKEN_FILE")" \
  -e "HERMES_BRIDGE_PORT=$PORT" \
  -e "HERMES_BRIDGE_TIMEOUT=${HERMES_BRIDGE_TIMEOUT:-180}" \
  "$HERMES_CONTAINER" \
  sh -lc 'cd /opt/hermes && python3 /tmp/workhub-hermes-bridge.py'
