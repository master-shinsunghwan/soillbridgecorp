#!/usr/bin/env bash
set -euo pipefail

BRIDGE_SOURCE="${BRIDGE_SOURCE:-/opt/soilbridgecorp/deploy/scripts/hermes_workhub_bridge.py}"
TOKEN_FILE="${TOKEN_FILE:-/opt/company-erp/workhub-hermes-bridge.token}"
CONTAINER_PATTERN="${CONTAINER_PATTERN:-^hermes-agent}"
PORT="${HERMES_BRIDGE_PORT:-4871}"
OPENAI_ENV_FILE="${OPENAI_ENV_FILE:-/opt/company-erp/workhub-openai.env}"

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
docker cp "$TOKEN_FILE" "$HERMES_CONTAINER:/tmp/workhub-hermes-bridge.token" >/dev/null
docker exec "$HERMES_CONTAINER" python3 -c 'import os, signal
for pid in os.listdir("/proc"):
    if not pid.isdigit():
        continue
    try:
        raw = open(f"/proc/{pid}/cmdline", "rb").read().replace(b"\0", b" ").decode("utf-8", "replace").strip()
    except OSError:
        continue
    if raw.startswith("python3 /tmp/workhub-hermes-bridge.py"):
        os.kill(int(pid), signal.SIGTERM)
' >/dev/null

OPENAI_EXEC_ENV=()
if [ -f "$OPENAI_ENV_FILE" ]; then
  set -a
  # shellcheck disable=SC1090
  . "$OPENAI_ENV_FILE"
  set +a
fi
WORKHUB_AI_TOOL_PROVIDER="codex"
HERMES_PROVIDER="openai-codex"
HERMES_MODEL="${HERMES_MODEL:-gpt-5.5}"
if [ -n "${OPENAI_API_KEY:-}" ]; then
  OPENAI_EXEC_ENV+=(-e "OPENAI_API_KEY=$OPENAI_API_KEY")
fi
if [ -n "${OPENAI_TEXT_MODEL:-}" ]; then
  OPENAI_EXEC_ENV+=(-e "OPENAI_TEXT_MODEL=$OPENAI_TEXT_MODEL")
fi
if [ -n "${OPENAI_IMAGE_MODEL:-}" ]; then
  OPENAI_EXEC_ENV+=(-e "OPENAI_IMAGE_MODEL=$OPENAI_IMAGE_MODEL")
fi
if [ -n "${FAL_KEY:-}" ]; then
  OPENAI_EXEC_ENV+=(-e "FAL_KEY=$FAL_KEY")
fi
OPENAI_EXEC_ENV+=(-e "WORKHUB_AI_TOOL_PROVIDER=$WORKHUB_AI_TOOL_PROVIDER")
OPENAI_EXEC_ENV+=(-e "HERMES_PROVIDER=$HERMES_PROVIDER")
OPENAI_EXEC_ENV+=(-e "HERMES_MODEL=$HERMES_MODEL")

docker exec "$HERMES_CONTAINER" python3 - <<'PY'
from pathlib import Path

import yaml

path = Path("/opt/data/config.yaml")
config = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
image_gen = config.setdefault("image_gen", {})
image_gen["provider"] = "openai-codex"
image_gen.setdefault("model", "gpt-image-2-medium")
openai_codex = image_gen.setdefault("openai-codex", {})
openai_codex.setdefault("model", "gpt-image-2-medium")
path.write_text(yaml.safe_dump(config, sort_keys=False, allow_unicode=True), encoding="utf-8")
PY

exec docker exec \
  -e "HERMES_BRIDGE_PORT=$PORT" \
  -e "HERMES_BRIDGE_TIMEOUT=${HERMES_BRIDGE_TIMEOUT:-240}" \
  "${OPENAI_EXEC_ENV[@]}" \
  "$HERMES_CONTAINER" \
  sh -lc 'export WORKHUB_HERMES_BRIDGE_TOKEN="$(cat /tmp/workhub-hermes-bridge.token)"; cd /opt/hermes && exec python3 /tmp/workhub-hermes-bridge.py'
