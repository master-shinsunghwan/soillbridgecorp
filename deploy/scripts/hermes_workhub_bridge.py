#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import subprocess
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any


HERMES_BIN = os.environ.get("HERMES_BIN", "/opt/hermes/.venv/bin/hermes")
HERMES_CWD = os.environ.get("HERMES_CWD", "/opt/hermes")
BRIDGE_HOST = os.environ.get("HERMES_BRIDGE_HOST", "0.0.0.0")
BRIDGE_PORT = int(os.environ.get("HERMES_BRIDGE_PORT", "4871"))
BRIDGE_TOKEN = os.environ.get("WORKHUB_HERMES_BRIDGE_TOKEN", "").strip()
REQUEST_TIMEOUT = int(os.environ.get("HERMES_BRIDGE_TIMEOUT", "180"))


def normalize_token(value: str) -> str:
    value = (value or "").strip()
    if value.lower().startswith("bearer "):
        return value.split(" ", 1)[1].strip()
    return value


def is_authorized(headers: dict[str, str], expected_token: str = BRIDGE_TOKEN) -> bool:
    expected_token = normalize_token(expected_token)
    if not expected_token:
        return True
    normalized = {str(key).lower(): str(value) for key, value in headers.items()}
    candidates = [
        normalized.get("authorization", ""),
        normalized.get("x-hermes-api-key", ""),
    ]
    return any(normalize_token(candidate) == expected_token for candidate in candidates)


def build_prompt(payload: dict[str, Any], mode: str) -> str:
    if mode == "automation":
        title = str(payload.get("title") or "Workhub automation request").strip()
        body = str(payload.get("body") or payload.get("message") or "").strip()
        return (
            "You are Hermes connected to Soillbridge Workhub.\n"
            "Reply in Korean. Convert the request into actionable steps, risks, and next actions.\n\n"
            f"Title: {title}\n"
            f"Request:\n{body}"
        )
    message = str(payload.get("message") or payload.get("prompt") or "").strip()
    return (
        "You are Hermes connected to Soillbridge Workhub.\n"
        "Reply in Korean with a concise, practical business answer.\n"
        "When the request is ambiguous, list what should be checked next.\n\n"
        f"Workhub message:\n{message}"
    )


def run_hermes(prompt: str) -> str:
    if not prompt.strip():
        raise ValueError("message is required")
    env = os.environ.copy()
    env.setdefault("HERMES_ACCEPT_HOOKS", "1")
    completed = subprocess.run(
        [HERMES_BIN, "-z", prompt],
        cwd=HERMES_CWD,
        env=env,
        text=True,
        capture_output=True,
        timeout=REQUEST_TIMEOUT,
        check=False,
    )
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "").strip()
        raise RuntimeError(detail[-1200:] or f"hermes exited with {completed.returncode}")
    return completed.stdout.strip()


class WorkhubHermesBridgeHandler(BaseHTTPRequestHandler):
    server_version = "WorkhubHermesBridge/1.0"

    def log_message(self, fmt: str, *args: Any) -> None:
        return

    def send_json(self, status: int, payload: dict[str, Any]) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def read_payload(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length") or "0")
        if length <= 0:
            return {}
        raw = self.rfile.read(length).decode("utf-8", errors="replace")
        return json.loads(raw or "{}")

    def do_GET(self) -> None:
        if self.path.rstrip("/") == "/health":
            self.send_json(200, {"ok": True, "status": "ready"})
            return
        self.send_json(404, {"ok": False, "error": "not_found"})

    def do_POST(self) -> None:
        if not is_authorized(dict(self.headers)):
            self.send_json(401, {"ok": False, "error": "unauthorized"})
            return
        if self.path not in ("/api/chat", "/api/automation"):
            self.send_json(404, {"ok": False, "error": "not_found"})
            return
        mode = "automation" if self.path == "/api/automation" else "chat"
        try:
            payload = self.read_payload()
            prompt = build_prompt(payload, mode)
            answer = run_hermes(prompt)
            self.send_json(200, {"ok": True, "answer": answer, "text": answer})
        except subprocess.TimeoutExpired:
            self.send_json(504, {"ok": False, "error": "timeout"})
        except Exception as exc:
            self.send_json(500, {"ok": False, "error": str(exc)})


def main() -> None:
    httpd = ThreadingHTTPServer((BRIDGE_HOST, BRIDGE_PORT), WorkhubHermesBridgeHandler)
    print(f"Workhub Hermes bridge listening on {BRIDGE_HOST}:{BRIDGE_PORT}", flush=True)
    httpd.serve_forever()


if __name__ == "__main__":
    main()
