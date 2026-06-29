#!/usr/bin/env python3
from __future__ import annotations

import json
import base64
import mimetypes
import os
import re
import subprocess
import urllib.error
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any


HERMES_BIN = os.environ.get("HERMES_BIN", "/opt/hermes/.venv/bin/hermes")
HERMES_CWD = os.environ.get("HERMES_CWD", "/opt/hermes")
BRIDGE_HOST = os.environ.get("HERMES_BRIDGE_HOST", "0.0.0.0")
BRIDGE_PORT = int(os.environ.get("HERMES_BRIDGE_PORT", "4871"))
BRIDGE_TOKEN = os.environ.get("WORKHUB_HERMES_BRIDGE_TOKEN", "").strip()
REQUEST_TIMEOUT = int(os.environ.get("HERMES_BRIDGE_TIMEOUT", "240"))
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
OPENAI_TEXT_MODEL = os.environ.get("OPENAI_TEXT_MODEL", "gpt-5.5")
OPENAI_IMAGE_MODEL = os.environ.get("OPENAI_IMAGE_MODEL", "gpt-image-2")
FAL_KEY = os.environ.get("FAL_KEY", "").strip()
AI_TOOL_PROVIDER = os.environ.get("WORKHUB_AI_TOOL_PROVIDER", "codex").strip().lower()
HERMES_PROVIDER = os.environ.get("HERMES_PROVIDER", "openai-codex").strip()
HERMES_MODEL = os.environ.get("HERMES_MODEL", "gpt-5.5").strip()
HERMES_IMAGE_CACHE_DIR = Path(os.environ.get("HERMES_IMAGE_CACHE_DIR", "/opt/data/cache/images")).resolve()


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
    workhub_context = payload.get("workhub_context")
    context_text = ""
    scope_text = ""
    if isinstance(workhub_context, dict) and workhub_context:
        context_text = "\n\nWorkhub context snapshot:\n" + json.dumps(workhub_context, ensure_ascii=False, indent=2)[:6000]
        sales_report = workhub_context.get("sales_report")
        if isinstance(sales_report, dict) and sales_report.get("scope") == "today":
            scope_text = (
                "\n\nSales data rule: the user asked for today's/current-day sales. "
                "Use only Workhub context sales_report.today for sales figures. "
                "Do not infer or mention monthly cumulative totals unless the user explicitly asks for month/cumulative data."
            )
    capabilities = payload.get("capabilities")
    capability_text = ""
    if isinstance(capabilities, dict) and capabilities:
        capability_text = "\n\nAvailable Workhub AI capabilities:\n" + json.dumps(capabilities, ensure_ascii=False, indent=2)[:2000]
    intent = str(payload.get("intent") or "").strip().lower()
    chat_mode = str(payload.get("mode") or payload.get("requested_mode") or "auto").strip().lower()
    intent_text = ""
    if intent == "web_search":
        intent_text = "\n\nRequested tool intent: web_search. Use the shared Hermes research/search backend when available."
    elif intent == "image_generation":
        intent_text = "\n\nRequested tool intent: image_generation. Use the shared Hermes image-generation backend when available."
    mode_text = {
        "automation": (
            "\n\nRequested Workhub mode: 업무자동화. Prioritize Workhub business data, operational workflows, "
            "and actionable automation steps. You may still answer generally if the request needs it."
        ),
        "general": (
            "\n\nRequested Workhub mode: 일반 AI. Answer like a general Codex/GPT assistant. "
            "Do not force the response into Workhub automation or internal database actions unless the user explicitly asks."
        ),
        "search": (
            "\n\nRequested Workhub mode: 자료검색. Treat this as a research/search task. "
            "Prefer current, source-aware investigation and summarize findings with links when available."
        ),
        "image": (
            "\n\nRequested Workhub mode: 이미지생성. Treat this as an image generation task. "
            "Create or route to image-generation tools and return the generated file path or image payload when possible."
        ),
        "auto": (
            "\n\nRequested Workhub mode: 자동선택. Choose the most useful capability among Workhub automation, general AI, search, and image generation."
        ),
    }.get(chat_mode, "")
    if mode == "automation":
        title = str(payload.get("title") or "Workhub automation request").strip()
        body = str(payload.get("body") or payload.get("message") or "").strip()
        return (
            "You are Hermes connected to Soillbridge Workhub.\n"
            "Reply in Korean. Convert the request into actionable steps, risks, and next actions.\n"
            "Use the provided Workhub context when it helps, but do not modify Workhub data.\n"
            "If the user asks for code or UI implementation, summarize the request and say Codex/developer work is needed.\n\n"
            f"Title: {title}\n"
            f"Request:\n{body}"
            f"{context_text}{scope_text}{capability_text}{mode_text}{intent_text}"
        )
    message = str(payload.get("message") or payload.get("prompt") or "").strip()
    return (
        "You are Hermes connected to Soillbridge Workhub.\n"
        "Reply in Korean with a concise, practical answer.\n"
        "When the request is ambiguous, list what should be checked next.\n\n"
        "Use the provided Workhub context when it helps, but do not modify Workhub data.\n"
        "Do not limit the user to Workhub-only tasks: you may support general AI answers, research/search, image generation, and Workhub automation according to the requested mode.\n"
        "If the user asks for code or UI implementation, summarize the request and say Codex/developer work is needed.\n\n"
        f"Workhub message:\n{message}"
        f"{context_text}{scope_text}{capability_text}{mode_text}{intent_text}"
    )


def hermes_command(prompt: str) -> list[str]:
    command = [HERMES_BIN, "--ignore-rules"]
    if HERMES_PROVIDER:
        command.extend(["--provider", HERMES_PROVIDER])
    if HERMES_MODEL:
        command.extend(["-m", HERMES_MODEL])
    command.extend(["-z", prompt])
    return command


def run_hermes(prompt: str) -> str:
    if not prompt.strip():
        raise ValueError("message is required")
    env = os.environ.copy()
    env.setdefault("HERMES_ACCEPT_HOOKS", "1")
    completed = subprocess.run(
        hermes_command(prompt),
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


def generated_image_payload(answer: str) -> dict[str, Any]:
    matches = re.findall(r"((?:[A-Za-z]:)?[/\\][^\s`]+?\.(?:png|jpg|jpeg|webp))", answer or "", flags=re.IGNORECASE)
    for raw_path in reversed(matches):
        path = Path(raw_path).resolve()
        try:
            if HERMES_IMAGE_CACHE_DIR not in path.parents:
                continue
            if not path.is_file():
                continue
            image_bytes = path.read_bytes()
        except OSError:
            continue
        image_mime = mimetypes.guess_type(path.name)[0] or "image/png"
        return {
            "image_base64": base64.b64encode(image_bytes).decode("ascii"),
            "image_mime": image_mime,
            "image_path": str(path),
        }
    return {}


def openai_request(path: str, payload: dict[str, Any]) -> dict[str, Any]:
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY is not configured on the Hermes bridge server")
    request = urllib.request.Request(
        f"https://api.openai.com{path}",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {OPENAI_API_KEY}",
            "Content-Type": "application/json; charset=utf-8",
            "Accept": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT) as response:
            body = response.read().decode("utf-8", errors="replace")
            return json.loads(body or "{}")
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"OpenAI API error {exc.code}: {body[:800]}") from exc


def response_output_text(payload: dict[str, Any]) -> str:
    if payload.get("output_text"):
        return str(payload["output_text"])
    parts: list[str] = []
    for item in payload.get("output", []) if isinstance(payload.get("output"), list) else []:
        for content in item.get("content", []) if isinstance(item, dict) else []:
            if not isinstance(content, dict):
                continue
            text = content.get("text") or content.get("output_text")
            if text:
                parts.append(str(text))
    return "\n".join(parts).strip()


def run_openai_web_search(payload: dict[str, Any]) -> dict[str, Any]:
    message = str(payload.get("message") or payload.get("prompt") or "").strip()
    if not message:
        raise ValueError("message is required")
    context = payload.get("workhub_context")
    input_text = message
    if isinstance(context, dict) and context:
        input_text += "\n\nWorkhub context snapshot:\n" + json.dumps(context, ensure_ascii=False, indent=2)[:4000]
    result = openai_request("/v1/responses", {
        "model": OPENAI_TEXT_MODEL,
        "input": input_text,
        "tools": [{"type": "web_search"}],
    })
    answer = response_output_text(result) or json.dumps(result, ensure_ascii=False)[:2000]
    return {"ok": True, "answer": answer, "text": answer, "provider": "openai", "capability": "web_search"}


def run_openai_image_generation(payload: dict[str, Any]) -> dict[str, Any]:
    prompt = str(payload.get("message") or payload.get("prompt") or "").strip()
    if not prompt:
        raise ValueError("message is required")
    result = openai_request("/v1/images/generations", {
        "model": OPENAI_IMAGE_MODEL,
        "prompt": prompt,
        "size": str(payload.get("size") or "1024x1024"),
        "n": 1,
    })
    data = result.get("data") if isinstance(result.get("data"), list) else []
    first = data[0] if data else {}
    image_base64 = first.get("b64_json") if isinstance(first, dict) else ""
    image_url = first.get("url") if isinstance(first, dict) else ""
    if not image_base64 and not image_url:
        raise RuntimeError("OpenAI image response did not include an image")
    answer = "이미지를 생성했습니다."
    return {
        "ok": True,
        "answer": answer,
        "text": answer,
        "provider": "openai",
        "capability": "image_generation",
        "image_base64": image_base64,
        "image_url": image_url,
        "image_mime": "image/png",
    }


def should_use_openai_for_intent(intent: str) -> bool:
    if intent not in {"web_search", "image_generation"}:
        return False
    if AI_TOOL_PROVIDER in {"codex", "openai-codex", "hermes", "shared", "slack", "slack_hermes"}:
        return False
    if AI_TOOL_PROVIDER in {"openai", "api"}:
        return True
    return False


def should_block_unconfigured_image_generation(intent: str) -> bool:
    if intent != "image_generation":
        return False
    if AI_TOOL_PROVIDER in {"codex", "openai-codex"}:
        return False
    if should_use_openai_for_intent(intent):
        return False
    return not FAL_KEY


def image_generation_not_configured_response() -> dict[str, Any]:
    openai_note = (
        "OpenAI 이미지 API를 쓰려면 OPENAI_API_KEY와 WORKHUB_AI_TOOL_PROVIDER=openai를 설정해야 합니다."
        if OPENAI_API_KEY
        else "OpenAI 이미지 API로 전환할 OPENAI_API_KEY도 없습니다."
    )
    answer = (
        "이미지 생성 설정이 아직 연결되지 않았습니다.\n\n"
        f"VPS의 Hermes 이미지 생성 백엔드에 FAL_KEY가 없습니다. {openai_note}\n"
        "서버 /opt/company-erp/workhub-openai.env에 FAL_KEY를 설정한 뒤 workhub-hermes-bridge를 재시작하면 Workhub에서 바로 이미지 생성 요청을 처리할 수 있습니다."
    )
    return {
        "ok": False,
        "answer": answer,
        "text": answer,
        "provider": "hermes",
        "capability": "image_generation",
        "missing_config": ["FAL_KEY"],
    }


def requested_intent(payload: dict[str, Any], mode: str) -> str:
    explicit = str(payload.get("intent") or "").strip().lower()
    if explicit in {"web_search", "image_generation"}:
        return explicit
    chat_mode = str(payload.get("mode") or payload.get("requested_mode") or "").strip().lower()
    if chat_mode == "search":
        return "web_search"
    if chat_mode == "image":
        return "image_generation"
    if mode != "chat":
        return ""
    message = str(payload.get("message") or payload.get("prompt") or "").lower()
    if message.startswith(("/검색", "검색:", "search:")) or any(word in message for word in ("검색해", "찾아봐", "최신", "뉴스")):
        return "web_search"
    if message.startswith(("/이미지", "이미지:", "image:")) or any(word in message for word in ("이미지 만들어", "그림 만들어", "로고 만들어", "배너 만들어")):
        return "image_generation"
    return ""


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
            intent = requested_intent(payload, mode)
            if intent == "web_search" and should_use_openai_for_intent(intent):
                self.send_json(200, run_openai_web_search(payload))
                return
            if intent == "image_generation" and should_use_openai_for_intent(intent):
                self.send_json(200, run_openai_image_generation(payload))
                return
            if should_block_unconfigured_image_generation(intent):
                self.send_json(200, image_generation_not_configured_response())
                return
            prompt = build_prompt(payload, mode)
            answer = run_hermes(prompt)
            response = {"ok": True, "answer": answer, "text": answer}
            if intent == "image_generation":
                response.update(generated_image_payload(answer))
            self.send_json(200, response)
        except subprocess.TimeoutExpired:
            self.send_json(504, {"ok": False, "error": "timeout"})
        except Exception as exc:
            self.send_json(500, {"ok": False, "error": str(exc)})


class ReusableThreadingHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = True


def main() -> None:
    httpd = ReusableThreadingHTTPServer((BRIDGE_HOST, BRIDGE_PORT), WorkhubHermesBridgeHandler)
    print(f"Workhub Hermes bridge listening on {BRIDGE_HOST}:{BRIDGE_PORT}", flush=True)
    httpd.serve_forever()


if __name__ == "__main__":
    main()
