from __future__ import annotations

import argparse
import html
import json
import os
import sys
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen


APP_TITLE = "(주)소일브릿지 업무자동화"
DEFAULT_APP_URL = "https://workhub.soilbridgecorp.cloud/"
APP_USER_AGENT = "SoilbridgeWorkhubDesktop/1.0"
STARTUP_SCRIPT_NAME = "SoilbridgeWorkhubDesktop_AutoStart.vbs"

LOCAL_APPDATA = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
DEFAULT_STORAGE_DIR = LOCAL_APPDATA / "SoilbridgeWorkhubDesktop" / "WebViewData"

_window = None


def resolve_app_url(cli_url: str | None = None) -> str:
    app_url = (cli_url or os.environ.get("WORKHUB_DESKTOP_URL") or DEFAULT_APP_URL).strip()
    parsed = urlparse(app_url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("WORKHUB_DESKTOP_URL은 http 또는 https 주소여야 합니다.")
    return app_url


def desktop_storage_dir() -> Path:
    configured = os.environ.get("WORKHUB_DESKTOP_STORAGE_DIR")
    if configured:
        return Path(configured).expanduser()
    return DEFAULT_STORAGE_DIR


def startup_folder() -> Path | None:
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return None
    return Path(appdata) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"


def register_startup_launch() -> None:
    if os.environ.get("WORKHUB_DESKTOP_DISABLE_AUTOSTART", "").strip() == "1":
        return
    folder = startup_folder()
    if not folder:
        return
    try:
        folder.mkdir(parents=True, exist_ok=True)
        startup_script = folder / STARTUP_SCRIPT_NAME
        executable = Path(sys.executable).resolve()
        if getattr(sys, "frozen", False):
            command = f'"{executable}"'
        else:
            command = f'"{executable}" "{Path(__file__).resolve()}"'
        escaped_command = command.replace('"', '""')
        script = (
            'Set shell = CreateObject("WScript.Shell")\n'
            f'shell.Run "{escaped_command}", 1, False\n'
        )
        if not startup_script.exists() or startup_script.read_text(encoding="utf-16", errors="ignore") != script:
            startup_script.write_text(script, encoding="utf-16")
    except Exception:
        return


def app_is_reachable(app_url: str, timeout: float = 6.0) -> tuple[bool, str]:
    request = Request(app_url, headers={"User-Agent": APP_USER_AGENT})
    try:
        with urlopen(request, timeout=timeout) as response:
            status = getattr(response, "status", 200)
            if status >= 500:
                return False, f"서버 응답 오류: HTTP {status}"
            return True, f"HTTP {status}"
    except HTTPError as exc:
        if exc.code < 500:
            return True, f"HTTP {exc.code}"
        return False, f"서버 응답 오류: HTTP {exc.code}"
    except URLError as exc:
        reason = getattr(exc, "reason", exc)
        return False, f"연결 실패: {reason}"
    except Exception as exc:  # noqa: BLE001
        return False, f"연결 확인 실패: {exc}"


def offline_html(app_url: str, message: str) -> str:
    escaped_url = html.escape(app_url)
    escaped_message = html.escape(message)
    return f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{APP_TITLE}</title>
  <style>
    :root {{
      color-scheme: light;
      font-family: "Malgun Gothic", "Segoe UI", Arial, sans-serif;
      background: #f5f7fb;
      color: #172033;
    }}
    body {{
      min-height: 100vh;
      margin: 0;
      display: grid;
      place-items: center;
    }}
    main {{
      width: min(520px, calc(100vw - 48px));
      padding: 32px;
      border: 1px solid #dbe3f0;
      border-radius: 14px;
      background: #ffffff;
      box-shadow: 0 20px 60px rgba(19, 38, 67, 0.12);
    }}
    .mark {{
      width: 52px;
      height: 52px;
      border-radius: 16px;
      display: grid;
      place-items: center;
      background: linear-gradient(135deg, #2457d6, #12a47f);
      color: #fff;
      font-weight: 800;
      letter-spacing: 0;
      margin-bottom: 18px;
    }}
    h1 {{
      margin: 0 0 10px;
      font-size: 22px;
      line-height: 1.35;
    }}
    p {{
      margin: 0;
      line-height: 1.7;
      color: #526070;
      word-break: keep-all;
    }}
    code {{
      display: block;
      margin-top: 16px;
      padding: 12px;
      border-radius: 10px;
      background: #f2f5fa;
      color: #344054;
      white-space: normal;
      word-break: break-all;
      font-family: "Cascadia Mono", Consolas, monospace;
      font-size: 13px;
    }}
    button {{
      margin-top: 22px;
      border: 0;
      border-radius: 10px;
      padding: 11px 16px;
      background: #2457d6;
      color: #fff;
      font-weight: 700;
      cursor: pointer;
    }}
    button:disabled {{
      background: #93a4c3;
      cursor: wait;
    }}
    #status {{
      margin-top: 14px;
      min-height: 24px;
      font-size: 14px;
      color: #6a7484;
    }}
  </style>
</head>
<body>
  <main>
    <div class="mark">SB</div>
    <h1>업무자동화 서버에 연결하지 못했어.</h1>
    <p>인터넷 연결이나 VPS 상태를 확인한 뒤 다시 시도해줘.</p>
    <code>{escaped_url}<br>{escaped_message}</code>
    <button id="retry" type="button">다시 연결</button>
    <p id="status" aria-live="polite"></p>
  </main>
  <script>
    const button = document.querySelector("#retry");
    const status = document.querySelector("#status");
    button.addEventListener("click", async () => {{
      button.disabled = true;
      status.textContent = "연결 확인 중...";
      try {{
        const result = await pywebview.api.retry();
        if (!result.ok) {{
          status.textContent = result.message || "아직 연결되지 않았어.";
          button.disabled = false;
        }}
      }} catch (error) {{
        status.textContent = "앱 내부 연결 확인에 실패했어.";
        button.disabled = false;
      }}
    }});
  </script>
</body>
</html>"""


class WorkhubDesktopApi:
    def __init__(self, app_url: str):
        self.app_url = app_url

    def retry(self) -> dict[str, object]:
        ok, message = app_is_reachable(self.app_url)
        if ok and _window is not None:
            _window.load_url(self.app_url)
        return {"ok": ok, "message": message}


def run_desktop_app(app_url: str, *, debug: bool = False, skip_preflight: bool = False) -> None:
    import webview

    storage_dir = desktop_storage_dir()
    storage_dir.mkdir(parents=True, exist_ok=True)

    api = WorkhubDesktopApi(app_url)
    ok, message = (True, "skipped") if skip_preflight else app_is_reachable(app_url)

    global _window
    window_kwargs = {
        "width": 1500,
        "height": 940,
        "min_size": (1180, 760),
        "confirm_close": False,
        "background_color": "#f5f7fb",
        "text_select": True,
        "zoomable": False,
        "js_api": api,
    }
    if ok:
        _window = webview.create_window(APP_TITLE, app_url, **window_kwargs)
    else:
        _window = webview.create_window(APP_TITLE, html=offline_html(app_url, message), **window_kwargs)

    webview.start(
        debug=debug,
        private_mode=False,
        storage_path=str(storage_dir),
        user_agent=APP_USER_AGENT,
    )


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Soilbridge Workhub desktop app")
    parser.add_argument("--url", help="Workhub URL. Defaults to WORKHUB_DESKTOP_URL or the production VPS URL.")
    parser.add_argument("--debug", action="store_true", help="Enable pywebview debug mode.")
    parser.add_argument("--skip-preflight", action="store_true", help="Open the app without checking URL reachability first.")
    parser.add_argument("--health-check", action="store_true", help="Check whether the configured Workhub URL is reachable and exit.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        app_url = resolve_app_url(args.url)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    if args.health_check:
        ok, message = app_is_reachable(app_url)
        print(json.dumps({"ok": ok, "url": app_url, "message": message}, ensure_ascii=False))
        return 0 if ok else 1

    register_startup_launch()
    run_desktop_app(app_url, debug=args.debug, skip_preflight=args.skip_preflight)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
