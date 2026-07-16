from __future__ import annotations

import argparse
import base64
import hashlib
import html
import json
import os
import re
import subprocess
import sys
import threading
import time
import uuid
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen


APP_TITLE = "(주)소일브릿지 업무자동화"
APP_VERSION = "1.1.0"
DEFAULT_APP_URL = "https://workhub.soilbridgecorp.cloud/"
DEFAULT_UPDATE_MANIFEST_URL = (
    "https://raw.githubusercontent.com/master-shinsunghwan/soillbridgecorp/"
    "main/static/desktop_update.json"
)
APP_USER_AGENT = f"SoilbridgeWorkhubDesktop/{APP_VERSION}"
STARTUP_SCRIPT_NAME = "SoilbridgeWorkhubDesktop_AutoStart.vbs"
MAX_UPDATE_BYTES = 100 * 1024 * 1024

LOCAL_APPDATA = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
DESKTOP_APP_DIR = LOCAL_APPDATA / "SoilbridgeWorkhubDesktop"
DEFAULT_STORAGE_DIR = DESKTOP_APP_DIR / "WebViewData"
DESKTOP_SETTINGS_PATH = DESKTOP_APP_DIR / "settings.json"
DESKTOP_UPDATE_DIR = DESKTOP_APP_DIR / "Updates"
PENDING_UPDATE_PATH = DESKTOP_UPDATE_DIR / "pending_update.json"
DEFAULT_DOWNLOAD_DIR = Path.home() / "Downloads"

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


def read_desktop_settings() -> dict[str, object]:
    try:
        data = json.loads(DESKTOP_SETTINGS_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def write_json_atomic(path: Path, payload: dict[str, object], *, encoding: str = "utf-8") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        temp_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding=encoding,
        )
        temp_path.replace(path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def write_desktop_settings(settings: dict[str, object]) -> None:
    write_json_atomic(DESKTOP_SETTINGS_PATH, settings)


def desktop_download_dir() -> Path:
    configured = os.environ.get("WORKHUB_DESKTOP_DOWNLOAD_DIR")
    if configured and configured.strip():
        return Path(configured).expanduser()
    saved = str(read_desktop_settings().get("download_dir") or "").strip()
    if saved:
        return Path(saved).expanduser()
    return DEFAULT_DOWNLOAD_DIR


def desktop_download_settings() -> dict[str, object]:
    settings = read_desktop_settings()
    saved = str(settings.get("download_dir") or "").strip()
    environment_override = bool(os.environ.get("WORKHUB_DESKTOP_DOWNLOAD_DIR", "").strip())
    return {
        "ok": True,
        "path": str(desktop_download_dir()),
        "default_path": str(DEFAULT_DOWNLOAD_DIR),
        "customized": bool(saved) and not environment_override,
        "source": "environment" if environment_override else ("custom" if saved else "default"),
    }


def set_desktop_download_dir(folder: str | Path) -> Path:
    raw_path = str(folder or "").strip()
    if not raw_path:
        raise ValueError("저장할 폴더를 선택해주세요.")
    path = Path(raw_path).expanduser()
    path.mkdir(parents=True, exist_ok=True)
    if not path.is_dir():
        raise NotADirectoryError("선택한 경로가 폴더가 아닙니다.")
    settings = read_desktop_settings()
    settings["download_dir"] = str(path)
    write_desktop_settings(settings)
    return path


def reset_desktop_download_dir() -> None:
    settings = read_desktop_settings()
    settings.pop("download_dir", None)
    write_desktop_settings(settings)


def version_parts(value: str) -> tuple[int, int, int, int]:
    match = re.fullmatch(r"v?(\d+(?:\.\d+){1,3})", str(value or "").strip(), re.IGNORECASE)
    if not match:
        raise ValueError("업데이트 버전 형식이 올바르지 않습니다.")
    parts = [int(part) for part in match.group(1).split(".")]
    padded = (parts + [0, 0, 0, 0])[:4]
    return padded[0], padded[1], padded[2], padded[3]


def is_newer_version(candidate: str, current: str = APP_VERSION) -> bool:
    return version_parts(candidate) > version_parts(current)


def secure_update_url(value: str, label: str) -> str:
    url = str(value or "").strip()
    parsed = urlparse(url)
    local_http = parsed.scheme == "http" and parsed.hostname in {"127.0.0.1", "localhost"}
    if (parsed.scheme != "https" and not local_http) or not parsed.netloc:
        raise ValueError(f"{label}은 HTTPS 주소여야 합니다.")
    return url


def desktop_update_manifest_url() -> str:
    configured = os.environ.get("WORKHUB_DESKTOP_UPDATE_MANIFEST_URL", "").strip()
    return secure_update_url(configured or DEFAULT_UPDATE_MANIFEST_URL, "업데이트 확인 주소")


def validate_desktop_update_manifest(payload: object) -> dict[str, object]:
    if not isinstance(payload, dict):
        raise ValueError("업데이트 정보 형식이 올바르지 않습니다.")
    version = str(payload.get("version") or "").strip()
    version_parts(version)
    download_url = secure_update_url(str(payload.get("url") or ""), "업데이트 파일 주소")
    sha256 = str(payload.get("sha256") or "").strip().lower()
    if not re.fullmatch(r"[0-9a-f]{64}", sha256):
        raise ValueError("업데이트 파일 검증값이 올바르지 않습니다.")
    try:
        size = int(payload.get("size") or 0)
    except (TypeError, ValueError) as exc:
        raise ValueError("업데이트 파일 크기가 올바르지 않습니다.") from exc
    if size <= 0 or size > MAX_UPDATE_BYTES:
        raise ValueError("업데이트 파일 크기가 허용 범위를 벗어났습니다.")
    return {
        "version": version,
        "url": download_url,
        "sha256": sha256,
        "size": size,
        "published_at": str(payload.get("published_at") or "").strip(),
        "notes": str(payload.get("notes") or "").strip(),
    }


def fetch_desktop_update_manifest(manifest_url: str | None = None, timeout: float = 6.0) -> dict[str, object]:
    url = secure_update_url(manifest_url or desktop_update_manifest_url(), "업데이트 확인 주소")
    request = Request(
        url,
        headers={
            "User-Agent": APP_USER_AGENT,
            "Accept": "application/json",
            "Cache-Control": "no-cache",
        },
    )
    with urlopen(request, timeout=timeout) as response:
        final_url = getattr(response, "geturl", lambda: url)()
        secure_update_url(final_url, "업데이트 확인 주소")
        data = response.read(64 * 1024 + 1)
    if len(data) > 64 * 1024:
        raise ValueError("업데이트 정보가 너무 큽니다.")
    try:
        payload = json.loads(data.decode("utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("업데이트 정보를 읽지 못했습니다.") from exc
    return validate_desktop_update_manifest(payload)


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def desktop_update_log(message: str) -> None:
    try:
        DESKTOP_APP_DIR.mkdir(parents=True, exist_ok=True)
        log_path = DESKTOP_APP_DIR / "update.log"
        if log_path.exists() and log_path.stat().st_size > 256 * 1024:
            log_path.write_text("", encoding="utf-8")
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(f"[{timestamp}] {message}\n")
    except OSError:
        return


def read_pending_desktop_update() -> dict[str, object] | None:
    try:
        payload = json.loads(PENDING_UPDATE_PATH.read_text(encoding="utf-8"))
        manifest = validate_desktop_update_manifest(payload)
        if not is_newer_version(str(manifest["version"])):
            return None
        staged_path = Path(str(payload.get("path") or "")).resolve()
        if staged_path.parent != DESKTOP_UPDATE_DIR.resolve() or not staged_path.is_file():
            return None
        if staged_path.stat().st_size != int(manifest["size"]):
            return None
        if file_sha256(staged_path) != str(manifest["sha256"]):
            return None
        return {**manifest, "path": str(staged_path)}
    except (FileNotFoundError, json.JSONDecodeError, OSError, ValueError):
        return None


def download_desktop_update(manifest: dict[str, object], timeout: float = 45.0) -> Path:
    DESKTOP_UPDATE_DIR.mkdir(parents=True, exist_ok=True)
    version = str(manifest["version"])
    expected_size = int(manifest["size"])
    expected_hash = str(manifest["sha256"])
    target = DESKTOP_UPDATE_DIR / f"SoilbridgeWorkhub-{version}.exe"
    if target.is_file() and target.stat().st_size == expected_size and file_sha256(target) == expected_hash:
        return target
    if target.exists():
        target.unlink()
    temp_path = target.with_name(f".{target.name}.{uuid.uuid4().hex}.part")
    request = Request(
        str(manifest["url"]),
        headers={"User-Agent": APP_USER_AGENT, "Accept": "application/octet-stream"},
    )
    total = 0
    digest = hashlib.sha256()
    try:
        with urlopen(request, timeout=timeout) as response, temp_path.open("wb") as handle:
            final_url = getattr(response, "geturl", lambda: str(manifest["url"]))()
            secure_update_url(final_url, "업데이트 파일 주소")
            for chunk in iter(lambda: response.read(1024 * 1024), b""):
                total += len(chunk)
                if total > MAX_UPDATE_BYTES or total > expected_size:
                    raise ValueError("업데이트 파일 크기가 예상보다 큽니다.")
                digest.update(chunk)
                handle.write(chunk)
        if total != expected_size:
            raise ValueError("업데이트 파일 크기가 일치하지 않습니다.")
        if digest.hexdigest() != expected_hash:
            raise ValueError("업데이트 파일 검증값이 일치하지 않습니다.")
        temp_path.replace(target)
        return target
    finally:
        if temp_path.exists():
            temp_path.unlink()


def stage_available_desktop_update(manifest_url: str | None = None) -> dict[str, object]:
    manifest = fetch_desktop_update_manifest(manifest_url)
    if not is_newer_version(str(manifest["version"])):
        return {"ok": True, "update_available": False, "version": APP_VERSION}
    pending = read_pending_desktop_update()
    if pending and pending["version"] == manifest["version"] and pending["sha256"] == manifest["sha256"]:
        return {"ok": True, "update_available": True, "staged": True, **pending}
    staged_path = download_desktop_update(manifest)
    pending_payload = {**manifest, "path": str(staged_path)}
    write_json_atomic(PENDING_UPDATE_PATH, pending_payload)
    desktop_update_log(f"업데이트 {manifest['version']} 다운로드 및 검증 완료")
    return {"ok": True, "update_available": True, "staged": True, **pending_payload}


def desktop_updates_enabled() -> bool:
    disabled = os.environ.get("WORKHUB_DESKTOP_DISABLE_UPDATES", "").strip() == "1"
    return os.name == "nt" and bool(getattr(sys, "frozen", False)) and not disabled


def check_desktop_update_in_background() -> None:
    if not desktop_updates_enabled():
        return
    try:
        stage_available_desktop_update()
    except Exception as exc:  # noqa: BLE001
        desktop_update_log(f"업데이트 확인 실패: {exc}")


def start_desktop_update_check() -> None:
    if not desktop_updates_enabled():
        return
    threading.Thread(
        target=check_desktop_update_in_background,
        name="workhub-desktop-updater",
        daemon=True,
    ).start()


def powershell_literal(value: str | Path) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def pending_update_script(pending: dict[str, object], target_executable: Path, parent_pid: int) -> str:
    source = powershell_literal(str(pending["path"]))
    target = powershell_literal(target_executable)
    pending_path = powershell_literal(PENDING_UPDATE_PATH)
    expected_hash = powershell_literal(str(pending["sha256"]).upper())
    return f'''$ErrorActionPreference = "Stop"
$Source = {source}
$Target = {target}
$Pending = {pending_path}
$ExpectedHash = {expected_hash}
$ParentPid = {int(parent_pid)}
$Backup = "$Target.update-backup"

try {{
  Wait-Process -Id $ParentPid -ErrorAction SilentlyContinue
  if ((Get-FileHash -Algorithm SHA256 -LiteralPath $Source).Hash.ToUpperInvariant() -ne $ExpectedHash) {{
    throw "Update source hash mismatch"
  }}
  if (Test-Path -LiteralPath $Backup) {{ Remove-Item -LiteralPath $Backup -Force }}
  if (Test-Path -LiteralPath $Target) {{ Copy-Item -LiteralPath $Target -Destination $Backup -Force }}
  $Copied = $false
  for ($Attempt = 0; $Attempt -lt 40; $Attempt += 1) {{
    try {{
      Copy-Item -LiteralPath $Source -Destination $Target -Force
      $Copied = $true
      break
    }} catch {{
      Start-Sleep -Milliseconds 250
    }}
  }}
  if (-not $Copied) {{ throw "Update target remained locked" }}
  if ((Get-FileHash -Algorithm SHA256 -LiteralPath $Target).Hash.ToUpperInvariant() -ne $ExpectedHash) {{
    throw "Updated executable hash mismatch"
  }}
  Remove-Item -LiteralPath $Pending -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $Source -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $Backup -Force -ErrorAction SilentlyContinue
  Start-Process -FilePath $Target -WorkingDirectory (Split-Path -Parent $Target)
}} catch {{
  if (Test-Path -LiteralPath $Backup) {{
    Copy-Item -LiteralPath $Backup -Destination $Target -Force -ErrorAction SilentlyContinue
  }}
  if (Test-Path -LiteralPath $Target) {{
    Start-Process -FilePath $Target -WorkingDirectory (Split-Path -Parent $Target)
  }}
}} finally {{
  Start-Sleep -Milliseconds 300
  Remove-Item -LiteralPath $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue
}}
'''


def launch_pending_desktop_update() -> bool:
    if not desktop_updates_enabled():
        return False
    pending = read_pending_desktop_update()
    if not pending:
        return False
    try:
        target_executable = Path(sys.executable).resolve()
        script_path = DESKTOP_UPDATE_DIR / f"apply_update_{uuid.uuid4().hex}.ps1"
        script_path.write_text(
            pending_update_script(pending, target_executable, os.getpid()),
            encoding="utf-8-sig",
        )
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | getattr(subprocess, "DETACHED_PROCESS", 0)
        subprocess.Popen(
            [
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
            ],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
            creationflags=creation_flags,
        )
        desktop_update_log(f"업데이트 {pending['version']} 적용 시작")
        return True
    except Exception as exc:  # noqa: BLE001
        desktop_update_log(f"업데이트 적용 시작 실패: {exc}")
        return False


def safe_download_filename(filename: str | None, fallback: str = "workhub_download.xlsx") -> str:
    value = str(filename or "").strip() or fallback
    value = value.replace("\\", "_").replace("/", "_").replace(":", "_")
    value = re.sub(r"[\x00-\x1f<>|?*]+", "_", value)
    value = value.strip(" .")
    return value or fallback


def unique_download_path(folder: Path, filename: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    candidate = folder / filename
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    for index in range(1, 1000):
        numbered = folder / f"{stem} ({index}){suffix}"
        if not numbered.exists():
            return numbered
    raise FileExistsError("Cannot find an available download filename.")


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
        self._active_downloads: dict[str, dict[str, object]] = {}

    def retry(self) -> dict[str, object]:
        ok, message = app_is_reachable(self.app_url)
        if ok and _window is not None:
            _window.load_url(self.app_url)
        return {"ok": ok, "message": message}

    def getDownloadSettings(self) -> dict[str, object]:
        return desktop_download_settings()

    def chooseDownloadFolder(self) -> dict[str, object]:
        if _window is None:
            return {"ok": False, "message": "앱 창을 찾지 못했습니다."}
        try:
            import webview

            selected = _window.create_file_dialog(
                webview.FOLDER_DIALOG,
                directory=str(desktop_download_dir()),
            )
            if not selected:
                return {**desktop_download_settings(), "cancelled": True}
            selected_path = selected[0] if isinstance(selected, (list, tuple)) else selected
            set_desktop_download_dir(str(selected_path))
            return desktop_download_settings()
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "message": str(exc)}

    def resetDownloadFolder(self) -> dict[str, object]:
        try:
            reset_desktop_download_dir()
            return desktop_download_settings()
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "message": str(exc)}

    def get_download_settings(self) -> dict[str, object]:
        return self.getDownloadSettings()

    def choose_download_folder(self) -> dict[str, object]:
        return self.chooseDownloadFolder()

    def reset_download_folder(self) -> dict[str, object]:
        return self.resetDownloadFolder()

    def saveDownload(self, filename: str, base64_data: str) -> dict[str, object]:
        try:
            data = base64.b64decode(str(base64_data or ""), validate=True)
            download_dir = desktop_download_dir()
            path = unique_download_path(download_dir, safe_download_filename(filename))
            path.write_bytes(data)
            return {"ok": True, "path": str(path), "filename": path.name, "size": len(data)}
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "message": str(exc)}

    def save_download(self, filename: str, base64_data: str) -> dict[str, object]:
        return self.saveDownload(filename, base64_data)

    def beginDownload(self, filename: str) -> dict[str, object]:
        try:
            download_dir = desktop_download_dir()
            final_path = unique_download_path(download_dir, safe_download_filename(filename))
            download_id = uuid.uuid4().hex
            temp_path = final_path.with_name(f".{final_path.name}.{download_id}.part")
            handle = temp_path.open("wb")
            self._active_downloads[download_id] = {
                "handle": handle,
                "temp_path": temp_path,
                "final_path": final_path,
                "size": 0,
            }
            return {"ok": True, "id": download_id, "filename": final_path.name}
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "message": str(exc)}

    def appendDownloadChunk(self, download_id: str, base64_data: str) -> dict[str, object]:
        try:
            download = self._active_downloads.get(str(download_id or ""))
            if not download:
                return {"ok": False, "message": "download session not found"}
            data = base64.b64decode(str(base64_data or ""), validate=True)
            handle = download["handle"]
            handle.write(data)
            download["size"] = int(download.get("size") or 0) + len(data)
            return {"ok": True, "size": download["size"]}
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "message": str(exc)}

    def finishDownload(self, download_id: str) -> dict[str, object]:
        download = self._active_downloads.pop(str(download_id or ""), None)
        if not download:
            return {"ok": False, "message": "download session not found"}
        try:
            handle = download["handle"]
            handle.close()
            temp_path = download["temp_path"]
            final_path = download["final_path"]
            temp_path.replace(final_path)
            return {
                "ok": True,
                "path": str(final_path),
                "filename": final_path.name,
                "size": int(download.get("size") or 0),
            }
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "message": str(exc)}

    def cancelDownload(self, download_id: str) -> dict[str, object]:
        download = self._active_downloads.pop(str(download_id or ""), None)
        if not download:
            return {"ok": True}
        try:
            handle = download["handle"]
            handle.close()
            temp_path = download["temp_path"]
            if temp_path.exists():
                temp_path.unlink()
            return {"ok": True}
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "message": str(exc)}

    def begin_download(self, filename: str) -> dict[str, object]:
        return self.beginDownload(filename)

    def append_download_chunk(self, download_id: str, base64_data: str) -> dict[str, object]:
        return self.appendDownloadChunk(download_id, base64_data)

    def finish_download(self, download_id: str) -> dict[str, object]:
        return self.finishDownload(download_id)

    def cancel_download(self, download_id: str) -> dict[str, object]:
        return self.cancelDownload(download_id)


def run_desktop_app(app_url: str, *, debug: bool = False, skip_preflight: bool = False) -> None:
    import webview

    # pywebview blocks WebView2 downloads by default. Workhub creates Excel and
    # shared-file downloads from Blob URLs, so the desktop shell must opt in.
    webview.settings["ALLOW_DOWNLOADS"] = True

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

    start_desktop_update_check()
    webview.start(
        debug=debug,
        private_mode=False,
        storage_path=str(storage_dir),
        user_agent=APP_USER_AGENT,
    )


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Soilbridge Workhub desktop app")
    parser.add_argument("--version", action="version", version=APP_VERSION)
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

    if not args.health_check and launch_pending_desktop_update():
        return 0

    if args.health_check:
        ok, message = app_is_reachable(app_url)
        print(json.dumps({"ok": ok, "url": app_url, "message": message}, ensure_ascii=False))
        return 0 if ok else 1

    register_startup_launch()
    run_desktop_app(app_url, debug=args.debug, skip_preflight=args.skip_preflight)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
