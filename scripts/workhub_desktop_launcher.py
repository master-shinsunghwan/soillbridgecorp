from __future__ import annotations

import sys
import threading
import time
import webbrowser
from urllib.request import urlopen

from workhub_delivery_app import run


PORT = 8765
URL = f"http://127.0.0.1:{PORT}/"
APP_TITLE = "(주)소일브릿지 업무자동화"


def is_running() -> bool:
    try:
        with urlopen(URL, timeout=1) as response:
            return response.status == 200
    except Exception:
        return False


def main() -> None:
    server_thread: threading.Thread | None = None

    if not is_running():
        server_thread = threading.Thread(target=run, kwargs={"port": PORT}, daemon=True)
        server_thread.start()

        for _ in range(40):
            if is_running():
                break
            time.sleep(0.25)
        else:
            raise RuntimeError("Workhub 서버를 시작하지 못했습니다.")

    try:
        import webview

        webview.create_window(
            APP_TITLE,
            URL,
            width=1480,
            height=920,
            min_size=(1180, 760),
            confirm_close=True,
        )
        webview.start()
    except Exception:
        webbrowser.open(URL)
        try:
            while server_thread and server_thread.is_alive():
                time.sleep(1)
        except KeyboardInterrupt:
            sys.exit(0)
        return

    if server_thread:
        try:
            while server_thread.is_alive():
                time.sleep(1)
        except KeyboardInterrupt:
            sys.exit(0)
    else:
        sys.exit(0)


if __name__ == "__main__":
    main()
