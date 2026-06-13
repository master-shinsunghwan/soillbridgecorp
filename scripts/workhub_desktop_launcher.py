from __future__ import annotations

import sys
import threading
import time
import webbrowser
from urllib.request import urlopen

from workhub_delivery_app import run


PORT = 8765
URL = f"http://127.0.0.1:{PORT}/"


def is_running() -> bool:
    try:
        with urlopen(URL, timeout=1) as response:
            return response.status == 200
    except Exception:
        return False


def main() -> None:
    if is_running():
        webbrowser.open(URL)
        return

    server_thread = threading.Thread(target=run, kwargs={"port": PORT}, daemon=True)
    server_thread.start()

    for _ in range(40):
        if is_running():
            webbrowser.open(URL)
            break
        time.sleep(0.25)

    try:
        while server_thread.is_alive():
            time.sleep(1)
    except KeyboardInterrupt:
        sys.exit(0)


if __name__ == "__main__":
    main()
