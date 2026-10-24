"""Start the local web app and open the default browser after the port is ready."""
from __future__ import annotations

import os
import threading
import time
import urllib.request
import webbrowser

from app import app


HOST = os.getenv("PAPER_CHECKER_HOST", "127.0.0.1")
PORT = int(os.getenv("PAPER_CHECKER_PORT", "5000"))
URL = f"http://127.0.0.1:{PORT}/"


def open_browser_when_ready() -> None:
    deadline = time.time() + 20
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(URL, timeout=1) as response:
                if response.status < 500:
                    webbrowser.open(URL, new=2)
                    return
        except Exception:
            time.sleep(0.25)
    # Still open the page so the user gets a useful browser tab if the app
    # responds slowly on an older classroom computer.
    webbrowser.open(URL, new=2)


if __name__ == "__main__":
    threading.Thread(target=open_browser_when_ready, daemon=True).start()
    print(f"Starting Paper Format Checker at {URL}", flush=True)
    app.run(debug=False, host=HOST, port=PORT)

