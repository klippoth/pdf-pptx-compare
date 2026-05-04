from __future__ import annotations

import socket
import threading
import time
import webbrowser

import uvicorn

from app.main import app


def _find_available_port(preferred_port: int = 8000) -> int:
    for candidate in (preferred_port, 8001, 8002, 8003, 0):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(("127.0.0.1", candidate))
            except OSError:
                continue
            port = sock.getsockname()[1]
        if candidate == 0 or port == candidate:
            return port
    return preferred_port


def _open_browser_when_ready(url: str, host: str, port: int) -> None:
    for _ in range(50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            if sock.connect_ex((host, port)) == 0:
                webbrowser.open(url)
                return
        time.sleep(0.2)
    webbrowser.open(url)


def main() -> None:
    port = _find_available_port()
    url = f"http://127.0.0.1:{port}"
    opener = threading.Thread(target=_open_browser_when_ready, args=(url, "127.0.0.1", port))
    opener.daemon = True
    opener.start()

    config = uvicorn.Config(
        app=app,
        host="127.0.0.1",
        port=port,
        reload=False,
        log_level="info",
    )
    uvicorn.Server(config).run()


if __name__ == "__main__":
    main()
