from __future__ import annotations

import ctypes
from datetime import datetime
import multiprocessing
import os
import socket
import sys
import threading
import time
import traceback
import webbrowser
from pathlib import Path

import uvicorn

_STREAM_HANDLES = []


def _launcher_log_path() -> Path:
    if sys.platform == "win32":
        local_app_data = os.environ.get("LOCALAPPDATA")
        if local_app_data:
            log_dir = Path(local_app_data) / "PDFtoPPTXReference"
            log_dir.mkdir(parents=True, exist_ok=True)
            return log_dir / "launcher.log"
    return Path.cwd() / "launcher.log"


def _append_launcher_log(message: str) -> Path:
    log_path = _launcher_log_path()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with log_path.open("a", encoding="utf-8") as handle:
        handle.write(f"[{timestamp}] {message}\n")
    return log_path


def _show_windows_error(message: str) -> None:
    if sys.platform != "win32":
        return
    ctypes.windll.user32.MessageBoxW(0, message, "PDFtoPPTXReference Launch Error", 0x10)


def _ensure_standard_streams() -> None:
    # PyInstaller windowed apps on Windows can start with stdout/stderr set to None.
    # Uvicorn's default logging checks .isatty(), so give it harmless streams.
    for stream_name in ("stdin", "stdout", "stderr"):
        if getattr(sys, stream_name) is not None:
            continue
        mode = "r" if stream_name == "stdin" else "w"
        handle = open(os.devnull, mode, encoding="utf-8")
        _STREAM_HANDLES.append(handle)
        setattr(sys, stream_name, handle)


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
    multiprocessing.freeze_support()
    _ensure_standard_streams()
    port = _find_available_port()
    url = f"http://127.0.0.1:{port}"
    _append_launcher_log(f"Launching local server at {url}")
    opener = threading.Thread(target=_open_browser_when_ready, args=(url, "127.0.0.1", port))
    opener.daemon = True
    opener.start()

    from app.main import app

    config = uvicorn.Config(
        app=app,
        host="127.0.0.1",
        port=port,
        reload=False,
        access_log=False,
        log_config=None,
        log_level="warning",
    )
    uvicorn.Server(config).run()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        error_text = traceback.format_exc()
        log_path = _append_launcher_log(error_text.rstrip())
        _show_windows_error(
            "The app could not start.\n\n"
            f"Details were written to:\n{log_path}\n\n"
            "Please send that file to support."
        )
        if sys.platform == "win32":
            sys.exit(1)
        raise
