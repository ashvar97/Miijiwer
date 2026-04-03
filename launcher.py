"""
launcher.py – entry point for the packaged .exe
Single-instance guarded. Starts Streamlit, opens browser, adds system tray.
"""
import sys
import os

# ── CRITICAL: must be the very first thing before any other imports ──────────
# Prevents PyInstaller child processes from re-running main() on Windows.
import multiprocessing
multiprocessing.freeze_support()
# ─────────────────────────────────────────────────────────────────────────────

import socket
import subprocess
import threading
import webbrowser
import time
import urllib.request


# ── Single-instance lock via a bound socket ──────────────────────────────────
_LOCK_PORT = 47200  # arbitrary port just for the lock

def acquire_instance_lock():
    """Returns a bound socket (lock) or None if another instance is running."""
    lock_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    lock_sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 0)
    try:
        lock_sock.bind(("127.0.0.1", _LOCK_PORT))
        return lock_sock  # We are the first instance
    except OSError:
        return None  # Another instance already holds the lock
# ─────────────────────────────────────────────────────────────────────────────


def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


def get_log_path():
    """Returns path for the log file next to the exe (or script during dev)."""
    if hasattr(sys, "_MEIPASS"):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "streamlit_error.log")


def get_streamlit_cmd(app_script, port):
    if hasattr(sys, "_MEIPASS"):
        cmd = [sys.executable, "-m", "streamlit"]
    else:
        scripts_dir = os.path.dirname(sys.executable)
        streamlit_exe = os.path.join(scripts_dir, "streamlit.exe")
        cmd = [streamlit_exe] if os.path.exists(streamlit_exe) else [sys.executable, "-m", "streamlit"]

    return cmd + [
        "run", app_script,
        "--server.port", port,
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
        "--global.developmentMode", "false",
    ]


def wait_for_streamlit(url, timeout=90):
    """Poll until Streamlit responds, then open the browser once."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            urllib.request.urlopen(url, timeout=2)
            webbrowser.open(url)
            return
        except Exception:
            time.sleep(1)
    webbrowser.open(url)


def run_streamlit_threaded(app_script, port):
    """Fallback: run streamlit in-process if subprocess fails."""
    os.environ["STREAMLIT_SERVER_PORT"] = port
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"

    def _run():
        from streamlit.web import cli as stcli
        sys.argv = ["streamlit", "run", app_script]
        stcli.main()

    threading.Thread(target=_run, daemon=True).start()


def main():
    # ── Guard: only one instance allowed ────────────────────────────────────
    lock = acquire_instance_lock()
    if lock is None:
        # Another instance is already running — just focus its browser tab
        webbrowser.open("http://localhost:8501")
        sys.exit(0)
    # ────────────────────────────────────────────────────────────────────────

    if hasattr(sys, "_MEIPASS"):
        os.chdir(os.path.dirname(sys.executable))

    app_script = resource_path("automationtoolstreamlit19.py")
    port = "8501"
    url = f"http://localhost:{port}"

    # ── Verify the app script exists before trying to launch ─────────────────
    log_path = get_log_path()
    if not os.path.exists(app_script):
        with open(log_path, "w") as f:
            f.write(f"ERROR: app script not found at: {app_script}\n")
            f.write(f"sys._MEIPASS = {getattr(sys, '_MEIPASS', 'N/A')}\n")
            f.write(f"sys.executable = {sys.executable}\n")
        sys.exit(1)
    # ─────────────────────────────────────────────────────────────────────────

    env = os.environ.copy()
    if hasattr(sys, "_MEIPASS"):
        env["PYTHONPATH"] = sys._MEIPASS

    creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0

    proc = None
    try:
        # ── Log stdout+stderr to file so failures are visible ────────────────
        log_file = open(log_path, "w", buffering=1)
        log_file.write(f"Launching Streamlit...\n")
        log_file.write(f"app_script: {app_script}\n")
        log_file.write(f"cmd: {get_streamlit_cmd(app_script, port)}\n\n")
        log_file.flush()
        # ─────────────────────────────────────────────────────────────────────

        proc = subprocess.Popen(
            get_streamlit_cmd(app_script, port),
            creationflags=creationflags,
            stdout=log_file,
            stderr=log_file,
            env=env,
        )
    except Exception as e:
        with open(log_path, "a") as f:
            f.write(f"\nsubprocess.Popen failed: {e}\n")
            f.write("Falling back to threaded mode...\n")
        run_streamlit_threaded(app_script, port)

    threading.Thread(target=wait_for_streamlit, args=(url,), daemon=True).start()

    # ── System tray (main thread — required on Windows) ──────────────────────
    try:
        import pystray
        from PIL import Image, ImageDraw

        img = Image.new("RGB", (64, 64), color=(255, 255, 255))
        draw = ImageDraw.Draw(img)
        draw.ellipse([8, 8, 56, 56], fill=(46, 139, 87))

        def on_open(icon, item):
            webbrowser.open(url)

        def on_quit(icon, item):
            icon.stop()
            if proc:
                proc.terminate()
            lock.close()
            os._exit(0)

        menu = pystray.Menu(
            pystray.MenuItem("Open App", on_open),
            pystray.MenuItem("Quit", on_quit),
        )

        icon = pystray.Icon("GEP Tool", img, "GEP Package Tool", menu)
        icon.run()

    except Exception:
        if proc:
            proc.wait()
        else:
            while True:
                time.sleep(1)


if __name__ == "__main__":
    main()
