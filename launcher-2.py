"""
launcher.py – entry point for the packaged .exe
Starts the Streamlit server, opens the browser, and adds a system tray icon
with a Quit option to cleanly shut everything down.
"""
import sys
import os
import subprocess
import threading
import webbrowser
import time
import urllib.request


def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


def wait_for_streamlit(url, timeout=60):
    """Poll until Streamlit is actually serving, then open the browser."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            urllib.request.urlopen(url, timeout=1)
            webbrowser.open(url)
            return
        except Exception:
            time.sleep(0.5)
    # Fallback: open anyway after timeout
    webbrowser.open(url)


def main():
    if hasattr(sys, "_MEIPASS"):
        os.chdir(os.path.dirname(sys.executable))

    app_script = resource_path("automationtoolstreamlit19.py")
    port = "8501"
    url = f"http://localhost:{port}"

    # Build the streamlit command
    streamlit_cmd = [
        sys.executable, "-m", "streamlit", "run", app_script,
        "--server.port", port,
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
        "--global.developmentMode", "false",
    ]

    # Launch Streamlit as a subprocess (much more reliable than threading)
    creationflags = 0
    if sys.platform == "win32":
        creationflags = subprocess.CREATE_NO_WINDOW  # Hide console window on Windows

    proc = subprocess.Popen(
        streamlit_cmd,
        creationflags=creationflags,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    # Open browser only once Streamlit is actually ready
    threading.Thread(target=wait_for_streamlit, args=(url,), daemon=True).start()

    # Run system tray on MAIN thread (required on Windows)
    try:
        import pystray
        from PIL import Image, ImageDraw

        # Create a simple green circle icon
        img = Image.new("RGB", (64, 64), color=(255, 255, 255))
        draw = ImageDraw.Draw(img)
        draw.ellipse([8, 8, 56, 56], fill=(46, 139, 87))

        def on_open(icon, item):
            webbrowser.open(url)

        def on_quit(icon, item):
            icon.stop()
            proc.terminate()
            os._exit(0)

        menu = pystray.Menu(
            pystray.MenuItem("Open App", on_open),
            pystray.MenuItem("Quit", on_quit),
        )

        icon = pystray.Icon("GEP Tool", img, "GEP Package Tool", menu)
        icon.run()  # Blocks main thread — correct on Windows

    except Exception:
        # If tray fails, just wait for the subprocess to end
        proc.wait()


if __name__ == "__main__":
    main()
