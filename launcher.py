"""
launcher.py – entry point for the packaged .exe
Starts the Streamlit server, opens the browser, and adds a system tray icon
with a Quit option to cleanly shut everything down.
"""
import sys
import os
import threading
import webbrowser
import time


def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


def open_browser():
    time.sleep(4)
    webbrowser.open("http://localhost:8501")


def start_streamlit(app_script):
    """Run Streamlit in a background thread."""
    from streamlit.web import cli as stcli
    sys.argv = ["streamlit", "run", app_script]
    stcli.main()


def main():
    if hasattr(sys, "_MEIPASS"):
        os.chdir(os.path.dirname(sys.executable))

    app_script = resource_path("automationtoolstreamlit19.py")

    os.environ["STREAMLIT_SERVER_PORT"] = "8501"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"

    # Start Streamlit in background thread
    t = threading.Thread(target=start_streamlit, args=(app_script,), daemon=True)
    t.start()

    # Open browser in background thread
    threading.Thread(target=open_browser, daemon=True).start()

    # Run system tray on MAIN thread (required on Windows)
    try:
        import pystray
        from PIL import Image, ImageDraw

        # Create a simple green circle icon
        img = Image.new("RGB", (64, 64), color=(255, 255, 255))
        draw = ImageDraw.Draw(img)
        draw.ellipse([8, 8, 56, 56], fill=(46, 139, 87))

        def on_open(icon, item):
            webbrowser.open("http://localhost:8501")

        def on_quit(icon, item):
            icon.stop()
            os._exit(0)

        menu = pystray.Menu(
            pystray.MenuItem("Open App", on_open),
            pystray.MenuItem("Quit", on_quit),
        )

        icon = pystray.Icon("GEP Tool", img, "GEP Package Tool", menu)
        icon.run()  # Blocks main thread — this is correct on Windows

    except Exception as e:
        # If tray fails, just wait for streamlit thread
        t.join()


if __name__ == "__main__":
    main()
