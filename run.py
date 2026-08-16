#!/usr/bin/env python3
"""Launcher for Heli Beauty Studio — starts the Flask server and opens the browser.

Run with:  python run.py
When bundled as a PyInstaller EXE, double-clicking the .exe does the same.
Data is stored in a `data/` folder next to this file (or next to the .exe),
so it is easy to back up or migrate to a server later.

If anything fails, the full error is written to `app_error.log` next to the
executable so it can be reported.
"""
import os
import sys
import traceback

# Make sure we run from the directory containing this file (matters when frozen).
BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

LOG_PATH = os.path.join(BASE, "app_error.log")


def log_err(msg):
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write("\n=== %s ===\n%s\n" % (msg, traceback.format_exc()))
    except Exception:
        pass


HOST = "127.0.0.1"
PORT = 5000


def open_browser():
    import time
    import webbrowser
    time.sleep(2.5)
    try:
        webbrowser.open(f"http://{HOST}:{PORT}/")
    except Exception:
        pass


def main():
    try:
        from web_app import app, init_all
        init_all()
    except Exception:
        log_err("STARTUP ERROR (before server start)")
        print("Startup failed — see app_error.log next to this file.")
        input("Press Enter to exit...")
        sys.exit(1)

    # Open browser automatically after the server starts.
    import threading
    threading.Thread(target=open_browser, daemon=True).start()
    print("=" * 50)
    print("  Heli Beauty Studio — سالن زیبایی هلیا")
    print("=" * 50)
    print(f"  باز کردن در مرورگر:  http://{HOST}:{PORT}/")
    print("  برای بستن برنامه، این پنجره را ببندید (Ctrl+C)")
    print("=" * 50)
    try:
        app.run(host=HOST, port=PORT, debug=False, use_reloader=False)
    except KeyboardInterrupt:
        print("\nبرنامه بسته شد.")


if __name__ == "__main__":
    main()
