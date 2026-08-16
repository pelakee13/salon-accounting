#!/usr/bin/env python3
"""Launcher for Heli Beauty Studio — starts the Flask server and opens the browser.

Run with:  python run.py
When bundled as a PyInstaller EXE, double-clicking the .exe does the same.
Data is stored in a `data/` folder next to this file (or next to the .exe),
so it is easy to back up or migrate to a server later.
"""
import os
import sys
import webbrowser
import threading
import time

# Make sure we run from the directory containing this file (matters when frozen).
BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

from web_app import app, init_all

HOST = "127.0.0.1"
PORT = 5000


def open_browser():
    time.sleep(2.5)
    try:
        webbrowser.open(f"http://{HOST}:{PORT}/")
    except Exception:
        pass


def main():
    init_all()
    # Open browser automatically after the server starts.
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
