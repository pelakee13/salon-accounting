# Build script for Windows EXE (run on Windows with Python 3.11+ installed).
# Usage:  pip install -r requirements.txt pyinstaller
#         python build.py
# Output: dist/HeliBeautyStudio.exe  (single file, portable)
import os
import sys
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))


def main():
    # Data folder must NOT be bundled (it is read/write at runtime next to exe).
    # Only bundle the Python code, templates and static assets.
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", "HeliBeautyStudio",
        "--onefile",
        "--windowed",          # no console window on Windows (GUI-like)
        "--add-data", os.path.join(HERE, "templates") + os.pathsep + "templates",
        "--add-data", os.path.join(HERE, "static") + os.pathsep + "static",
        "--hidden-import", "web_app",
        "--hidden-import", "auth",
        "--hidden-import", "openpyxl",
        "--hidden-import", "reportlab",
        "--hidden-import", "reportlab.pdfgen",
        "--hidden-import", "reportlab.pdfbase",
        "--hidden-import", "arabic_reshaper",
        "--hidden-import", "bidi",
        "--hidden-import", "bidi.algorithm",
        "--hidden-import", "werkzeug.security",
        "--hidden-import", "jinja2",
        "--hidden-import", "flask",
        os.path.join(HERE, "run.py"),
    ]
    print("Running:", " ".join(cmd))
    subprocess.check_call(cmd)


if __name__ == "__main__":
    main()
