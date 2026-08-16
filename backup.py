#!/usr/bin/env python3
"""Daily backup of the Heli Beauty Studio data folder.

Creates a timestamped ZIP of the `data/` folder in `backups/`.
Run manually:   python backup.py
Schedule daily (Windows Task Scheduler / macOS launchd / cron):
    python backup.py
Keeps the most recent 30 backups and deletes older ones.
"""
import os
import sys
import zipfile
import shutil
from datetime import datetime

# Run from the directory containing this file (matters when frozen as EXE).
BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

DATA_DIR = os.path.join(BASE, "data")
BACKUP_DIR = os.path.join(BASE, "backups")
KEEP = 30  # number of recent backups to keep


def jalali_str():
    try:
        from web_app import PersianDate
        return PersianDate.today_str().replace("/", "-")
    except Exception:
        return datetime.now().strftime("%Y-%m-%d")


def main():
    if not os.path.isdir(DATA_DIR):
        print("data/ folder not found, nothing to back up.")
        return
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    jdate = jalali_str()
    zip_name = os.path.join(BACKUP_DIR, f"backup-{stamp}-({jdate}).zip")
    with zipfile.ZipFile(zip_name, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(DATA_DIR):
            for f in files:
                fp = os.path.join(root, f)
                z.write(fp, os.path.relpath(fp, BASE))
    print(f"Backup created: {zip_name}")

    # Rotate old backups
    existing = sorted(
        (os.path.join(BACKUP_DIR, f) for f in os.listdir(BACKUP_DIR) if f.endswith(".zip")),
        key=os.path.getmtime,
    )
    for old in existing[:-KEEP]:
        os.remove(old)
        print(f"Removed old backup: {old}")


if __name__ == "__main__":
    main()
