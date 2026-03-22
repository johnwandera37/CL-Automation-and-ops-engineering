"""
ETIMS Backup Uploader
Uploads latest ETIMS backup files from petrol station PCs to Google Drive.
"""

import os
import sys
import json
import time
import socket
import logging
import argparse
from pathlib import Path
from datetime import datetime

from drive_service import DriveService
from backup_finder import find_latest_backup
from error_logger import log_critical_error

# ── Base directory: works both as .py script and as PyInstaller exe ──────────
# When frozen (exe), __file__ points to the temp extraction folder.
# sys.executable always points to the actual exe location.
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = BASE_DIR

# ── Local logging with rotation (max 2 MB, keep 3 backup files) ─────────────
from logging.handlers import RotatingFileHandler
log_path = BASE_DIR / "etims_backup.log"
_file_handler = RotatingFileHandler(
    log_path, maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8"
)
_file_handler.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s"))
_console_handler = logging.StreamHandler(sys.stdout)
_console_handler.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s"))
logging.basicConfig(level=logging.INFO, handlers=[_file_handler, _console_handler])
log = logging.getLogger(__name__)


# ── Helpers ──────────────────────────────────────────────────────────────────

def load_config() -> dict:
    config_path = BASE_DIR / "config.json"
    if not config_path.exists():
        log.error("config.json not found. Run with --setup first.")
        sys.exit(1)
    with open(config_path, "r", encoding="utf-8-sig") as fh:  # utf-8-sig handles BOM
        return json.load(fh)


def save_config(config: dict):
    config_path = BASE_DIR / "config.json"
    with open(config_path, "w", encoding="utf-8") as fh:
        json.dump(config, fh, indent=2)


def is_internet_available() -> bool:
    try:
        socket.setdefaulttimeout(5)
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.connect(("8.8.8.8", 53))
        return True
    except OSError:
        return False


def wait_for_internet(max_wait_minutes: int = 120) -> bool:
    """Poll every 30 s until internet is up or timeout."""
    log.info("No internet. Waiting (up to %d min)…", max_wait_minutes)
    deadline = time.time() + max_wait_minutes * 60
    while time.time() < deadline:
        time.sleep(30)
        if is_internet_available():
            log.info("Internet connection restored.")
            return True
        log.info("Still offline… retrying in 30 s")
    return False


# ── Core upload logic ─────────────────────────────────────────────────────────

def run_backup(config: dict, drive: DriveService) -> bool:
    backup_path   = config.get("backup_path", "")
    station_label = f"{config['station_type']} {config['station_name']}"

    # 1. Find the latest backup (auto-detects if backup_path not set)
    latest, backup_dir = find_latest_backup(backup_path)
    if not latest:
        location = str(backup_dir) if backup_dir else "any known location"
        msg = f"[{station_label}] No ETIMS backup file found in {location}"
        log.error(msg)
        log_critical_error(drive, config, msg)
        return False
    log.info("Using backup directory: %s", backup_dir)

    log.info("Latest backup: %s", latest.name)

    try:
        root_id      = config["drive_root_folder_id"]
        type_folder  = drive.get_or_create_folder(config["station_type"], root_id)
        sub_folder   = drive.get_or_create_folder(config["station_name"], type_folder)

        # 2. Skip if already uploaded
        if drive.file_exists(latest.name, sub_folder):
            log.info("File already on Drive – nothing to do.")
            return True

        # 3. Upload
        log.info("Uploading %s …", latest.name)
        drive.upload_file(str(latest), sub_folder)
        log.info("Upload complete.")
        return True

    except Exception as exc:
        msg = f"[{station_label}] Upload failed: {exc}"
        log.error(msg, exc_info=True)
        log_critical_error(drive, config, msg)
        return False


# ── Cleanup old backups ───────────────────────────────────────────────────────

def cleanup_old_files(config: dict, drive: DriveService, keep: int = 1):
    """Delete old ETIMS backups from Drive, keeping only the `keep` newest."""
    root_id     = config["drive_root_folder_id"]
    type_folder = drive.get_or_create_folder(config["station_type"], root_id)
    sub_folder  = drive.get_or_create_folder(config["station_name"], type_folder)

    files = drive.list_backup_files(sub_folder)
    # Sort by name descending (name contains date, e.g. ETIMS202602240813)
    files_sorted = sorted(files, key=lambda f: f["name"], reverse=True)

    to_delete = files_sorted[keep:]
    if not to_delete:
        log.info("Nothing to delete – %d file(s) present, keeping %d.", len(files_sorted), keep)
        return

    log.info("Deleting %d old backup(s), keeping latest %d…", len(to_delete), keep)
    for f in to_delete:
        drive.delete_file(f["id"])
        log.info("  Deleted: %s", f["name"])
    log.info("Cleanup done.")


# ── First-run setup wizard ────────────────────────────────────────────────────

def run_setup():
    print("\n=== ETIMS Backup Uploader – First-time Setup ===\n")

    station_type  = input("Station type  (e.g. Rubis, Shell, Total, Regnol): ").strip()
    station_name  = input("Station name  (e.g. Eldoret Town, Nakuru West):    ").strip()
    root_folder   = input("Google Drive root folder ID (paste from Drive URL): ").strip()
    svc_email     = input(f"Service account email [{DEFAULT_SA_EMAIL}]:         ").strip()
    schedule_time = input("Daily backup time (24h, default 01:00):            ").strip() or "01:00"

    config = {
        "service_account_file": "credentials.json",
        "service_account_email": svc_email or DEFAULT_SA_EMAIL,
        "station_type": station_type,
        "station_name": station_name,
        "drive_root_folder_id": root_folder,
        "schedule_time": schedule_time,
        "max_retries": 3,
        "retry_interval_seconds": 300,
        "max_wait_minutes": 120,
        "log_spreadsheet_name": "ETIMSBackupLog"
    }

    save_config(config)
    print("\n✓ config.json saved.")
    print("  Place credentials.json in the same folder, then run without --setup.\n")


# ── Entry point ───────────────────────────────────────────────────────────────

DEFAULT_SA_EMAIL = "kra-checker-service@kra-auto-checker.iam.gserviceaccount.com"

def main():
    parser = argparse.ArgumentParser(description="ETIMS Backup Uploader")
    parser.add_argument("--setup",   action="store_true", help="Run the first-time setup wizard")
    parser.add_argument("--cleanup", action="store_true", help="Delete old Drive backups, keep latest N")
    parser.add_argument("--keep",    type=int, default=1,  help="Files to keep with --cleanup (default 1)")
    parser.add_argument("--test",          action="store_true", help="Test Drive connectivity and exit")
    parser.add_argument("--detect-backup", action="store_true", help="Auto-detect backup folder and print path")
    args = parser.parse_args()

    if args.setup:
        run_setup()
        return

    # ── Detect backup folder mode ──
    if getattr(args, 'detect_backup', False):
        from backup_finder import detect_backup_dir
        found = detect_backup_dir()
        if found:
            print(f"FOUND:{found}")
            sys.exit(0)
        else:
            print("NOTFOUND")
            sys.exit(1)

    config = load_config()

    # ── Wait for internet ──
    if not is_internet_available():
        if not wait_for_internet(config.get("max_wait_minutes", 120)):
            log.critical("Internet unavailable after waiting. Exiting.")
            # We can't log to Drive either if there's no internet – local log is the record.
            sys.exit(1)

    creds_path = BASE_DIR / config.get("service_account_file", "credentials.json")
    drive = DriveService(str(creds_path))

    # ── Test connectivity mode ──
    if args.test:
        try:
            root_id = config["drive_root_folder_id"]
            drive.get_or_create_folder("_etims_test_ping", root_id)
            drive.delete_file(
                drive._service.files().list(
                    q=f"name='_etims_test_ping' and '{root_id}' in parents and trashed=false",
                    fields="files(id)"
                ).execute().get("files", [{}])[0].get("id", "")
            )
            log.info("Connection test PASSED.")
            sys.exit(0)
        except Exception as exc:
            log.error("Connection test FAILED: %s", exc)
            sys.exit(1)

    # ── Cleanup mode ──
    if args.cleanup:
        cleanup_old_files(config, drive, keep=args.keep)
        return

    # -- Normal upload with retries --
    # Config reloaded each attempt so edits to config.json take effect without restart
    max_retries    = config.get("max_retries", 3)
    retry_interval = config.get("retry_interval_seconds", 300)

    for attempt in range(1, max_retries + 1):
        config = load_config()
        creds_path = BASE_DIR / config.get("service_account_file", "credentials.json")
        drive = DriveService(str(creds_path))

        log.info("=== Attempt %d / %d ===", attempt, max_retries)
        if run_backup(config, drive):
            break
        if attempt < max_retries:
            retry_interval = config.get("retry_interval_seconds", 300)
            log.info("Waiting %d s before next attempt...", retry_interval)
            time.sleep(retry_interval)
    else:
        log.critical("All %d attempts failed. Check ETIMSBackupLog on Drive.", max_retries)


if __name__ == "__main__":
    main()
