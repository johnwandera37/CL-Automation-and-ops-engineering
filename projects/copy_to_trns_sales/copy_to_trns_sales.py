"""
KRA JSON File Recovery Script
===============================
Copies JSON files from processedArchive date folders back to trnsSales
for resubmission to KRA.

The script will interactively ask for:
  - Station PIN folder path
  - Start date (required)
  - End date (optional — leave blank to copy everything from start date onwards)

Rules:
- Only processes date folders within the specified date range
- Only copies JSON files with date modified within the specified range
- Skips files that already exist in trnsSales
- Logs all actions to a log file
"""

import os
import shutil
import logging
from datetime import datetime


def get_station_path() -> str:
    """Prompt the user to enter the KRA PIN base folder path."""
    print("\n" + "=" * 60)
    print("  KRA JSON File Recovery Script")
    print("=" * 60)
    print("\nEnter the full path to the station's KRA PIN folder.")
    print("Example: C:\\Users\\Administrator\\AppData\\EbmData\\P052030219E_01")
    print()

    while True:
        path = input("Station PIN folder path: ").strip().strip('"').strip("'")

        if not path:
            print("  [!] Path cannot be empty. Please try again.\n")
            continue

        if not os.path.isdir(path):
            print(f"  [!] Folder not found: {path}")
            print("      Please check the path and try again.\n")
            continue

        return path


def get_dates() -> tuple:
    """Interactively prompt the user for start and optional end dates."""
    print("\n" + "-" * 60)
    print("  Date Range Selection")
    print("-" * 60)
    print("  Format: YYYYMMDD  e.g. 20260101 for 01/01/2026")
    print()

    # --- Start date (required) ---
    start_date = None
    start_str = None
    while True:
        raw = input("  Start date (required): ").strip()
        if not raw:
            print("  [!] Start date cannot be empty. Please try again.\n")
            continue
        try:
            start_date = datetime.strptime(raw, "%Y%m%d")
            start_str = raw
            break
        except ValueError:
            print(f"  [!] Invalid date '{raw}'. Please use YYYYMMDD format e.g. 20260101\n")

    # --- End date (optional) ---
    end_date = None
    end_str = None
    while True:
        raw = input("  End date   (optional, leave blank for no end): ").strip()
        if not raw:
            print("  [i] No end date set — will copy from start date onwards.")
            break
        try:
            end_date = datetime.strptime(raw, "%Y%m%d")
            if end_date < start_date:
                print(f"  [!] End date cannot be before start date ({start_str}). Please try again.\n")
                end_date = None
                continue
            # Set to end of day so files modified anytime on that day are included
            end_date = end_date.replace(hour=23, minute=59, second=59)
            end_str = raw
            break
        except ValueError:
            print(f"  [!] Invalid date '{raw}'. Please use YYYYMMDD format e.g. 20260131\n")

    return start_date, end_date, start_str, end_str


def setup_logging(base_path: str) -> str:
    """Set up logging to both console and a log file saved inside the PIN folder."""
    log_filename = f"kra_copy_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    log_path = os.path.join(base_path, log_filename)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler()
        ]
    )
    return log_path


def validate_structure(pin_folder: str):
    """Check that the expected folder structure exists and return key paths."""
    processed_archive = os.path.join(pin_folder, "Data", "processed", "processedArchive")
    trns_sales = os.path.join(pin_folder, "Data", "resend", "trnsSales")

    errors = []
    if not os.path.isdir(processed_archive):
        errors.append(f"processedArchive folder not found:\n    {processed_archive}")
    if not os.path.isdir(trns_sales):
        errors.append(f"trnsSales folder not found:\n    {trns_sales}")

    if errors:
        print("\n  [!] Folder structure validation failed:")
        for e in errors:
            print(f"      - {e}")
        print("\n  Please verify the path is correct and try again.")
        raise SystemExit(1)

    return processed_archive, trns_sales


def is_valid_date_folder(folder_name: str, start_str: str, end_str) -> bool:
    """Return True if the folder name is a YYYYMMDD date within the specified range."""
    if len(folder_name) != 8 or not folder_name.isdigit():
        return False
    if folder_name < start_str:
        return False
    if end_str and folder_name > end_str:
        return False
    return True


def is_within_date_range(modified: datetime, start_date: datetime, end_date) -> bool:
    """Return True if the file's modified date falls within the date range."""
    if modified < start_date:
        return False
    if end_date and modified > end_date:
        return False
    return True


def get_file_modified_datetime(filepath: str) -> datetime:
    """Return the file's last modified datetime (local, naive)."""
    mtime = os.path.getmtime(filepath)
    return datetime.fromtimestamp(mtime)


def copy_json_files(processed_archive: str, trns_sales: str,
                    start_date: datetime, end_date, start_str: str, end_str) -> dict:
    """Main copy logic. Returns a summary dict with counts."""
    summary = {
        "date_folders_scanned": 0,
        "files_checked": 0,
        "files_copied": 0,
        "files_skipped_existing": 0,
        "files_skipped_out_of_range": 0,
        "errors": 0,
    }

    # Get and sort all date folders within range
    try:
        all_entries = os.listdir(processed_archive)
    except PermissionError:
        logging.error(f"Permission denied accessing: {processed_archive}")
        raise SystemExit(1)

    date_folders = sorted([
        f for f in all_entries
        if os.path.isdir(os.path.join(processed_archive, f))
        and is_valid_date_folder(f, start_str, end_str)
    ])

    if not date_folders:
        range_msg = f">= {start_str}" if not end_str else f"{start_str} → {end_str}"
        logging.warning(f"No valid date folders found in processedArchive ({range_msg}).")
        return summary

    logging.info(f"Found {len(date_folders)} date folder(s) to process: {date_folders[0]} → {date_folders[-1]}")
    logging.info(f"Destination (trnsSales): {trns_sales}")
    logging.info("-" * 60)

    for folder_name in date_folders:
        folder_path = os.path.join(processed_archive, folder_name)
        summary["date_folders_scanned"] += 1

        try:
            json_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".json")]
        except PermissionError:
            logging.error(f"Permission denied reading folder: {folder_path}")
            summary["errors"] += 1
            continue

        if not json_files:
            logging.info(f"[{folder_name}] No JSON files found — skipping.")
            continue

        logging.info(f"[{folder_name}] Found {len(json_files)} JSON file(s).")

        for filename in json_files:
            src = os.path.join(folder_path, filename)
            dst = os.path.join(trns_sales, filename)
            summary["files_checked"] += 1

            # Check file modified date
            try:
                modified = get_file_modified_datetime(src)
            except Exception as e:
                logging.error(f"  Could not read modified date for {filename}: {e}")
                summary["errors"] += 1
                continue

            if not is_within_date_range(modified, start_date, end_date):
                logging.info(f"  SKIP (out of range {modified.strftime('%Y-%m-%d')}) → {filename}")
                summary["files_skipped_out_of_range"] += 1
                continue

            # Check if already exists in trnsSales
            if os.path.exists(dst):
                logging.info(f"  SKIP (already exists in trnsSales) → {filename}")
                summary["files_skipped_existing"] += 1
                continue

            # Copy the file
            try:
                shutil.copy2(src, dst)  # copy2 preserves metadata
                logging.info(f"  COPIED (modified {modified.strftime('%Y-%m-%d %H:%M')}) → {filename}")
                summary["files_copied"] += 1
            except Exception as e:
                logging.error(f"  ERROR copying {filename}: {e}")
                summary["errors"] += 1

    return summary


def print_summary(summary: dict, log_path: str):
    """Print a final summary to console and log."""
    msg = f"""
{'=' * 60}
  COPY OPERATION COMPLETE
{'=' * 60}
  Date folders scanned  : {summary['date_folders_scanned']}
  JSON files checked    : {summary['files_checked']}
  Files COPIED          : {summary['files_copied']}
  Skipped (out of range): {summary['files_skipped_out_of_range']}
  Skipped (exist)       : {summary['files_skipped_existing']}
  Errors                : {summary['errors']}
{'=' * 60}
  Log saved to: {log_path}
{'=' * 60}
"""
    logging.info(msg)


def main():
    # Step 1: Get station path from user
    pin_folder = get_station_path()

    # Step 2: Get date range interactively
    start_date, end_date, start_str, end_str = get_dates()

    # Step 3: Set up logging
    log_path = setup_logging(pin_folder)

    # Step 4: Log startup info
    logging.info(f"Script started. Station folder: {pin_folder}")
    if end_date:
        logging.info(f"Date range: {start_date.strftime('%Y-%m-%d')} → {end_date.strftime('%Y-%m-%d')} (inclusive)")
    else:
        logging.info(f"Date range: {start_date.strftime('%Y-%m-%d')} onwards")

    # Step 5: Validate folder structure
    processed_archive, trns_sales = validate_structure(pin_folder)
    logging.info(f"processedArchive: {processed_archive}")
    logging.info(f"trnsSales       : {trns_sales}")

    # Step 6: Confirm before proceeding
    print(f"\n  processedArchive : {processed_archive}")
    print(f"  trnsSales        : {trns_sales}")
    if end_date:
        print(f"\n  Date range : {start_date.strftime('%d/%m/%Y')} → {end_date.strftime('%d/%m/%Y')} (inclusive)")
    else:
        print(f"\n  Date range : {start_date.strftime('%d/%m/%Y')} onwards")
    print("  Files that already exist in trnsSales will be SKIPPED.\n")

    confirm = input("  Type YES to proceed: ").strip()
    if confirm != "YES":
        logging.info("Operation cancelled by user.")
        print("\n  Operation cancelled.")
        return

    # Step 7: Copy files
    logging.info("User confirmed. Starting copy operation...")
    summary = copy_json_files(processed_archive, trns_sales,
                               start_date, end_date, start_str, end_str)

    # Step 8: Print summary
    print_summary(summary, log_path)


if __name__ == "__main__":
    main()