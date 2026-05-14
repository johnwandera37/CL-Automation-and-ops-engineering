"""
Clear Sheets Utility
Clears all data rows (keeps headers) from the four monitoring sheets.
Run once from the Builder folder:
    python clear_sheets.py
"""

import json
import os
import sys

logging_suppress = True
import logging
logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.ERROR)

try:
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
except ImportError:
    print("ERROR: Google API libraries not installed.")
    sys.exit(1)

# ── Config ────────────────────────────────────────────────────────────────────
# Point these to your local config.json and credentials.json
BASE_DIR       = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH    = os.path.join(BASE_DIR, "config.json")
CREDS_PATH     = os.path.join(BASE_DIR, "credentials.json")

SHEETS_TO_CLEAR = [
    {
        "name"    : "Report",
        "headers" : ["Timestamp", "Station", "AnyDesk", "Check Date",
                     "Status", "Trans Date", "Invoice", "QR Link", "Details"],
    },
    {
        "name"    : "Logs",
        "headers" : ["Timestamp", "Station", "AnyDesk", "Level", "Message"],
    },
    {
        "name"    : "Station Status",
        "headers" : ["Station", "AnyDesk", "Last Seen", "Status",
                     "IP Address", "Disk Space", "SQL Server", "CPU Temp",
                     "Heartbeat Interval (min)"],
    },
    {
        "name"    : "Heartbeat Error Logs",
        "headers" : ["Timestamp", "Station", "AnyDesk", "Level", "Message"],
    },
]


def main():
    # ── Load config ───────────────────────────────────────────────────────────
    if not os.path.exists(CONFIG_PATH):
        print(f"ERROR: config.json not found at {CONFIG_PATH}")
        print("Copy config.json to the same folder as this script.")
        sys.exit(1)

    if not os.path.exists(CREDS_PATH):
        print(f"ERROR: credentials.json not found at {CREDS_PATH}")
        sys.exit(1)

    with open(CONFIG_PATH) as f:
        config = json.load(f)

    spreadsheet_id = config.get("spreadsheet_id")
    if not spreadsheet_id:
        print("ERROR: spreadsheet_id not found in config.json")
        sys.exit(1)

    # ── Authenticate ──────────────────────────────────────────────────────────
    creds = service_account.Credentials.from_service_account_file(
        CREDS_PATH,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    service = build("sheets", "v4", credentials=creds)

    print(f"\nSpreadsheet: {spreadsheet_id}")
    print()

    # ── Confirm ───────────────────────────────────────────────────────────────
    print("This will clear ALL data rows from the following sheets:")
    for s in SHEETS_TO_CLEAR:
        print(f"  • {s['name']}")
    print("\nHeaders will be preserved. This cannot be undone.")
    print()
    confirm = input("Type YES to confirm: ").strip()
    if confirm.upper() != "YES":
        print("Cancelled.")
        return

    print()

    # ── Clear each sheet ──────────────────────────────────────────────────────
    for sheet in SHEETS_TO_CLEAR:
        name    = sheet["name"]
        headers = sheet["headers"]
        col     = _col_letter(len(headers) - 1)

        try:
            # Step 1: Check sheet exists
            meta     = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            existing = [s["properties"]["title"] for s in meta.get("sheets", [])]
            if name not in existing:
                print(f"  [--] {name} — sheet not found, skipping")
                continue

            # Step 2: Remove any data protection first
            sheet_meta = next(
                (s for s in meta.get("sheets", [])
                 if s["properties"]["title"] == name), None
            )
            if sheet_meta:
                del_requests = []
                for prot in sheet_meta.get("protectedRanges", []):
                    del_requests.append({
                        "deleteProtectedRange": {
                            "protectedRangeId": prot["protectedRangeId"]
                        }
                    })
                if del_requests:
                    service.spreadsheets().batchUpdate(
                        spreadsheetId=spreadsheet_id,
                        body={"requests": del_requests}
                    ).execute()
                    print(f"  [OK] {name} — protection removed")

            # Step 3: Clear all content
            service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id,
                range=f"{name}!A:Z"
            ).execute()

            # Step 4: Restore headers
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{name}!A1:{col}1",
                valueInputOption="RAW",
                body={"values": [headers]}
            ).execute()

            print(f"  [OK] {name} — cleared, headers restored")

        except Exception as e:
            print(f"  [ERROR] {name} — {e}")


    # Clean up local flag files so protection and date separators
    # are re-applied fresh on the next program run
    print("\n  Cleaning up local flag files...")
    import glob
    cleaned = 0
    for pattern in [".protected_*", ".lastdate_*"]:
        for flag in glob.glob(os.path.join(BASE_DIR, pattern)):
            try:
                os.remove(flag)
                cleaned += 1
            except Exception:
                pass
    if cleaned:
        print(f"  [OK] Removed {cleaned} local flag file(s)")

    print("\nDone. All sheets cleared and headers restored.")
    print()


def _col_letter(idx: int) -> str:
    result = ""
    while idx >= 0:
        result = chr(idx % 26 + 65) + result
        idx = idx // 26 - 1
    return result


if __name__ == "__main__":
    main()
