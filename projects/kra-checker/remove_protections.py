"""
Remove All Sheet Protections
Removes ALL protections from the four monitoring sheets.
Run once from Builder folder:
    python remove_protections.py
"""

import json
import os
import sys
import logging

logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.ERROR)

try:
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
except ImportError:
    print("ERROR: Google API libraries not installed.")
    sys.exit(1)

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
CREDS_PATH  = os.path.join(BASE_DIR, "credentials.json")

SHEETS_TO_FIX = ["Report", "Logs", "Station Status", "Heartbeat Error Logs"]


def main():
    if not os.path.exists(CONFIG_PATH):
        print(f"ERROR: config.json not found at {CONFIG_PATH}")
        sys.exit(1)
    if not os.path.exists(CREDS_PATH):
        print(f"ERROR: credentials.json not found at {CREDS_PATH}")
        sys.exit(1)

    with open(CONFIG_PATH) as f:
        config = json.load(f)

    spreadsheet_id = config.get("spreadsheet_id")
    if not spreadsheet_id:
        print("ERROR: spreadsheet_id not in config.json")
        sys.exit(1)

    creds = service_account.Credentials.from_service_account_file(
        CREDS_PATH,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    service = build("sheets", "v4", credentials=creds)

    print(f"\nSpreadsheet: {spreadsheet_id}")
    print(f"Removing protections from: {', '.join(SHEETS_TO_FIX)}\n")

    # Get full spreadsheet metadata including all protections
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        includeGridData=False
    ).execute()

    requests = []

    for sheet in meta.get("sheets", []):
        title    = sheet["properties"]["title"]
        sheet_id = sheet["properties"]["sheetId"]

        if title not in SHEETS_TO_FIX:
            continue

        # Collect all protected ranges on this sheet
        protected_ranges = sheet.get("protectedRanges", [])

        if not protected_ranges:
            print(f"  [--] {title} — no protections found")
            continue

        for prot in protected_ranges:
            prot_id = prot["protectedRangeId"]
            desc    = prot.get("description", "no description")
            requests.append({
                "deleteProtectedRange": {
                    "protectedRangeId": prot_id
                }
            })
            print(f"  [QUEUED] {title} — removing protection: {desc}")

    if not requests:
        print("\nNo protections found on any sheet. Nothing to do.")
        input("\nPress Enter to exit...")
        return

    # Also unprotect the sheet itself (sheet-level lock)
    # This requires updating each sheet's properties
    for sheet in meta.get("sheets", []):
        title    = sheet["properties"]["title"]
        sheet_id = sheet["properties"]["sheetId"]

        if title not in SHEETS_TO_FIX:
            continue

        # Remove sheet-level hidden protection by updating properties
        requests.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId"   : sheet_id,
                    "hidden"    : False,
                    "tabColor"  : {},
                },
                "fields": "hidden"
            }
        })

    print(f"\nApplying {len(requests)} change(s)...")

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()

    print("\n[OK] All protections removed.")
    print("The sheets should now be fully editable again.")
    print("\nNOTE: The new exes will NOT re-apply any protection.")
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
