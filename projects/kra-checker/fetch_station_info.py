"""
Fetch Station Info
Looks up the station name in the Automation Helper spreadsheet
using the local AnyDesk code as the key.

Usage (standalone):
    python fetch_station_info.py <credentials_file> <automation_helper_sheet_id>
"""

import json
import sys
import os
import logging

logger = logging.getLogger(__name__)

try:
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
except ImportError:
    print("ERROR: Google API libraries not installed. Run: pip install google-api-python-client google-auth")
    sys.exit(1)


def fetch_station_by_anydesk(
    anydesk_code: str,
    credentials_file: str,
    automation_helper_id: str,
) -> str | None:
    """
    Look up the station name in the 'Station Mapping' sheet.
    Sheet layout expected:  | Station Name | AnyDesk Code |
    Returns station name string, or None if not found.
    """
    try:
        creds = service_account.Credentials.from_service_account_file(
            credentials_file,
            scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
        )
        service = build("sheets", "v4", credentials=creds)

        result = (
            service.spreadsheets()
            .values()
            .get(spreadsheetId=automation_helper_id, range="Station Mapping!A:B")
            .execute()
        )

        rows = result.get("values", [])

        for row in rows[1:]:  # skip header
            if len(row) < 2:
                continue
            station_name = row[0].strip()
            sheet_anydesk = str(row[1]).strip()

            if sheet_anydesk == str(anydesk_code).strip():
                print(f"✓ Found station: {station_name}")
                return station_name

        print(f"✗ No station found for AnyDesk code: {anydesk_code}")
        return None

    except Exception as e:
        print(f"✗ Error fetching station info: {e}")
        return None


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    if not os.path.exists("detected_anydesk.txt"):
        print("ERROR: detected_anydesk.txt not found. Run anydesk_detector.py first.")
        sys.exit(1)

    with open("detected_anydesk.txt", "r") as f:
        anydesk_code = f.read().strip()

    if len(sys.argv) < 3:
        print("Usage: python fetch_station_info.py <credentials_file> <automation_helper_sheet_id>")
        sys.exit(1)

    credentials_file = sys.argv[1]
    automation_helper_id = sys.argv[2]

    station_name = fetch_station_by_anydesk(anydesk_code, credentials_file, automation_helper_id)

    if station_name:
        with open("detected_station.txt", "w") as f:
            f.write(station_name)
        print(f"Saved station name to detected_station.txt")
        sys.exit(0)
    else:
        sys.exit(1)
