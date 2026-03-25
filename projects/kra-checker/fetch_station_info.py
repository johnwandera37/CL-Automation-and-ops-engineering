"""
Fetch Station Info
Looks up station name in the Automation Helper spreadsheet using the AnyDesk
code as the key. If the station is not found and a name is provided manually,
it writes the new entry back to the Station Mapping sheet automatically.

Usage (standalone):
    python fetch_station_info.py <credentials_file> <automation_helper_sheet_id>
"""

import sys
import os
import logging

# Suppress the noisy file_cache warning from google library
logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.ERROR)
logger = logging.getLogger(__name__)

try:
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
except ImportError:
    print("ERROR: Google API libraries not installed.")
    sys.exit(1)


def _get_service(credentials_file: str):
    creds = service_account.Credentials.from_service_account_file(
        credentials_file,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return build("sheets", "v4", credentials=creds)


def fetch_station_by_anydesk(anydesk_code, credentials_file, automation_helper_id):
    """
    Look up station name by AnyDesk code in the 'Station Mapping' sheet.
    Sheet layout: | Station Name | AnyDesk Code |
    Returns station name string, or None if not found.
    """
    try:
        service = _get_service(credentials_file)
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
            if str(row[1]).strip() == str(anydesk_code).strip():
                print(f"Found station: {row[0].strip()}")
                return row[0].strip()
        print(f"No station found for AnyDesk code: {anydesk_code}")
        return None
    except Exception as e:
        print(f"Error fetching station info: {e}")
        return None


def add_station_to_sheet(station_name, anydesk_code, credentials_file, automation_helper_id):
    """
    Append a new station row to the 'Station Mapping' sheet.
    Called when a station is entered manually during installation.
    Returns True on success.
    """
    try:
        service = _get_service(credentials_file)
        service.spreadsheets().values().append(
            spreadsheetId=automation_helper_id,
            range="Station Mapping!A:B",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": [[station_name, str(anydesk_code)]]},
        ).execute()
        print(f"Added '{station_name}' ({anydesk_code}) to Station Mapping sheet")
        return True
    except Exception as e:
        print(f"Could not write to Station Mapping sheet: {e}")
        return False


if __name__ == "__main__":
    if not os.path.exists("detected_anydesk.txt"):
        print("ERROR: detected_anydesk.txt not found. Run anydesk_detector.py first.")
        sys.exit(1)

    with open("detected_anydesk.txt") as f:
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
        sys.exit(0)
    else:
        sys.exit(1)
