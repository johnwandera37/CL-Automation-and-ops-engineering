"""
error_logger.py – Log critical failures to a Google Sheet (ETIMSBackupLog).

The sheet ID is stored in config.json as  log_sheet_id.
Columns:  A=Timestamp  B=Station Type  C=Station Name  D=Error Message

Failed rows are written with RED bold text via the Sheets batchUpdate API.
"""

import logging
from datetime import datetime

from google.oauth2 import service_account
from googleapiclient.discovery import build

log = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

ERROR_TEXT_FORMAT = {
    "textFormat": {
        "bold": True,
        "foregroundColor": {"red": 0.8, "green": 0.0, "blue": 0.0},
    }
}

HEADER_FORMAT = {
    "backgroundColor": {"red": 0.75, "green": 0.0, "blue": 0.0},
    "textFormat": {
        "bold": True,
        "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
    },
    "horizontalAlignment": "CENTER",
}


def _get_sheets_service(credentials_path: str):
    creds = service_account.Credentials.from_service_account_file(
        credentials_path, scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _ensure_headers(service, sheet_id: str):
    """Write header row if the sheet is empty."""
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id, range="Sheet1!A1:D1"
    ).execute()

    if result.get("values"):
        return

    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range="Sheet1!A1:D1",
        valueInputOption="RAW",
        body={"values": [["Timestamp", "Station Type", "Station Name", "Error Message"]]}
    ).execute()

    service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={
            "requests": [
                {
                    "repeatCell": {
                        "range": {"sheetId": 0, "startRowIndex": 0, "endRowIndex": 1,
                                  "startColumnIndex": 0, "endColumnIndex": 4},
                        "cell": {"userEnteredFormat": HEADER_FORMAT},
                        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
                    }
                },
                *[{
                    "updateDimensionProperties": {
                        "range": {"sheetId": 0, "dimension": "COLUMNS",
                                  "startIndex": i, "endIndex": i + 1},
                        "properties": {"pixelSize": w},
                        "fields": "pixelSize",
                    }
                } for i, w in enumerate([180, 140, 180, 500])]
            ]
        }
    ).execute()


def log_critical_error(drive, config: dict, message: str):
    """
    Append a red error row to the Google Sheet in config['log_sheet_id'].
    'drive' arg kept for API compatibility but Sheets uses its own service.
    """
    sheet_id = config.get("log_sheet_id", "").strip()
    if not sheet_id:
        log.error(
            "log_sheet_id is not set in config.json – cannot write to error log. "
            "Create a Google Sheet, share it with the service account as Editor, "
            "then paste the sheet ID into config.json."
        )
        return

    creds_path   = config.get("service_account_file", "credentials.json")
    station_type = config.get("station_type", "Unknown")
    station_name = config.get("station_name", "Unknown")
    timestamp    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        service = _get_sheets_service(creds_path)

        _ensure_headers(service, sheet_id)

        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id, range="Sheet1!A:A"
        ).execute()
        next_row = len(result.get("values", [])) + 1

        service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"Sheet1!A{next_row}:D{next_row}",
            valueInputOption="RAW",
            body={"values": [[timestamp, station_type, station_name, message]]}
        ).execute()

        service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={
                "requests": [{
                    "repeatCell": {
                        "range": {"sheetId": 0,
                                  "startRowIndex": next_row - 1, "endRowIndex": next_row,
                                  "startColumnIndex": 0, "endColumnIndex": 4},
                        "cell": {"userEnteredFormat": ERROR_TEXT_FORMAT},
                        "fields": "userEnteredFormat.textFormat",
                    }
                }]
            }
        ).execute()

        log.info("Error logged to Google Sheet (row %d).", next_row)

    except Exception as exc:
        log.error("Could not write to Google Sheet error log: %s", exc, exc_info=True)
