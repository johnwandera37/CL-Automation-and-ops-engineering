"""
KRA Auto-Checker  v2.1
Runs daily at the configured time via Windows Task Scheduler (VBS launcher).
  - Queries SQL Server for a random ETIMS transaction
  - Checks the QR link against the KRA portal
  - Writes result to the Report sheet
  - Logs all steps to the Logs sheet
  - Saves failed transactions for overnight retry
Updates are handled by heartbeat_monitor.exe — not this program.
"""

import pyodbc
import requests
import re
import json
import logging
import time
import sys
import os
from datetime import datetime
from typing import Optional, Tuple, Dict

GLOBAL_LOGGER = None
logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.ERROR)

# ── Hide console when running via Task Scheduler ──────────────────────────────
def _hide_console():
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)
    except Exception:
        pass

_hide_console()

# ── Updates handled by heartbeat_monitor (runs every 30 min) ─────────────────

BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

try:
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
except ImportError as e:
    print(f"Missing library: {e}")
    sys.exit(1)

from config_loader import ConfigLoader


# ─────────────────────────────────────────────────────────────────────────────
# Sheet Logger
# ─────────────────────────────────────────────────────────────────────────────

class SheetLogger:
    """Writes log entries to both console and Google Sheets Logs tab."""

    LEVEL_EMOJI = {
        "INFO"   : "🟢 INFO",
        "WARNING": "🟡 WARNING",
        "ERROR"  : "🔴 ERROR",
        "SUCCESS": "🟢 SUCCESS",
        "UPDATE" : "🔄 UPDATE",
        "DEBUG"  : "⚪ DEBUG",
    }

    def __init__(self, sheets_manager, config: ConfigLoader):
        self.sheets  = sheets_manager
        self.config  = config
        self._log    = logging.getLogger(__name__)
        self.log_buffer = []

    def _write(self, level: str, message: str):
        emoji = {"INFO": "🟢", "WARNING": "🟡", "ERROR": "🔴", "SUCCESS": "🟢"}.get(level, "⚪")
        getattr(self._log, level.lower() if level != "SUCCESS" else "info")(
            f"{emoji} {message}"
        )

        self.log_buffer.append(
            (
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),  # exact time captured now
                level,
                message
            )
        )

    def info(self, m):    self._write("INFO", m)
    def warning(self, m): self._write("WARNING", m)
    def error(self, m):   self._write("ERROR", m)
    def success(self, m): self._write("SUCCESS", m)

    def flush(self):
        if not self.log_buffer:
            return
        try:
            self.sheets.add_log_entries(self.log_buffer)
        except Exception:
            pass

        self.log_buffer.clear()


# ─────────────────────────────────────────────────────────────────────────────
# KRA Link Checker
# ─────────────────────────────────────────────────────────────────────────────

class KRAChecker:

    def __init__(self, config: ConfigLoader, logger: SheetLogger):
        self.config  = config
        self.logger  = logger
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        })

    def check_kra_link(self, qr_link: str) -> Tuple[str, str, Optional[str], Optional[str]]:
        """
        Check a QR link against the KRA portal.
        Returns (status, details, transaction_date, invoice_number).
        status: SUCCESS | NOT_SUBMITTED | ERROR
        """
        max_retries = int(self.config.get("max_retries", 3))
        timeout     = int(self.config.get("timeout", 15))
        retry_delay = int(self.config.get("retry_delay", 30))

        for attempt in range(max_retries + 1):
            try:
                response = self.session.get(qr_link, timeout=timeout)
                if response.status_code == 200:
                    return self._parse_response(response.text)
                elif response.status_code in (500, 502, 503, 504):
                    return ("ERROR", f"KRA server error {response.status_code}", None, None)
                else:
                    return ("ERROR", f"HTTP {response.status_code}", None, None)

            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                if attempt < max_retries:
                    self.logger.warning(f"Network error (attempt {attempt+1}/{max_retries}), retrying...")
                    time.sleep(retry_delay)
                else:
                    kind = "Timeout" if isinstance(e, requests.exceptions.Timeout) else "Connection error"
                    return ("ERROR", f"{kind} after {attempt+1} attempts", None, None)
            except Exception as e:
                return ("ERROR", f"Unexpected error: {e}", None, None)

        return ("ERROR", "Max retries exceeded", None, None)

    def _parse_response(self, html: str) -> Tuple[str, str, Optional[str], Optional[str]]:
        lower = html.lower()
        if "invoice number" in lower and "scu information" in lower:
            return ("SUCCESS", "Transaction submitted to KRA",
                    self._extract_date(html), self._extract_invoice(html))
        if "could not be verified" in lower or "try again later" in lower:
            return ("NOT_SUBMITTED", "Invoice not verified by KRA", None, None)
        return ("ERROR", "Page loaded but status unclear", None, None)

    @staticmethod
    def _extract_date(html: str) -> Optional[str]:
        for p in [
            r'<span[^>]*>(\d{1,2})/(\d{1,2})/(\d{4})\s+\d{2}:\d{2}:\d{2}</span>',
            r'(\d{1,2})/(\d{1,2})/(\d{4})\s+\d{2}:\d{2}:\d{2}',
        ]:
            m = re.search(p, html)
            if m:
                return f"{int(m.group(2))}/{int(m.group(1))}/{m.group(3)}"
        return None

    @staticmethod
    def _extract_invoice(html: str) -> Optional[str]:
        m = re.search(r'Invoice Number\s*[:\s]+([A-Z0-9/]+)', html, re.IGNORECASE)
        return m.group(1).strip() if m else None


# ─────────────────────────────────────────────────────────────────────────────
# Database Manager
# ─────────────────────────────────────────────────────────────────────────────

class DatabaseManager:

    def __init__(self, config: ConfigLoader, logger: SheetLogger):
        self.config = config
        self.logger = logger

    def _conn_str(self) -> str:
        return (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={self.config.get('sql_server', '.\\SQLEXPRESS')};"
            f"DATABASE={self.config.get('sql_database', 'ETIMS')};"
            f"UID={self.config.get('sql_username', 'sa')};"
            f"PWD={self.config.get('sql_password', '')}"
        )

    def get_random_transactions(self, check_date: datetime) -> Dict[str, Optional[Dict]]:
        """
        Fetch:
        - one Fuel Card transaction
        - one Non-Fuel-Card transaction

        Returns:
        {
            "fuel_card": {...} or None,
            "other": {...} or None
        }
        """

        time_ranges = [
            ("16:00:00", "19:00:00", "4 PM – 7 PM"),
            ("12:00:00", "19:00:00", "12 PM – 7 PM"),
            ("08:00:00", "19:00:00", "8 AM – 7 PM"),
            ("00:00:00", "23:59:59", "any time today"),
        ]

        queries = {
            "fuel_card": """
                SELECT TOP 1
                    TransDateTime,
                    QRLink,
                    PaymentMode
                FROM ETPumpSales
                WHERE TransDateTime BETWEEN ? AND ?
                AND QRLink IS NOT NULL
                AND LTRIM(RTRIM(QRLink)) <> ''
                AND PaymentMode = 'Fuel Card'
                ORDER BY NEWID()
            """,

            "other": """
                SELECT TOP 1
                    TransDateTime,
                    QRLink,
                    PaymentMode
                FROM ETPumpSales
                WHERE TransDateTime BETWEEN ? AND ?
                AND QRLink IS NOT NULL
                AND LTRIM(RTRIM(QRLink)) <> ''
                AND PaymentMode <> 'Fuel Card'
                ORDER BY NEWID()
            """
        }

        results = {
            "fuel_card": None,
            "other": None
        }

        try:
            conn = pyodbc.connect(self._conn_str())
            cursor = conn.cursor()

            for tx_type, query in queries.items():

                for start_t, end_t, label in time_ranges:

                    t_start = check_date.strftime(f"%Y-%m-%d {start_t}")
                    t_end   = check_date.strftime(f"%Y-%m-%d {end_t}")

                    cursor.execute(query, (t_start, t_end))
                    row = cursor.fetchone()

                    if row:
                        self.logger.info(
                            f"Found {tx_type} transaction ({label}): {row[0]}"
                        )

                        results[tx_type] = {
                            "TransDateTime": row[0],
                            "QRLink": row[1],
                            "PaymentMode": row[2],
                            "CheckDate": check_date.strftime("%Y-%m-%d"),
                        }

                        break
                    else:
                        self.logger.info(
                            f"No {tx_type.replace('_', ' ').title()} transactions in {label}, widening search…"
                        )

                if not results[tx_type]:
                    self.logger.warning(
                        f"No {tx_type.replace('_', ' ').title()} transaction found"
                    )

            cursor.close()
            conn.close()

            return results

        except Exception as e:
            self.logger.error(f"Database error: {e}")
            return results


# ─────────────────────────────────────────────────────────────────────────────
# Google Sheets Manager
# ─────────────────────────────────────────────────────────────────────────────

class GoogleSheetsManager:

    REPORT_HEADERS = [
        "Timestamp", "Station", "AnyDesk", "Check Date",
        "Status", "Trans Date", "Invoice", "QR Link", "Details",
    ]
    LOG_HEADERS = ["Timestamp", "Station", "AnyDesk", "Level", "Message"]

    COLUMN_WIDTHS = {
        "Report": [160, 180, 130, 100, 130, 100, 130, 300, 250],
        "Logs"  : [160, 180, 130, 90,  400],
    }

    LEVEL_EMOJI = {
        "INFO"   : "🟢 INFO",
        "WARNING": "🟡 WARNING",
        "ERROR"  : "🔴 ERROR",
        "SUCCESS": "🟢 SUCCESS",
        "UPDATE" : "🔄 UPDATE",
        "DEBUG"  : "⚪ DEBUG",
    }

    def __init__(self, config: ConfigLoader):
        self.config = config
        self._log   = logging.getLogger(__name__)
        self.service = self._authenticate()
        self._sheet_metadata_cache = None  # cached once per run
        self._sheet_id_cache       = {}    # sheet_name -> sheet_id

    def _get_spreadsheet_metadata(self):
        """Fetch and cache spreadsheet metadata — called at most once per run."""
        if self._sheet_metadata_cache is None:
            self._sheet_metadata_cache = self.service.spreadsheets().get(
                spreadsheetId=self.config.get("spreadsheet_id")
            ).execute()
        return self._sheet_metadata_cache

    def _authenticate(self):
        sa_file = self.config.get("service_account_file", "credentials.json")
        creds   = service_account.Credentials.from_service_account_file(
            sa_file, scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        return build("sheets", "v4", credentials=creds)

    # ── Public write methods ──────────────────────────────────────────

    def add_report_entries(self, entries: list):
        """Write multiple report rows in a single API call — keeps station rows together."""
        sid = self.config.get("spreadsheet_id")
        self._ensure_sheet("Report", self.REPORT_HEADERS)
        self._heal_schema("Report", self.REPORT_HEADERS)

        today = datetime.now().strftime("%d/%m/%Y")
        self._insert_date_separator("Report", today, len(self.REPORT_HEADERS))

        values = []
        for data in entries:
            raw = data.get("status", "")
            if "SUCCESS"         in raw: status_cell = "🟢 SUCCESS"
            elif "NOT_SUBMITTED" in raw: status_cell = "🔴 NOT SUBMITTED"
            elif "ERROR"         in raw: status_cell = "🟡 ERROR"
            elif "NO DATA"       in raw: status_cell = "⚪ NO DATA"
            else:                        status_cell = raw

            values.append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                self.config.get("station_name", ""),
                self.config.get("anydesk_code", ""),
                data.get("check_date", ""),
                status_cell,
                data.get("transaction_date", ""),
                data.get("invoice_number", ""),
                data.get("qr_link", ""),
                data.get("details", ""),
            ])

        self.service.spreadsheets().values().append(
            spreadsheetId=sid,
            range="Report!A:I",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": values},
        ).execute()
        self._auto_resize("Report")
        self._log.info(f"Report: {len(values)} entries written")

    # For backward compatibility, currently not being used
    def add_log_entry(self, level: str, message: str):
        try:
            sid = self.config.get("spreadsheet_id")
            self._ensure_sheet("Logs", self.LOG_HEADERS)

            today = datetime.now().strftime("%d/%m/%Y")
            self._insert_date_separator("Logs", today, len(self.LOG_HEADERS))

            level_display = self.LEVEL_EMOJI.get(level.upper(), level)
            values = [[
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                self.config.get("station_name", ""),
                self.config.get("anydesk_code", ""),
                level_display,
                message,
            ]]
            self.service.spreadsheets().values().append(
                spreadsheetId=sid,
                range="Logs!A:E",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": values},
            ).execute()
            self._auto_resize("Logs")
        except Exception:
            pass  # never crash main flow


    def add_log_entries(self, entries):
        """
        Bulk append log entries in a single API call.
        entries = [(timestamp, level, message), ...]
        """

        try:
            sid = self.config.get("spreadsheet_id")

            self._ensure_sheet("Logs", self.LOG_HEADERS)

            today = datetime.now().strftime("%d/%m/%Y")
            self._insert_date_separator("Logs", today, len(self.LOG_HEADERS))

            values = []
            for timestamp, level, message in entries:
                values.append([
                    timestamp,
                    self.config.get("station_name", ""),
                    self.config.get("anydesk_code", ""),
                    self.LEVEL_EMOJI.get(level.upper(), level),
                    message
                ])

            self.service.spreadsheets().values().append(
                spreadsheetId=sid,
                range="Logs!A:E",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": values}
            ).execute()

            self._auto_resize("Logs")

        except Exception:
            pass


    # ── Internal helpers ──────────────────────────────────────────────

    def _ensure_sheet(self, name: str, headers: list):
        meta     = self._get_spreadsheet_metadata()
        existing = [s["properties"]["title"] for s in meta.get("sheets", [])]
        if name in existing:
            return
        # Sheet doesn't exist — create it and invalidate cache
        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.config.get("spreadsheet_id"),
            body={"requests": [{"addSheet": {"properties": {"title": name}}}]},
        ).execute()
        self._sheet_metadata_cache = None  # invalidate so next call re-fetches
        col = self._col(len(headers) - 1)
        self.service.spreadsheets().values().update(
            spreadsheetId=self.config.get("spreadsheet_id"),
            range=f"{name}!A1:{col}1",
            valueInputOption="RAW",
            body={"values": [headers]},
        ).execute()
        self._log.info(f"Created sheet: {name}")

    def _heal_schema(self, sheet_name: str, expected: list):
        result  = self.service.spreadsheets().values().get(
            spreadsheetId=self.config.get("spreadsheet_id"),
            range=f"{sheet_name}!A1:Z1",
        ).execute()
        current = result.get("values", [[]])[0]
        if current != expected:
            self.service.spreadsheets().values().update(
                spreadsheetId=self.config.get("spreadsheet_id"),
                range=f"{sheet_name}!A1",
                valueInputOption="RAW",
                body={"values": [expected]},
            ).execute()
            self._log.warning(f"Auto-healed schema: {sheet_name}")

    def _get_sheet_id(self, sheet_name: str) -> int:
        if sheet_name not in self._sheet_id_cache:
            meta = self._get_spreadsheet_metadata()
            for s in meta.get("sheets", []):
                if s["properties"]["title"] == sheet_name:
                    self._sheet_id_cache[sheet_name] = s["properties"]["sheetId"]
                    return s["properties"]["sheetId"]
            raise ValueError(f"Sheet not found: {sheet_name}")
        return self._sheet_id_cache[sheet_name]

    
    """ 
    Let one station Write today's date to a known cell (Z1). All stations
    check that cell first, First station to write it wins, rest skip. One read call,
    one write call, cell acts as a shared flag across all stations
    """
    def _date_separator_written_today(self, sheet_name: str, date_str: str) -> bool:
        """Check shared flag cell — much faster than scanning column A."""
        try:
            flag_cell = "Z1" if sheet_name == "Report" else "F1"
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.config.get("spreadsheet_id"),
                range=f"{sheet_name}!{flag_cell}"
            ).execute()
            val = result.get("values", [[]])[0][0] if result.get("values") else ""
            return val.strip() == date_str
        except Exception:
            return False

    def _mark_separator_written(self, sheet_name: str, date_str: str):
        flag_cell = "Z1" if sheet_name == "Report" else "F1"
        try:
            self.service.spreadsheets().values().update(
                spreadsheetId=self.config.get("spreadsheet_id"),
                range=f"{sheet_name}!{flag_cell}",
                valueInputOption="RAW",
                body={"values": [[date_str]]}
            ).execute()
        except Exception:
            pass

    def _insert_date_separator(self, sheet_name: str, date_str: str, num_cols: int):
        """
        Appends a plain merged date row when the date changes.
        Uses a shared flag cell (Z1 for Report, F1 for Logs) visible to all stations.
        First station to run each day inserts the separator, rest skip.
        """
        if self._date_separator_written_today(sheet_name, date_str):
            return False

        try:
            sheet_id = self._get_sheet_id(sheet_name)
            self.service.spreadsheets().values().append(
                spreadsheetId=self.config.get("spreadsheet_id"),
                range=f"{sheet_name}!A:A",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": [[date_str]]}
            ).execute()

            rows = self.service.spreadsheets().values().get(
                spreadsheetId=self.config.get("spreadsheet_id"),
                range=f"{sheet_name}!A:A"
            ).execute().get("values", [])
            next_row = len(rows)

            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.config.get("spreadsheet_id"),
                body={"requests": [{
                    "mergeCells": {
                        "range": {
                            "sheetId"         : sheet_id,
                            "startRowIndex"   : next_row - 1,
                            "endRowIndex"     : next_row,
                            "startColumnIndex": 0,
                            "endColumnIndex"  : num_cols,
                        },
                        "mergeType": "MERGE_ALL"
                    }
                }]}
            ).execute()

            self._mark_separator_written(sheet_name, date_str)
            return True

        except Exception as e:
            self._log.warning(f"[DATE SEP] {sheet_name}: {e}")
            return False

    def _auto_resize(self, sheet_name: str):
        widths = self.COLUMN_WIDTHS.get(sheet_name)
        if not widths:
            return
        try:
            sheet_id = self._get_sheet_id(sheet_name)
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.config.get("spreadsheet_id"),
                body={"requests": [
                    {
                        "updateDimensionProperties": {
                            "range"     : {"sheetId": sheet_id, "dimension": "COLUMNS",
                                           "startIndex": i, "endIndex": i + 1},
                            "properties": {"pixelSize": px},
                            "fields"    : "pixelSize"
                        }
                    }
                    for i, px in enumerate(widths)
                ]}
            ).execute()
        except Exception as e:
            self._log.warning(f"Column resize failed for {sheet_name}: {e}")

    @staticmethod
    def _col(idx: int) -> str:
        result = ""
        while idx >= 0:
            result = chr(idx % 26 + 65) + result
            idx = idx // 26 - 1
        return result


# ─────────────────────────────────────────────────────────────────────────────
# Retry Manager
# ─────────────────────────────────────────────────────────────────────────────

class RetryManager:

    def __init__(self, config: ConfigLoader, logger: SheetLogger):
        self.config     = config
        self.logger     = logger
        self.retry_file = config.get(
            "retry_file", os.path.join(BASE_DIR, "retry_transaction.json")
        )

    def save(self, transactions: list, retry_count: int = 0):
        """
        Save failed transactions for retry.
        """

        data = {
            "retry_count": retry_count,
            "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "transactions": [],
        }

        for tx in transactions:

            data["transactions"].append({
                "transaction_date": str(tx["TransDateTime"]),
                "qr_link": tx["QRLink"],
                "check_date": tx["CheckDate"],
                "payment_mode": tx.get("PaymentMode", "UNKNOWN"),
                "last_status": tx.get("LastStatus", "ERROR"),
            })

        with open(self.retry_file, "w") as f:
            json.dump(data, f, indent=4)

        self.logger.info(
            f"Saved {len(data['transactions'])} transaction(s) for retry"
        )

    def load(self) -> Optional[Dict]:
        if not os.path.exists(self.retry_file):
            return None
        if time.time() - os.path.getmtime(self.retry_file) > 86400:
            self.logger.warning("Retry file stale (>24h) — removing")
            os.remove(self.retry_file)
            return None
        with open(self.retry_file) as f:
            return json.load(f)

    def delete(self):
        if os.path.exists(self.retry_file):
            os.remove(self.retry_file)
            self.logger.info("Cleared retry file")

    def schedule(self, retry_time: str):
        exe_path  = os.path.abspath(sys.argv[0])
        task_name = f"KRA_Retry_{retry_time.replace(':', '')}"
        os.system(f'schtasks /delete /tn "{task_name}" /f >nul 2>&1')
        os.system(
            f'schtasks /create /tn "{task_name}" '
            f'/tr "\\"{exe_path}\\" --retry" '
            f'/sc once /st {retry_time} /f /rl highest'
        )
        self.logger.info(f"Retry task scheduled at {retry_time}")


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )

    config        = ConfigLoader()
    sheets        = GoogleSheetsManager(config)
    logger        = SheetLogger(sheets, config)
    global GLOBAL_LOGGER
    GLOBAL_LOGGER = logger
    db            = DatabaseManager(config, logger)
    kra           = KRAChecker(config, logger)
    retry_manager = RetryManager(config, logger)

    is_retry    = "--retry" in sys.argv
    retry_hours = config.get("retry_hours", [0, 2, 4])

    logger.info("=" * 60)
    logger.info("KRA Auto-Checker")
    logger.info(f"Station : {config.get('station_name')}  ({config.get('anydesk_code')})")
    logger.info(f"Mode    : {'RETRY' if is_retry else 'INITIAL'}")

    # ── Determine which transaction to check ─────────────────────────
    report_entries   = []
    failed_transactions = []

    if is_retry:
        retry_data = retry_manager.load()
        if not retry_data:
            logger.error("No retry data found — aborting")
            return
        retry_data["retry_count"] += 1

        logger.info(f"Retry attempt #{retry_data['retry_count']}")

        transaction_list = [
            {
                "QRLink"      : tx["qr_link"],
                "TransDateTime": tx["transaction_date"],
                "CheckDate"   : tx["check_date"],
                "PaymentMode" : tx.get("payment_mode", "UNKNOWN"),
            }
            for tx in retry_data.get("transactions", [])
        ]

    else:
        check_date   = datetime.now()
        # # Change the date to test if data in db is old
        # check_date = datetime(2025, 8, 19)
        transactions = db.get_random_transactions(check_date)

        transaction_list = []

        if transactions["fuel_card"]:
            transaction_list.append(transactions["fuel_card"])
        else:
            report_entries.append({
                "check_date"      : check_date.strftime("%Y-%m-%d"),
                "status"          : "NO DATA",
                "transaction_date": "N/A",
                "invoice_number"  : "N/A",
                "qr_link"         : "N/A",
                "details"         : "No Fuel Card transaction found today",
            })

        if transactions["other"]:
            transaction_list.append(transactions["other"])
        else:
            report_entries.append({
                "check_date"      : check_date.strftime("%Y-%m-%d"),
                "status"          : "NO DATA",
                "transaction_date": "N/A",
                "invoice_number"  : "N/A",
                "qr_link"         : "N/A",
                "details"         : "No non-fuel transaction found today",
            })

    # ── Check all transactions against KRA first ──────────────────────
    for transaction in transaction_list:

        logger.info(
            f"Checking [{transaction.get('PaymentMode', 'UNKNOWN')}]: "
            f"{transaction['QRLink']}"
        )

        status, details, trans_date, invoice = kra.check_kra_link(
            transaction["QRLink"]
        )

        # TEMPORARY TEST FOR RETRY LOGIC, ENSURE check_date is given a date where transaction
        # Can be found in the db, e.g check_date = datetime(2025, 8, 19) 
        # then run  kra_auto_checker.py --retry
        # status = "ERROR"
        # details = "TEST RETRY FLOW"
        # TEMPORARY TEST

        logger.info(f"Result: {status} — {details}")

        entry = {
            "check_date"      : transaction["CheckDate"],
            "status"          : status,
            "transaction_date": trans_date or "N/A",
            "invoice_number"  : invoice    or "N/A",
            "qr_link"         : transaction["QRLink"],
            "details"         : f"[{transaction.get('PaymentMode')}] {details}",
        }

        if status == "SUCCESS":
            logger.success("Transaction confirmed with KRA")

        elif status in ("NOT_SUBMITTED", "ERROR"):

            logger.warning(f"{status}: {details}")
            transaction["LastStatus"] = status
            failed_transactions.append(transaction)

        report_entries.append(entry)

    # ── Write all results in one batch API call ───────────────────────
    if report_entries:
        sheets.add_report_entries(report_entries)


    # ── Handle retries ────────────────────────────────────────────────
    if failed_transactions:
        retry_count = (
            retry_data["retry_count"]
            if is_retry
            else 0
        )

        if retry_count < len(retry_hours):
            next_retry_time = f"{retry_hours[retry_count]:02d}:00"
            retry_manager.save(failed_transactions, retry_count)
            retry_manager.schedule(next_retry_time)
            logger.info(
                f"Next retry at {next_retry_time} "
                f"for {len(failed_transactions)} transaction(s)"
            )
        else:
            logger.warning("All retries exhausted")
            retry_manager.delete()
    else:
        retry_manager.delete()
    logger.info("=" * 60)
    logger.flush()


if __name__ == "__main__":
    try:
        main()

    except KeyboardInterrupt:
        print("\n⚠ Interrupted")

        if GLOBAL_LOGGER:
            GLOBAL_LOGGER.flush()

        sys.exit(0)

    except Exception as e:
        import traceback

        print(f"\n❌ Fatal error: {e}")
        traceback.print_exc()

        if GLOBAL_LOGGER:
            GLOBAL_LOGGER.flush()

        sys.exit(1)
