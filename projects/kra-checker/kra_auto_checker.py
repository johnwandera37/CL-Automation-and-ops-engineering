"""
KRA Auto-Checker  v2.1
Runs daily at 19:00 (via Windows Task Scheduler).
  • Pulls a random transaction from the ETIMS SQL database
  • Checks whether its QR link returns a valid KRA receipt
  • Writes result to the 'Report' sheet
  • Logs all steps to the 'Logs' sheet
  • On failure, saves the transaction and schedules overnight retries
    using times defined in the Global Config sheet (retry_hours)
"""

import pyodbc
import requests
import re
import json
import logging
logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.ERROR)
import time
import sys
import os
from datetime import datetime
from typing import Optional, Tuple, Dict

# ── Auto-update check (must be first) ────────────────────────────────────────
try:
    from auto_updater import check_and_update
    check_and_update()
except Exception as e:
    print(f"[KRA] Update check skipped: {e}")

# ── Base directory ────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

# ── Google Sheets ─────────────────────────────────────────────────────────────
try:
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
except ImportError:
    print("Google API libraries not installed. Run: pip install google-api-python-client google-auth")
    sys.exit(1)

from config_loader import ConfigLoader


# ─────────────────────────────────────────────────────────────────────────────
# Sheet Logger  (console + Google Sheets)
# ─────────────────────────────────────────────────────────────────────────────

class SheetLogger:
    EMOJI = {"INFO": "🟢", "WARNING": "🟡", "ERROR": "🔴", "SUCCESS": "🟢", "DEBUG": "⚪"}

    def __init__(self, sheets_manager, config: ConfigLoader):
        self.sheets = sheets_manager
        self.config = config
        self._log = logging.getLogger(__name__)

    def _write(self, level: str, message: str):
        emoji = self.EMOJI.get(level, "⚪")
        getattr(self._log, level.lower() if level != "SUCCESS" else "info")(
            f"{emoji} {message}"
        )
        try:
            self.sheets.add_log_entry(level, message)
        except Exception as e:
            self._log.error(f"Sheet log write failed: {e}")

    def info(self, m):    self._write("INFO", m)
    def warning(self, m): self._write("WARNING", m)
    def error(self, m):   self._write("ERROR", m)
    def success(self, m): self._write("SUCCESS", m)


# ─────────────────────────────────────────────────────────────────────────────
# KRA Checker
# ─────────────────────────────────────────────────────────────────────────────

class KRAChecker:

    def __init__(self, config: ConfigLoader, logger: SheetLogger):
        self.config = config
        self.logger = logger
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })

    def check_kra_link(self, qr_link: str) -> Tuple[str, str, Optional[str], Optional[str]]:
        """
        Check a QR link against the KRA portal.
        Returns (status, details, transaction_date, invoice_number).
        status is one of: SUCCESS | NOT_SUBMITTED | ERROR
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
                    self.logger.warning(
                        f"Network error ({attempt+1}/{max_retries}), retrying in {retry_delay}s…"
                    )
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
            return (
                "SUCCESS",
                "Transaction submitted to KRA",
                self._extract_date(html),
                self._extract_invoice(html),
            )
        if "could not be verified" in lower or "try again later" in lower:
            return ("NOT_SUBMITTED", "Invoice not verified by KRA", None, None)
        return ("ERROR", "Page loaded but status unclear", None, None)

    @staticmethod
    def _extract_date(html: str) -> Optional[str]:
        patterns = [
            r'<span[^>]*>(\d{1,2})/(\d{1,2})/(\d{4})\s+\d{2}:\d{2}:\d{2}</span>',
            r'(\d{1,2})/(\d{1,2})/(\d{4})\s+\d{2}:\d{2}:\d{2}',
        ]
        for p in patterns:
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

    def get_random_transaction(self, check_date: datetime) -> Optional[Dict]:
        """
        Fetch one random transaction for check_date.
        Progressively widens the time window if narrower ranges return nothing.
        """
        time_ranges = [
            ("16:00:00", "19:00:00", "4 PM – 7 PM"),
            ("12:00:00", "19:00:00", "12 PM – 7 PM"),
            ("08:00:00", "19:00:00", "8 AM – 7 PM"),
            ("00:00:00", "23:59:59", "any time today"),
        ]

        query = """
            SELECT TOP 1 TransDateTime, QRLink
            FROM ETPumpSales
            WHERE TransDateTime BETWEEN ? AND ?
              AND QRLink IS NOT NULL
              AND QRLink != ''
            ORDER BY NEWID()
        """

        try:
            conn = pyodbc.connect(self._conn_str())
            cursor = conn.cursor()

            for start_t, end_t, label in time_ranges:
                t_start = check_date.strftime(f"%Y-%m-%d {start_t}")
                t_end   = check_date.strftime(f"%Y-%m-%d {end_t}")
                cursor.execute(query, (t_start, t_end))
                row = cursor.fetchone()
                if row:
                    self.logger.info(f"Found transaction ({label}): {row[0]}")
                    cursor.close()
                    conn.close()
                    return {
                        "TransDateTime": row[0],
                        "QRLink"       : row[1],
                        "CheckDate"    : check_date.strftime("%Y-%m-%d"),
                    }
                self.logger.info(f"No transactions in {label}, widening search…")

            cursor.close()
            conn.close()
            self.logger.warning(f"No transactions found for {check_date.date()}")
            return None

        except Exception as e:
            self.logger.error(f"Database error: {e}")
            return None


# ─────────────────────────────────────────────────────────────────────────────
# Google Sheets Manager
# ─────────────────────────────────────────────────────────────────────────────

class GoogleSheetsManager:

    REPORT_HEADERS = [
        "Timestamp", "Station", "AnyDesk", "Check Date",
        "Status", "Trans Date", "Invoice", "QR Link", "Details",
    ]
    LOG_HEADERS = ["Timestamp", "Station", "AnyDesk", "Level", "Message"]

    def __init__(self, config: ConfigLoader):
        self.config = config
        self._log = logging.getLogger(__name__)
        self.service = self._authenticate()

    def _authenticate(self):
        sa_file = self.config.get("service_account_file", "credentials.json")
        creds = service_account.Credentials.from_service_account_file(
            sa_file, scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        return build("sheets", "v4", credentials=creds)

    # ── Public write methods ──────────────────────────────────────────

    def add_report_entry(self, data: Dict):
        sid = self.config.get("spreadsheet_id")
        self._ensure_sheet("Report", self.REPORT_HEADERS)
        self._heal_schema("Report", self.REPORT_HEADERS)

        raw_status = data.get("status", "")
        if "SUCCESS"       in raw_status: status_cell = "🟢 SUCCESS"
        elif "NOT_SUBMITTED" in raw_status: status_cell = "🔴 NOT SUBMITTED"
        elif "ERROR"       in raw_status: status_cell = "🟡 ERROR"
        elif "NO DATA"     in raw_status: status_cell = "⚪ NO DATA"
        else:                             status_cell = raw_status

        values = [[
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            self.config.get("station_name", ""),
            self.config.get("anydesk_code", ""),
            data.get("check_date", ""),
            status_cell,
            data.get("transaction_date", ""),
            data.get("invoice_number", ""),
            data.get("qr_link", ""),
            data.get("details", ""),
        ]]
        self.service.spreadsheets().values().append(
            spreadsheetId=sid,
            range="Report!A:I",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": values},
        ).execute()
        self._log.info("Report entry written")

    def add_log_entry(self, level: str, message: str):
        try:
            sid = self.config.get("spreadsheet_id")
            self._ensure_sheet("Logs", self.LOG_HEADERS)
            values = [[
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                self.config.get("station_name", ""),
                self.config.get("anydesk_code", ""),
                level,
                message,
            ]]
            self.service.spreadsheets().values().append(
                spreadsheetId=sid,
                range="Logs!A:E",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": values},
            ).execute()
        except Exception:
            pass  # never let logging crash the main flow

    # ── Internal helpers ──────────────────────────────────────────────

    def _ensure_sheet(self, name: str, headers: list):
        meta = self.service.spreadsheets().get(
            spreadsheetId=self.config.get("spreadsheet_id")
        ).execute()
        existing = [s["properties"]["title"] for s in meta.get("sheets", [])]
        if name in existing:
            return
        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.config.get("spreadsheet_id"),
            body={"requests": [{"addSheet": {"properties": {"title": name}}}]},
        ).execute()
        col = self._col(len(headers) - 1)
        self.service.spreadsheets().values().update(
            spreadsheetId=self.config.get("spreadsheet_id"),
            range=f"{name}!A1:{col}1",
            valueInputOption="RAW",
            body={"values": [headers]},
        ).execute()
        self._log.info(f"Created sheet: {name}")

    def _heal_schema(self, sheet_name: str, expected: list):
        result = self.service.spreadsheets().values().get(
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
        self.config = config
        self.logger = logger
        self.retry_file = config.get("retry_file", os.path.join(BASE_DIR, "retry_transaction.json"))

    def save(self, transaction: Dict, status: str):
        data = {
            "transaction_date": str(transaction["TransDateTime"]),
            "qr_link"         : transaction["QRLink"],
            "check_date"      : transaction["CheckDate"],
            "last_status"     : status,
            "retry_count"     : 0,
            "saved_at"        : datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        with open(self.retry_file, "w") as f:
            json.dump(data, f, indent=4)
        self.logger.info("Saved transaction for retry")

    def load(self) -> Optional[Dict]:
        if not os.path.exists(self.retry_file):
            return None
        age = time.time() - os.path.getmtime(self.retry_file)
        if age > 86400:
            self.logger.warning("Retry file is stale (>24h) — removing")
            os.remove(self.retry_file)
            return None
        with open(self.retry_file) as f:
            return json.load(f)

    def delete(self):
        if os.path.exists(self.retry_file):
            os.remove(self.retry_file)
            self.logger.info("Cleared retry file")

    def schedule(self, retry_time: str):
        """Register a one-time Windows Scheduled Task for the retry."""
        exe_path  = os.path.abspath(sys.argv[0])
        task_name = f"KRA_Retry_{retry_time.replace(':', '')}"
        os.system(f'schtasks /delete /tn "{task_name}" /f >nul 2>&1')
        cmd = (
            f'schtasks /create /tn "{task_name}" '
            f'/tr "\\"{exe_path}\\" --retry" '
            f'/sc once /st {retry_time} /f /rl highest'
        )
        result = os.system(cmd)
        if result == 0:
            self.logger.info(f"Retry task scheduled at {retry_time}")
        else:
            self.logger.error(f"Failed to schedule retry at {retry_time}")


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
    db            = DatabaseManager(config, logger)
    kra           = KRAChecker(config, logger)
    retry_manager = RetryManager(config, logger)

    is_retry      = "--retry" in sys.argv
    retry_hours   = config.get("retry_hours", [0, 2, 4])

    logger.info("=" * 60)
    logger.info("KRA Auto-Checker")
    logger.info(f"Station : {config.get('station_name')}  ({config.get('anydesk_code')})")
    logger.info(f"Mode    : {'RETRY' if is_retry else 'INITIAL'}")

    # ── Determine which transaction to check ─────────────────────────

    if is_retry:
        retry_data = retry_manager.load()
        if not retry_data:
            logger.error("No retry data found — aborting")
            return
        retry_data["retry_count"] += 1
        logger.info(f"Retry attempt #{retry_data['retry_count']}")
        transaction = {
            "QRLink"      : retry_data["qr_link"],
            "TransDateTime": retry_data["transaction_date"],
            "CheckDate"   : retry_data["check_date"],
        }
    else:
        check_date  = datetime.now()
        transaction = db.get_random_transaction(check_date)
        if not transaction:
            logger.warning("No transactions found for today")
            sheets.add_report_entry({
                "check_date"      : check_date.strftime("%Y-%m-%d"),
                "status"          : "NO DATA",
                "transaction_date": "N/A",
                "invoice_number"  : "N/A",
                "qr_link"         : "N/A",
                "details"         : "No transactions found in database",
            })
            return

    # ── Check the QR link ────────────────────────────────────────────

    logger.info(f"Checking: {transaction['QRLink']}")
    status, details, trans_date, invoice = kra.check_kra_link(transaction["QRLink"])
    logger.info(f"Result: {status} — {details}")

    # ── Handle result ────────────────────────────────────────────────

    if status == "SUCCESS":
        logger.success("Transaction confirmed with KRA")
        sheets.add_report_entry({
            "check_date"      : transaction["CheckDate"],
            "status"          : "SUCCESS",
            "transaction_date": trans_date or "N/A",
            "invoice_number"  : invoice   or "N/A",
            "qr_link"         : transaction["QRLink"],
            "details"         : details,
        })
        retry_manager.delete()

    elif status in ("NOT_SUBMITTED", "ERROR"):
        logger.warning(f"{status}: {details}")

        current_hour = datetime.now().hour

        # Work out the next retry slot
        next_retry_time = None
        if current_hour in retry_hours:
            idx = retry_hours.index(current_hour)
            if idx + 1 < len(retry_hours):
                next_retry_time = f"{retry_hours[idx + 1]:02d}:00"
        else:
            # First failure — schedule first overnight retry
            next_retry_time = f"{retry_hours[0]:02d}:00" if retry_hours else None

        if next_retry_time:
            retry_manager.save(transaction, status)
            retry_manager.schedule(next_retry_time)
            logger.info(f"Next retry at {next_retry_time}")
        else:
            # All retries exhausted — write final result
            logger.warning("All retries exhausted — writing final report")
            status_cell = "🔴 NOT SUBMITTED" if status == "NOT_SUBMITTED" else "🟡 ERROR"
            sheets.add_report_entry({
                "check_date"      : transaction["CheckDate"],
                "status"          : status_cell,
                "transaction_date": trans_date or "N/A",
                "invoice_number"  : invoice    or "N/A",
                "qr_link"         : transaction["QRLink"],
                "details"         : details,
            })
            retry_manager.delete()

    logger.info("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n⚠ Interrupted")
        sys.exit(0)
    except Exception as e:
        import traceback
        print(f"\n❌ Fatal error: {e}")
        traceback.print_exc()
        sys.exit(1)
