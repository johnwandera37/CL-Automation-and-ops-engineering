"""
Station Heartbeat Monitor  v2.1
Runs every 30 minutes (via Windows Task Scheduler).
  • Checks internet, SQL Server, disk space, local IP
  • Writes a timestamped row to the 'Station Status' sheet
  • Status column is a formula (Online/Stale/Offline) — never written by Python
  • Logs errors to 'Heartbeat Error Logs' sheet
"""

import socket
import logging
logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.ERROR)
import sys
import os
import platform
from datetime import datetime
from typing import Optional, Dict

# ── Auto-update check (must be first) ────────────────────────────────────────
try:
    from auto_updater import check_and_update
    check_and_update()
except Exception as e:
    print(f"[Heartbeat] Update check skipped: {e}")

# ── Resolve base dir AFTER potential restart ──────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

# ── Third-party imports ───────────────────────────────────────────────────────
try:
    import pyodbc
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
except ImportError as e:
    print(f"Missing library: {e}")
    print("Run: pip install pyodbc google-api-python-client google-auth")
    sys.exit(1)

from config_loader import ConfigLoader


# ─────────────────────────────────────────────────────────────────────────────
# System Monitor
# ─────────────────────────────────────────────────────────────────────────────

class SystemMonitor:
    """Collect local system health metrics."""

    @staticmethod
    def check_internet() -> bool:
        try:
            socket.create_connection(("8.8.8.8", 53), timeout=5)
            return True
        except OSError:
            return False

    @staticmethod
    def get_local_ip() -> str:
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except Exception:
            return "N/A"

    @staticmethod
    def get_disk_space() -> str:
        """Return free space on C:\\ as a human-readable string."""
        try:
            if platform.system() == "Windows":
                import ctypes
                free_bytes = ctypes.c_ulonglong(0)
                ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                    ctypes.c_wchar_p("C:\\"),
                    None, None,
                    ctypes.pointer(free_bytes),
                )
                free_gb = free_bytes.value / (1024 ** 3)
                return f"{free_gb:.1f} GB"
        except Exception:
            pass
        return "N/A"

    @staticmethod
    def check_sql_server(config: ConfigLoader) -> bool:
        try:
            conn_str = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={config.get('sql_server', '.\\SQLEXPRESS')};"
                f"DATABASE={config.get('sql_database', 'ETIMS')};"
                f"UID={config.get('sql_username', 'sa')};"
                f"PWD={config.get('sql_password', '')}"
            )
            conn = pyodbc.connect(conn_str, timeout=5)
            conn.close()
            return True
        except Exception:
            return False


# ─────────────────────────────────────────────────────────────────────────────
# Google Sheets Manager
# ─────────────────────────────────────────────────────────────────────────────

class GoogleSheetsManager:

    STATION_STATUS_HEADERS = [
        "Station", "AnyDesk", "Last Seen", "Status",
        "IP Address", "Disk Space", "SQL Server", "CPU Temp",
        "Heartbeat Interval (min)",
    ]

    HEARTBEAT_ERROR_HEADERS = [
        "Timestamp", "Station", "AnyDesk", "Level", "Message"
    ]

    def __init__(self, config: ConfigLoader):
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.service = self._authenticate()

    def _authenticate(self):
        sa_file = self.config.get("service_account_file", "credentials.json")
        creds = service_account.Credentials.from_service_account_file(
            sa_file,
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
        return build("sheets", "v4", credentials=creds)

    # ── Public methods ────────────────────────────────────────────────

    def update_station_status(self, status_data: Dict):
        """Upsert this station's row in the 'Station Status' sheet."""
        spreadsheet_id = self.config.get("spreadsheet_id")

        self._ensure_sheet_exists("Station Status", self.STATION_STATUS_HEADERS)
        self._heal_schema("Station Status", self.STATION_STATUS_HEADERS)

        station_row = self._find_station_row()

        values = [[
            self.config.get("station_name", ""),
            self.config.get("anydesk_code", ""),
            status_data.get("last_seen", ""),
            "",                                       # formula-controlled — leave blank
            status_data.get("ip_address", ""),
            status_data.get("disk_space", ""),
            status_data.get("sql_status", ""),
            status_data.get("cpu_temp") or "",        # omit "N/A"
            status_data.get("heartbeat_interval", 30),
        ]]

        if station_row:
            self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"Station Status!A{station_row}:I{station_row}",
                valueInputOption="USER_ENTERED",
                body={"values": values},
            ).execute()
            self._write_status_formula(station_row)
            self.logger.info(f"Updated station status (row {station_row})")
        else:
            self.service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range="Station Status!A:I",
                valueInputOption="USER_ENTERED",
                insertDataOption="INSERT_ROWS",
                body={"values": values},
            ).execute()
            # Find the newly appended row and write the formula
            new_row = self._find_station_row()
            if new_row:
                self._write_status_formula(new_row)
            self.logger.info("Added new station row")

    def log_error(self, level: str, message: str):
        """Append a row to 'Heartbeat Error Logs' — never raises."""
        try:
            spreadsheet_id = self.config.get("spreadsheet_id")
            self._ensure_sheet_exists("Heartbeat Error Logs", self.HEARTBEAT_ERROR_HEADERS)
            self._heal_schema("Heartbeat Error Logs", self.HEARTBEAT_ERROR_HEADERS)

            values = [[
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                self.config.get("station_name", ""),
                self.config.get("anydesk_code", ""),
                level,
                message,
            ]]
            self.service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range="Heartbeat Error Logs!A:E",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": values},
            ).execute()
        except Exception:
            pass  # never let logging crash the heartbeat

    # ── Internal helpers ──────────────────────────────────────────────

    def _write_status_formula(self, row: int):
        """
        Write the Online/Stale/Offline formula into column D.
        Uses column I (heartbeat interval) so the threshold is self-adjusting.
        """
        formula = (
            f'=IF(C{row}="","⚪ Never",'
            f'IF((NOW()-DATEVALUE(LEFT(C{row},10))-TIMEVALUE(RIGHT(C{row},8)))*1440>I{row}*2,'
            f'"🔴 Offline",'
            f'IF((NOW()-DATEVALUE(LEFT(C{row},10))-TIMEVALUE(RIGHT(C{row},8)))*1440>I{row},'
            f'"🟡 Stale","🟢 Online")))'
        )
        self.service.spreadsheets().values().update(
            spreadsheetId=self.config.get("spreadsheet_id"),
            range=f"Station Status!D{row}",
            valueInputOption="USER_ENTERED",
            body={"values": [[formula]]},
        ).execute()

    def _find_station_row(self) -> Optional[int]:
        """Return 1-based row index for this station, or None."""
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.config.get("spreadsheet_id"),
                range="Station Status!A:B",
            ).execute()
            rows = result.get("values", [])
            anydesk = str(self.config.get("anydesk_code", "")).strip()
            station = str(self.config.get("station_name", "")).strip()

            # Priority 1: match by AnyDesk code
            for i, row in enumerate(rows):
                if len(row) >= 2 and str(row[1]).strip() == anydesk:
                    return i + 1
            # Priority 2: match by station name
            for i, row in enumerate(rows):
                if len(row) >= 1 and str(row[0]).strip() == station:
                    return i + 1
        except Exception:
            pass
        return None

    def _ensure_sheet_exists(self, name: str, headers: list):
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

        col = self._col_letter(len(headers) - 1)
        self.service.spreadsheets().values().update(
            spreadsheetId=self.config.get("spreadsheet_id"),
            range=f"{name}!A1:{col}1",
            valueInputOption="RAW",
            body={"values": [headers]},
        ).execute()
        self.logger.info(f"Created sheet: {name}")

    def _heal_schema(self, sheet_name: str, expected_headers: list):
        """Overwrite header row if it drifted."""
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.config.get("spreadsheet_id"),
            range=f"{sheet_name}!A1:Z1",
        ).execute()
        existing = result.get("values", [[]])[0]
        if existing != expected_headers:
            self.service.spreadsheets().values().update(
                spreadsheetId=self.config.get("spreadsheet_id"),
                range=f"{sheet_name}!A1",
                valueInputOption="RAW",
                body={"values": [expected_headers]},
            ).execute()
            self.logger.warning(f"Auto-healed schema for {sheet_name}")

    @staticmethod
    def _col_letter(idx: int) -> str:
        result = ""
        while idx >= 0:
            result = chr(idx % 26 + 65) + result
            idx = idx // 26 - 1
        return result


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )
    log = logging.getLogger(__name__)

    config = ConfigLoader()
    station = config.get("station_name", "Unknown")
    anydesk = config.get("anydesk_code", "N/A")
    heartbeat_interval = int(config.get("heartbeat_interval", 30))

    log.info("=" * 60)
    log.info("Heartbeat Monitor")
    log.info(f"Station : {station}  ({anydesk})")

    monitor = SystemMonitor()
    sheets  = GoogleSheetsManager(config)

    internet_ok = monitor.check_internet()
    local_ip    = monitor.get_local_ip()
    disk_space  = monitor.get_disk_space()
    sql_ok      = monitor.check_sql_server(config)

    log.info(f"Internet : {'🟢 OK' if internet_ok else '🔴 Down'}")
    log.info(f"Local IP : {local_ip}")
    log.info(f"Disk     : {disk_space}")
    log.info(f"SQL      : {'🟢 OK' if sql_ok else '🔴 Down'}")

    status_data = {
        "last_seen"          : datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ip_address"         : local_ip,
        "disk_space"         : disk_space,
        "sql_status"         : "🟢 OK" if sql_ok else "🔴 Down",
        "cpu_temp"           : None,          # placeholder for future sensor
        "heartbeat_interval" : heartbeat_interval,
    }

    try:
        sheets.update_station_status(status_data)
        log.info("Station Status updated")
    except Exception as e:
        log.error(f"Failed to update Station Status: {e}")
        try:
            sheets.log_error("ERROR", f"Heartbeat update failed: {e}")
        except Exception:
            pass

    log.info("Heartbeat complete")
    log.info("=" * 60)


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
