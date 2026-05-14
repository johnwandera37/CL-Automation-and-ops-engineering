"""
Station Heartbeat Monitor  v2.1
Runs on schedule via Windows Task Scheduler (called via VBS launcher).
  - Checks internet, SQL Server, disk space, local IP
  - Writes station status to Google Sheets (Station Status tab)
  - Manages auto-updates for BOTH heartbeat_monitor.exe and kra_checker.exe
  - Syncs Task Scheduler intervals from Global Config sheet
  - Logs errors to Heartbeat Error Logs tab
"""

import socket
import logging
import sys
import os
import platform
from datetime import datetime
from typing import Optional, Dict
from task_scheduler import (
    change_heartbeat_interval,
    change_kra_schedule_time,
)

# import task_scheduler
# print(task_scheduler.__file__)

logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.ERROR)

# ── Hide console window when running via Task Scheduler ──────────────────────
def _hide_console():
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)
    except Exception:
        pass

_hide_console()


# ─────────────────────────────────────────────────────────────────────────────
# Auto-updater — heartbeat manages updates for BOTH exes
# ─────────────────────────────────────────────────────────────────────────────

def _run_updates():
    """
    Runs in a background thread with 60s timeout so it never blocks heartbeat.
    Checks kra_checker.exe first (safe — not running), heartbeat last (causes restart).
    """
    import threading
    t = threading.Thread(target=_update_logic, daemon=True)
    t.start()
    t.join(timeout=60)
    if t.is_alive():
        logging.getLogger(__name__).warning("[UPDATER] Timed out after 60s — skipping")


def _update_logic():
    import json, io, subprocess, datetime as dt
    logger = logging.getLogger(__name__)
    try:
        from packaging import version as semver
        from config_loader import ConfigLoader
        from googleapiclient.discovery import build as gbuild
        from googleapiclient.http import MediaIoBaseDownload
        from google.oauth2 import service_account as gsa

        config     = ConfigLoader()
        base_dir   = os.path.dirname(os.path.abspath(sys.argv[0]))
        creds_file = os.path.join(base_dir, "credentials.json")

        creds_drive = gsa.Credentials.from_service_account_file(
            creds_file, scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
        drive_svc = gbuild("drive", "v3", credentials=creds_drive)

        creds_sheets = gsa.Credentials.from_service_account_file(
            creds_file, scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        sheets_svc = gbuild("sheets", "v4", credentials=creds_sheets)

        def _log_update(label, local_ver, remote_ver, success=True):
            try:
                sheets_svc.spreadsheets().values().append(
                    spreadsheetId=config.get("spreadsheet_id"),
                    range="Logs!A:E",
                    valueInputOption="RAW",
                    insertDataOption="INSERT_ROWS",
                    body={"values": [[
                        dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        config.get("station_name", ""),
                        config.get("anydesk_code", ""),
                        "🔄 UPDATE",
                        f"Auto-updated {label} from v{local_ver} to v{remote_ver}"
                    ]]}
                ).execute()
            except Exception:
                pass
            try:
                helper_id  = config.get("automation_helper_sheet_id")
                anydesk    = str(config.get("anydesk_code", "")).strip()
                timestamp  = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
                status_val = f"✅ v{remote_ver}" if success else "🔴 Failed"
                rows = sheets_svc.spreadsheets().values().get(
                    spreadsheetId=helper_id, range="Station Mapping!A:B"
                ).execute().get("values", [])
                row_num = None
                for i, row in enumerate(rows):
                    if len(row) >= 2 and str(row[1]).strip() == anydesk:
                        row_num = i + 1
                        break
                if row_num:
                    col_range = f"Station Mapping!D{row_num}:E{row_num}" if label == "kra_checker" \
                                else f"Station Mapping!F{row_num}:G{row_num}"
                    sheets_svc.spreadsheets().values().update(
                        spreadsheetId=helper_id,
                        range=col_range,
                        valueInputOption="RAW",
                        body={"values": [[status_val, timestamp]]}
                    ).execute()
            except Exception:
                pass

        def _download_exe(drive_id, dest_path):
            request    = drive_svc.files().get_media(fileId=drive_id)
            fh         = io.FileIO(dest_path, "wb")
            downloader = MediaIoBaseDownload(fh, request, chunksize=65536 * 16)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                if status:
                    logger.info(f"[UPDATER] {int(status.progress() * 100)}%")
            fh.close()
            return os.path.getsize(dest_path)

        def _write_version(key, value):
            try:
                p = os.path.join(base_dir, "config.json")
                with open(p) as f:
                    cfg = json.load(f)
                cfg[key] = value
                with open(p, "w") as f:
                    json.dump(cfg, f, indent=2)
            except Exception as e:
                logger.warning(f"[UPDATER] Could not write version: {e}")

        def _swap_bat(exe_path, new_exe, backup_exe, restart=False):
            name     = "apply_update.bat" if restart else "apply_kra_update.bat"
            bat_path = os.path.join(base_dir, name)
            lines    = [
                "@echo off", "timeout /t 2 >nul",
                f'if exist "{backup_exe}" del /f /q "{backup_exe}"',
                f'if exist "{exe_path}" ren "{exe_path}" "{os.path.basename(backup_exe)}"',
                f'ren "{new_exe}" "{os.path.basename(exe_path)}"',
            ]
            if restart:
                lines.append(f'start "" "{exe_path}"')
            lines += ['del /f /q "%~f0"', "exit"]
            with open(bat_path, "w") as f:
                f.write("\r\n".join(lines) + "\r\n")
            return bat_path

        exes = [
            {"label": "kra_checker",       "exe": "kra_checker.exe",
             "lkey": "current_version_kra",        "rkey": "remote_version_kra",
             "dkey": "kra_checker_drive_id",        "self": False},
            {"label": "heartbeat_monitor",  "exe": "heartbeat_monitor.exe",
             "lkey": "current_version_heartbeat",   "rkey": "remote_version_heartbeat",
             "dkey": "heartbeat_monitor_drive_id",   "self": True},
        ]

        hb_restart = False
        hb_exe_path = hb_new = hb_backup = None

        for exe in exes:
            local_ver  = config.get(exe["lkey"], "0.0.0")
            remote_ver = config.get(exe["rkey"], "0.0.0")
            drive_id   = config.get(exe["dkey"])
            logger.info(f"[UPDATER] {exe['label']}  local={local_ver}  remote={remote_ver}")

            if not drive_id or semver.parse(remote_ver) <= semver.parse(local_ver):
                logger.info(f"[UPDATER] {exe['label']} up to date")
                continue

            logger.info(f"[UPDATER] Updating {exe['label']} ({local_ver} -> {remote_ver})")
            exe_path   = os.path.join(base_dir, exe["exe"])
            new_exe    = exe_path + ".new"
            backup_exe = exe_path + ".old"

            size = _download_exe(drive_id, new_exe)
            size_mb = size / (1024 * 1024)
            logger.info(f"[UPDATER] Downloaded {size_mb:.1f} MB")

            if size < 10 * 1024 * 1024:
                logger.error(f"[UPDATER] Too small ({size_mb:.1f} MB) — aborting")
                if os.path.exists(new_exe):
                    os.remove(new_exe)
                _log_update(exe["label"], local_ver, remote_ver, success=False)
                continue

            _write_version(exe["lkey"], remote_ver)
            _log_update(exe["label"], local_ver, remote_ver, success=True)

            if exe["self"]:
                hb_restart  = True
                hb_exe_path = exe_path
                hb_new      = new_exe
                hb_backup   = backup_exe
            else:
                bat = _swap_bat(exe_path, new_exe, backup_exe, restart=False)
                subprocess.Popen(["cmd", "/c", bat], creationflags=subprocess.CREATE_NO_WINDOW)
                logger.info("[UPDATER] kra_checker.exe swap launched")

        if hb_restart and hb_exe_path:
            bat = _swap_bat(hb_exe_path, hb_new, hb_backup, restart=True)
            subprocess.Popen(["cmd", "/c", bat], creationflags=subprocess.CREATE_NO_WINDOW)
            logger.info("[UPDATER] Heartbeat hot-swap launched — restarting")
            sys.exit(0)

    except Exception as e:
        logging.getLogger(__name__).warning(f"[UPDATER] Skipped: {e}")


_run_updates()


# ── Resolve base dir AFTER potential restart ──────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

try:
    import pyodbc
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
except ImportError as e:
    print(f"Missing library: {e}")
    sys.exit(1)

from config_loader import ConfigLoader


# ─────────────────────────────────────────────────────────────────────────────
# System Monitor
# ─────────────────────────────────────────────────────────────────────────────

class SystemMonitor:

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
        try:
            if platform.system() == "Windows":
                import ctypes
                free_bytes = ctypes.c_ulonglong(0)
                ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                    ctypes.c_wchar_p("C:\\"), None, None, ctypes.pointer(free_bytes)
                )
                return f"{free_bytes.value / (1024 ** 3):.1f} GB"
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

    COLUMN_WIDTHS = {
        "Station Status"      : [180, 130, 160, 110, 140, 110, 80, 80, 160],
        "Heartbeat Error Logs": [160, 180, 130, 100, 400],
    }

    def __init__(self, config: ConfigLoader):
        self.config  = config
        self.logger  = logging.getLogger(__name__)
        self.service = self._authenticate()

    def _authenticate(self):
        sa_file = self.config.get("service_account_file", "credentials.json")
        creds   = service_account.Credentials.from_service_account_file(
            sa_file, scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        return build("sheets", "v4", credentials=creds)

    # ── Public ────────────────────────────────────────────────────────

    def update_station_status(self, status_data: Dict):
        sid = self.config.get("spreadsheet_id")
        self._ensure_sheet_exists("Station Status", self.STATION_STATUS_HEADERS)
        self._heal_schema("Station Status", self.STATION_STATUS_HEADERS)

        station_row = self._find_station_row()
        values = [[
            self.config.get("station_name", ""),
            self.config.get("anydesk_code", ""),
            status_data.get("last_seen", ""),
            "",   # formula-controlled
            status_data.get("ip_address", ""),
            status_data.get("disk_space", ""),
            status_data.get("sql_status", ""),
            status_data.get("cpu_temp") or "",
            status_data.get("heartbeat_interval", 30),
        ]]

        if station_row:
            self.service.spreadsheets().values().update(
                spreadsheetId=sid,
                range=f"Station Status!A{station_row}:I{station_row}",
                valueInputOption="USER_ENTERED",
                body={"values": values},
            ).execute()
            self._write_status_formula(station_row)
            self._auto_resize("Station Status")
            self.logger.info(f"Updated station status (row {station_row})")
        else:
            self.service.spreadsheets().values().append(
                spreadsheetId=sid,
                range="Station Status!A:I",
                valueInputOption="USER_ENTERED",
                insertDataOption="INSERT_ROWS",
                body={"values": values},
            ).execute()
            new_row = self._find_station_row()
            if new_row:
                self._write_status_formula(new_row)
            self._auto_resize("Station Status")
            self.logger.info("Added new station row")

    def log_error(self, level: str, message: str):
        try:
            sid = self.config.get("spreadsheet_id")
            self._ensure_sheet_exists("Heartbeat Error Logs", self.HEARTBEAT_ERROR_HEADERS)
            self._heal_schema("Heartbeat Error Logs", self.HEARTBEAT_ERROR_HEADERS)
            today = datetime.now().strftime("%d/%m/%Y")
            self._insert_date_separator("Heartbeat Error Logs", today, len(self.HEARTBEAT_ERROR_HEADERS))
            self.service.spreadsheets().values().append(
                spreadsheetId=sid,
                range="Heartbeat Error Logs!A:E",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": [[
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    self.config.get("station_name", ""),
                    self.config.get("anydesk_code", ""),
                    level,
                    message,
                ]]},
            ).execute()
        except Exception:
            pass

    # ── Internal ──────────────────────────────────────────────────────

    def _find_station_row(self) -> Optional[int]:
        try:
            rows    = self.service.spreadsheets().values().get(
                spreadsheetId=self.config.get("spreadsheet_id"),
                range="Station Status!A:B",
            ).execute().get("values", [])
            anydesk = str(self.config.get("anydesk_code", "")).strip()
            station = str(self.config.get("station_name", "")).strip()
            for i, row in enumerate(rows):
                if len(row) >= 2 and str(row[1]).strip() == anydesk:
                    return i + 1
            for i, row in enumerate(rows):
                if len(row) >= 1 and str(row[0]).strip() == station:
                    return i + 1
        except Exception:
            pass
        return None

    def _write_status_formula(self, row: int):
        """
        Online/Stale/Offline formula for column D.
        Last Seen must be a REAL datetime value.
        Column I contains heartbeat interval in minutes.
        Google Apps Script refreshes NOW() every minute.
        """

        formula = (
            f'=IF('
            f'C{row}="",'
            f'"⚪ Never",'
            f'IF('
            f'(NOW()-C{row})*1440>I{row}*2,'
            f'"🔴 Offline",'
            f'IF('
            f'(NOW()-C{row})*1440>I{row},'
            f'"🟡 Stale",'
            f'"🟢 Online"'
            f')'
            f')'
            f')'
        )

        try:
            self.service.spreadsheets().values().update(
                spreadsheetId=self.config.get("spreadsheet_id"),
                range=f"Station Status!D{row}",
                valueInputOption="USER_ENTERED",
                body={"values": [[formula]]},
            ).execute()

        except Exception as e:
            self.logger.warning(f"Could not write status formula: {e}")

    def _ensure_sheet_exists(self, name: str, headers: list):
        meta     = self.service.spreadsheets().get(
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

    def _heal_schema(self, sheet_name: str, expected: list):
        result   = self.service.spreadsheets().values().get(
            spreadsheetId=self.config.get("spreadsheet_id"),
            range=f"{sheet_name}!A1:Z1",
        ).execute()
        existing = result.get("values", [[]])[0]
        if existing != expected:
            self.service.spreadsheets().values().update(
                spreadsheetId=self.config.get("spreadsheet_id"),
                range=f"{sheet_name}!A1",
                valueInputOption="RAW",
                body={"values": [expected]},
            ).execute()
            self.logger.warning(f"Auto-healed schema for {sheet_name}")

    def _get_sheet_id(self, sheet_name: str) -> int:
        meta = self.service.spreadsheets().get(
            spreadsheetId=self.config.get("spreadsheet_id")
        ).execute()
        for s in meta.get("sheets", []):
            if s["properties"]["title"] == sheet_name:
                return s["properties"]["sheetId"]
        raise ValueError(f"Sheet not found: {sheet_name}")

    def _insert_date_separator(self, sheet_name: str, date_str: str, num_cols: int):
        """
        Appends a plain merged date row when the date changes.
        Flag file prevents duplicates — only the first run each day inserts it.
        No formatting — plain text only.
        """
        import pathlib
        flag_file = os.path.join(BASE_DIR, f".lastdate_{sheet_name.replace(' ', '_')}")
        try:
            if os.path.exists(flag_file) and pathlib.Path(flag_file).read_text().strip() == date_str:
                return False
        except Exception:
            pass
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
            pathlib.Path(flag_file).write_text(date_str)
            return True
        except Exception as e:
            self.logger.warning(f"[DATE SEP] {sheet_name}: {e}")
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
            self.logger.warning(f"Column resize failed for {sheet_name}: {e}")

    @staticmethod
    def _col_letter(idx: int) -> str:
        result = ""
        while idx >= 0:
            result = chr(idx % 26 + 65) + result
            idx = idx // 26 - 1
        return result


# ─────────────────────────────────────────────────────────────────────────────
# Task Scheduler Sync
# ─────────────────────────────────────────────────────────────────────────────

def sync_scheduled_tasks(config: ConfigLoader, log):
    # import json here, if import is from task_scheduler then json is not loaded at all
    # investigate WHY module-level import disappeared.
    import subprocess, json

    heartbeat_interval = int(config.get("heartbeat_interval", 30))
    kra_check_time = config.get("kra_check_time", "19:00")

    config_path = os.path.join(
        os.path.dirname(os.path.abspath(sys.argv[0])),
        "config.json"
    )
    
    try:
        with open(config_path) as f:
            local = json.load(f)
    except Exception as e:
        log.warning(f"[TASKS] Failed loading config cache: {e}")
        local = {}

    config_updates = {}

    cached_interval = int(local.get("heartbeat_interval", 0))
    if cached_interval != heartbeat_interval:
        log.info(
            f"[TASKS] Heartbeat interval changed: "
            f"{cached_interval} -> {heartbeat_interval} min"
        )

        if change_heartbeat_interval(heartbeat_interval, log):
            config_updates["heartbeat_interval"] = heartbeat_interval

    else:
        log.info(f"[TASKS] Heartbeat interval unchanged ({heartbeat_interval} min)")

    cached_time = local.get("kra_check_time", "")
    if cached_time != kra_check_time:

        log.info(
            f"[TASKS] KRA check time changed: "
            f"{cached_time} -> {kra_check_time}"
        )

        if change_kra_schedule_time(kra_check_time, log):
            config_updates["kra_check_time"] = kra_check_time

    else:
        log.info(f"[TASKS] KRA check time unchanged ({kra_check_time})")

    # update config.json cache
    for key in ["heartbeat_interval", "kra_check_time", "max_retries",
                "retry_delay", "timeout", "retry_hours", "spreadsheet_id"]:
        val = config.get(key)
        if val is not None and local.get(key) != val:
            config_updates[key] = val

    if config_updates:
        try:
            with open(config_path) as f:
                cfg = json.load(f)
            cfg.update(config_updates)
            with open(config_path, "w") as f:
                json.dump(cfg, f, indent=2)
            log.info(f"[TASKS] config.json updated: {list(config_updates.keys())}")
        except Exception as e:
            log.warning(f"[TASKS] Could not update config.json: {e}")


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

    config             = ConfigLoader()
    station            = config.get("station_name", "Unknown")
    anydesk            = config.get("anydesk_code", "N/A")
    heartbeat_interval = int(config.get("heartbeat_interval", 30))

    log.info("=" * 60)
    log.info("Heartbeat Monitor")
    log.info(f"Station : {station}  ({anydesk})")

    sync_scheduled_tasks(config, log)

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
        "last_seen"         : datetime.now().isoformat(sep=" "),
        "ip_address"        : local_ip,
        "disk_space"        : disk_space,
        "sql_status"        : "🟢 OK" if sql_ok else "🔴 Down",
        "cpu_temp"          : None,
        "heartbeat_interval": heartbeat_interval,
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
