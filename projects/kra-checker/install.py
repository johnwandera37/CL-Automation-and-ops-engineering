"""
KRA Auto-Checker Installer v2.4
- Auto-detects AnyDesk ID
- Looks up station name from Automation Helper sheet
- If station not found, prompts for name and writes it back to the sheet
- Self-elevates to Administrator (double-click friendly)
- No Python required on target stations (build with PyInstaller)
"""

# import getpass
import os
import sys
import json
import shutil
import subprocess
import ctypes
import logging
import task_scheduler

# Suppress google file_cache warning globally
logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.ERROR)

BASE_OPS_DIR = r"C:\Automation_and_ops_engineering"
INSTALL_DIR  = os.path.join(BASE_OPS_DIR, "KRA_Checker")
REQUIRED_FILES = [
    "kra_checker.exe",
    "heartbeat_monitor.exe",
    "uninstall.exe",
    "credentials.json",
    "config.json",
    "anydesk_detector.py",
    "fetch_station_info.py", 
]


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def self_elevate():
    """Re-launch this script as Administrator if not already elevated."""
    if is_admin():
        return
    print("Requesting Administrator privileges...")
    # Re-run as admin — works for both .py and .exe
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1
    )
    sys.exit(0)


def banner(msg):
    print(f"\n{'='*52}")
    print(f"  {msg}")
    print("=" * 52)


def step(n, total, msg):
    print(f"\n[{n}/{total}] {msg}...")


def main():
    # ── Auto-elevate to admin ─────────────────────────────────────────
    self_elevate()

    banner("johnwandera org")
    banner("KRA Auto-Checker Installation v2.4")

    script_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))

    # ── Step 1: Verify files ──────────────────────────────────────────
    step(1, 7, "Verifying files")
    for f in REQUIRED_FILES:
        path = os.path.join(script_dir, f)
        if not os.path.exists(path):
            print(f"  ERROR: {f} not found in installer directory!")
            input("\nPress Enter to exit...")
            sys.exit(1)
        print(f"  Found: {f}")

    # ── Step 2: Copy files ────────────────────────────────────────────
    step(2, 7, f"Copying files to {INSTALL_DIR}")
    os.makedirs(INSTALL_DIR, exist_ok=True)
    for f in REQUIRED_FILES:
        shutil.copy2(os.path.join(script_dir, f), os.path.join(INSTALL_DIR, f))
        print(f"  Copied: {f}")

    os.chdir(INSTALL_DIR)

    # ── Step 3: Detect AnyDesk ID ─────────────────────────────────────
    step(3, 7, "Detecting AnyDesk ID")

    # Import directly instead of subprocess so it works inside the exe
    sys.path.insert(0, INSTALL_DIR)
    detected_anydesk = None

    try:
        from anydesk_detector import get_anydesk_id
        detected_anydesk = get_anydesk_id()
    except Exception as e:
        print(f"  Detection error: {e}")

    if not detected_anydesk:
        print("\n  WARNING: Could not auto-detect AnyDesk ID.")
        print("  Make sure AnyDesk is installed and has been opened at least once.")
        detected_anydesk = input("\n  Enter AnyDesk code manually (or leave blank to skip): ").strip()
        if not detected_anydesk:
            detected_anydesk = "UNKNOWN"
            print("  Skipping — set anydesk_code in config.json manually later.")
    else:
        print(f"  AnyDesk ID: {detected_anydesk}")

    # ── Step 4: Look up station name from sheet ───────────────────────
    step(4, 7, "Looking up station in Automation Helper sheet")
    station_name = None
    helper_sheet_id = None
    creds_file = os.path.join(INSTALL_DIR, "credentials.json")

    try:
        with open(os.path.join(INSTALL_DIR, "config.json")) as f:
            config = json.load(f)
        helper_sheet_id = config.get("automation_helper_sheet_id", "")

        if not helper_sheet_id:
            print("  WARNING: automation_helper_sheet_id not in config.json")
        elif detected_anydesk == "UNKNOWN":
            print("  Skipping sheet lookup — AnyDesk code unknown")
        else:
            from fetch_station_info import fetch_station_by_anydesk
            station_name = fetch_station_by_anydesk(
                detected_anydesk, creds_file, helper_sheet_id
            )
            if station_name:
                print(f"  Station identified: {station_name}")
            else:
                print("  Station not found in sheet.")

    except Exception as e:
        print(f"  WARNING: Sheet lookup failed: {e}")

    # ── If not found — ask and write back to sheet ────────────────────
    if not station_name:
        print("\n  This station is not in the Automation Helper sheet yet.")
        station_name = input("  Enter station name to register: ").strip()
        if not station_name:
            print("  ERROR: Station name cannot be empty.")
            input("\nPress Enter to exit...")
            sys.exit(1)

        if helper_sheet_id and detected_anydesk != "UNKNOWN":
            print(f"\n  Adding '{station_name}' to the Station Mapping sheet...")
            try:
                from fetch_station_info import add_station_to_sheet
                success = add_station_to_sheet(
                    station_name, detected_anydesk, creds_file, helper_sheet_id
                )
                if success:
                    print(f"  Sheet updated — other tools will now recognise this station.")
                else:
                    print(f"  Could not update sheet automatically. Please add it manually.")
            except Exception as e:
                print(f"  Sheet write failed: {e} — please add manually.")
        else:
            print("  Skipping sheet update (missing sheet ID or AnyDesk code).")

    # ── Step 5: Confirm details ───────────────────────────────────────
    step(5, 7, "Final configuration")
    print(f"\n  Station name : {station_name}")
    print(f"  AnyDesk code : {detected_anydesk}")
    sql_password = input("\n  SQL Server password for 'sa' account: ").strip()

    # ── Step 6: Write config.json ─────────────────────────────────────
    step(6, 7, "Writing configuration")
    try:
        config_path = os.path.join(INSTALL_DIR, "config.json")
        with open(config_path) as f:
            config = json.load(f)

        config["station_name"] = station_name
        config["anydesk_code"] = detected_anydesk
        config["sql_password"] = sql_password

        with open(config_path, "w") as f:
            json.dump(config, f, indent=2)

        print("  config.json updated successfully")
    except Exception as e:
        print(f"  ERROR writing config.json: {e}")
        input("\nPress Enter to exit...")
        sys.exit(1)

    # Clean up temp files
    for tmp in ["detected_anydesk.txt", "detected_station.txt"]:
        path = os.path.join(INSTALL_DIR, tmp)
        if os.path.exists(path):
            os.remove(path)

    # ── Mark station as installed in Automation Helper sheet ─────────────
    print("  Updating Station Mapping sheet...")
    try:
        from googleapiclient.discovery import build as gbuild
        from google.oauth2 import service_account as gsa
        from datetime import datetime

        with open(os.path.join(INSTALL_DIR, "config.json")) as f:
            _cfg = json.load(f)

        _creds = gsa.Credentials.from_service_account_file(
            os.path.join(INSTALL_DIR, "credentials.json"),
            scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        _svc = gbuild("sheets", "v4", credentials=_creds)
        _helper_id = _cfg.get("automation_helper_sheet_id")
        _kra_ver   = _cfg.get("current_version_kra", "1.0.0")
        _hb_ver    = _cfg.get("current_version_heartbeat", "1.0.0")

        # Find the station row by AnyDesk code
        _result = _svc.spreadsheets().values().get(
            spreadsheetId=_helper_id,
            range="Station Mapping!A:B"
        ).execute()
        _rows = _result.get("values", [])
        _row_num = None
        for _i, _row in enumerate(_rows):
            if len(_row) >= 2 and str(_row[1]).strip() == str(detected_anydesk).strip():
                _row_num = _i + 1
                break

        if _row_num:
            # Ensure headers exist in C:G
            _headers = _svc.spreadsheets().values().get(
                spreadsheetId=_helper_id,
                range="Station Mapping!C1:G1"
            ).execute().get("values", [[]])
            if not _headers or _headers[0] != ["Installed", "KRA Checker", "KRA Updated", "Heartbeat Monitor", "HB Updated"]:
                _svc.spreadsheets().values().update(
                    spreadsheetId=_helper_id,
                    range="Station Mapping!C1:G1",
                    valueInputOption="RAW",
                    body={"values": [["Installed", "KRA Checker", "KRA Updated", "Heartbeat Monitor", "HB Updated"]]}
                ).execute()

            # Write install record
            _install_date = datetime.now().strftime("%Y-%m-%d %H:%M")
            _svc.spreadsheets().values().update(
                spreadsheetId=_helper_id,
                range=f"Station Mapping!C{_row_num}:G{_row_num}",
                valueInputOption="RAW",
                body={"values": [[
                    f"✅ {_install_date}",
                    _kra_ver,
                    "",
                    _hb_ver,
                    ""
                ]]}
            ).execute()
            print(f"  Station Mapping updated for row {_row_num}")
        else:
            print("  Could not find station row — skipping sheet update")
    except Exception as _e:
        print(f"  Could not update Station Mapping: {_e}")

    # Enable Task Scheduler history
    # history enabled automatically # no technician action needed # all stations become diagnosable remotely
    subprocess.run(
        'wevtutil set-log Microsoft-Windows-TaskScheduler/Operational /enabled:true',
        shell=True,
        capture_output=True
    )
    # Enabling history:
    # does NOT noticeably affect performance
    # does NOT noticeably affect disk space
    # is standard in enterprise automation systems

    # ── Step 7: Create Scheduled Tasks ───────────────────────────────
    step(7, 7, "Creating scheduled tasks")

    kra_exe       = os.path.join(INSTALL_DIR, "kra_checker.exe")
    heartbeat_exe = os.path.join(INSTALL_DIR, "heartbeat_monitor.exe")

    # Create silent VBS launchers — Task Scheduler calls these
    # so the exe starts with zero visible window

    kra_vbs       = os.path.join(INSTALL_DIR, "run_kra_checker.vbs")
    heartbeat_vbs = os.path.join(INSTALL_DIR, "run_heartbeat.vbs")

    task_scheduler.create_vbs_launcher(
        kra_vbs,
        kra_exe,
        INSTALL_DIR
    )

    task_scheduler.create_vbs_launcher(
        heartbeat_vbs,
        heartbeat_exe,
        INSTALL_DIR
    )

    # Read schedule values from deployed config.json
    try:
        with open(os.path.join(INSTALL_DIR, "config.json")) as _f:
            _cfg = json.load(_f)
        heartbeat_interval = int(_cfg.get("heartbeat_interval", 30))
        kra_check_time     = _cfg.get("kra_check_time", "19:00")
    except Exception:
        heartbeat_interval = 30
        kra_check_time     = "19:00"

    print(f"  Heartbeat interval : every {heartbeat_interval} min")
    print(f"  KRA check time     : {kra_check_time}")


    tasks = [
        ("KRA Auto Checker",  f'wscript.exe "{kra_vbs}"',       f"/sc daily /st {kra_check_time}"),
        ("Station Heartbeat", f'wscript.exe "{heartbeat_vbs}"', f"/sc minute /mo {heartbeat_interval} /st 00:00"),
    ]

    for task_name, task_cmd, schedule in tasks:
        # windows_user, windows_pass
        ok = task_scheduler.create_or_update_task(
        task_name, task_cmd, schedule,
        )
        if not ok:
            print(f"  ERROR creating '{task_name}'")


    print("\n  Verifying installation:")
    for f in ["kra_checker.exe", "heartbeat_monitor.exe", "credentials.json", "config.json"]:
        exists = os.path.exists(os.path.join(INSTALL_DIR, f))
        print(f"  {'[OK]' if exists else '[MISSING]'} {f}")

    for task_name in ["KRA Auto Checker", "Station Heartbeat"]:
        r = subprocess.run(
            f'schtasks /query /tn "{task_name}"',
            shell=True, capture_output=True
        )
        print(f"  {'[OK]' if r.returncode == 0 else '[MISSING]'} Task: {task_name}")

    # ── Done ──────────────────────────────────────────────────────────
    banner("Installation Complete!")
    print(f"  Station  : {station_name}")
    print(f"  AnyDesk  : {detected_anydesk}")
    print(f"  Directory: {INSTALL_DIR}")
    print(f"""
  Tasks:
    KRA Auto Checker  - daily at {kra_check_time}
    Station Heartbeat - every {heartbeat_interval} min

  To test manually:
    {kra_exe}
    {heartbeat_exe}

  Results in Google Spreadsheet:
    Report tab         - KRA check results
    Station Status tab - connectivity status
    Logs tab           - detailed event log
""")
    input("Press Enter to exit...")


if __name__ == "__main__":
    main()
