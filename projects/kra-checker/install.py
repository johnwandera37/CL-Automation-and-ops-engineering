"""
KRA Auto-Checker Installer v2.4
- Auto-detects AnyDesk ID
- Looks up station name from Automation Helper sheet
- If station not found, prompts for name and writes it back to the sheet
- Self-elevates to Administrator (double-click friendly)
- No Python required on target stations (build with PyInstaller)
"""

import os
import sys
import json
import shutil
import subprocess
import ctypes
import logging

# Suppress google file_cache warning globally
logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.ERROR)

BASE_OPS_DIR = r"C:\Automation_and_ops_engineering"
INSTALL_DIR  = os.path.join(BASE_OPS_DIR, "KRA_Checker")
REQUIRED_FILES = [
    "kra_checker.exe",
    "heartbeat_monitor.exe",
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

    # ── Step 7: Create Scheduled Tasks ───────────────────────────────
    step(7, 7, "Creating scheduled tasks")

    kra_exe       = os.path.join(INSTALL_DIR, "kra_checker.exe")
    heartbeat_exe = os.path.join(INSTALL_DIR, "heartbeat_monitor.exe")

    tasks = [
        ("KRA Auto Checker",  kra_exe,       "/sc daily /st 19:00"),
        ("Station Heartbeat", heartbeat_exe, "/sc minute /mo 30"),
    ]

    for task_name, exe_path, schedule in tasks:
        subprocess.run(
            f'schtasks /delete /tn "{task_name}" /f',
            shell=True, capture_output=True
        )
        result = subprocess.run(
            f'schtasks /create /tn "{task_name}" /tr "{exe_path}" {schedule} /f /rl highest',
            shell=True, capture_output=True
        )
        if result.returncode == 0:
            print(f"  Created: {task_name}")
        else:
            print(f"  ERROR creating task '{task_name}': {result.stderr.decode()}")

    # ── Verification ──────────────────────────────────────────────────
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
    KRA Auto Checker  - daily at 7:00 PM
    Station Heartbeat - every 30 minutes

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
