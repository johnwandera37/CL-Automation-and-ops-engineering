"""
KRA Auto-Checker Installer v2.2
Run as Administrator:  python install.py
"""

import os
import sys
import json
import shutil
import subprocess
import ctypes

INSTALL_DIR = r"C:\KRA_Checker"
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
    except:
        return False

def banner(msg):
    print(f"\n{'='*50}")
    print(f"  {msg}")
    print('='*50)

def step(n, total, msg):
    print(f"\n[{n}/{total}] {msg}...")

def main():
    banner("KRA Auto-Checker Installation v2.2")

    # ── Admin check ───────────────────────────────────────────────────
    if not is_admin():
        print("ERROR: Please run this script as Administrator.")
        print("  Right-click CMD -> Run as Administrator, then run: python install.py")
        input("\nPress Enter to exit...")
        sys.exit(1)

    script_dir = os.path.dirname(os.path.abspath(__file__))

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
    result = subprocess.run([sys.executable, "anydesk_detector.py"])
    detected_anydesk = None

    anydesk_file = os.path.join(INSTALL_DIR, "detected_anydesk.txt")
    if result.returncode == 0 and os.path.exists(anydesk_file):
        with open(anydesk_file) as f:
            detected_anydesk = f.read().strip()
        print(f"  AnyDesk ID: {detected_anydesk}")
    else:
        print("\n  WARNING: Could not auto-detect AnyDesk ID.")
        print("  Make sure AnyDesk is installed and has been opened at least once.")
        detected_anydesk = input("\n  Enter AnyDesk code manually (or leave blank to skip): ").strip()
        if not detected_anydesk:
            detected_anydesk = "UNKNOWN"
            print("  Skipping AnyDesk — set it manually in config.json later.")

    # ── Step 4: Look up station name ──────────────────────────────────
    step(4, 7, "Looking up station in Automation Helper sheet")
    station_name = None

    try:
        with open(os.path.join(INSTALL_DIR, "config.json")) as f:
            config = json.load(f)
        helper_sheet_id = config.get("automation_helper_sheet_id", "")

        if not helper_sheet_id:
            print("  WARNING: automation_helper_sheet_id not set in config.json")
        else:
            creds_file = os.path.join(INSTALL_DIR, "credentials.json")
            result = subprocess.run(
                [sys.executable, "fetch_station_info.py", creds_file, helper_sheet_id]
            )

            station_file = os.path.join(INSTALL_DIR, "detected_station.txt")
            if result.returncode == 0 and os.path.exists(station_file):
                with open(station_file) as f:
                    station_name = f.read().strip()
                print(f"  Station identified: {station_name}")
            else:
                print("  WARNING: AnyDesk code not found in Station Mapping sheet.")
                print("  Check that the code is listed in the Automation Helper spreadsheet.")

    except Exception as e:
        print(f"  WARNING: Could not read config: {e}")

    if not station_name:
        station_name = input("\n  Enter station name manually: ").strip()
        if not station_name:
            print("  ERROR: Station name cannot be empty.")
            input("\nPress Enter to exit...")
            sys.exit(1)

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
        # Delete existing
        subprocess.run(
            f'schtasks /delete /tn "{task_name}" /f',
            shell=True, capture_output=True
        )
        # Create new
        cmd = f'schtasks /create /tn "{task_name}" /tr "{exe_path}" {schedule} /f /rl highest'
        result = subprocess.run(cmd, shell=True, capture_output=True)
        if result.returncode == 0:
            print(f"  Created: {task_name}")
        else:
            print(f"  ERROR: Could not create task '{task_name}'")
            print(f"  {result.stderr.decode()}")

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
  Tasks created:
    KRA Auto Checker  - daily at 7:00 PM
    Station Heartbeat - every 30 minutes

  To test manually:
    {kra_exe}
    {heartbeat_exe}

  View results in your Google Spreadsheet:
    Report tab         - KRA check results
    Station Status tab - connectivity status
    Logs tab           - detailed event log
""")
    input("Press Enter to exit...")


if __name__ == "__main__":
    main()
