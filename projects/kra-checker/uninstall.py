"""
KRA Auto-Checker Uninstaller
Removes scheduled tasks and the installation directory.
Self-elevates to Administrator.
"""

import os
import sys
import shutil
import subprocess
import ctypes

INSTALL_DIR  = r"C:\Automation_and_ops_engineering\KRA_Checker"
BASE_OPS_DIR = r"C:\Automation_and_ops_engineering"
TASK_NAMES   = ["KRA Auto Checker", "Station Heartbeat"]


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def self_elevate():
    if is_admin():
        return
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1
    )
    sys.exit(0)


def banner(msg):
    print(f"\n{'='*52}")
    print(f"  {msg}")
    print("=" * 52)


def main():
    self_elevate()

    banner("KRA Auto-Checker Uninstaller")

    if not os.path.exists(INSTALL_DIR):
        print(f"\n  Nothing to uninstall — {INSTALL_DIR} not found.")
        input("\nPress Enter to exit...")
        return

    print(f"\n  This will remove:")
    print(f"    • Scheduled tasks: {', '.join(TASK_NAMES)}")
    print(f"    • Installation folder: {INSTALL_DIR}")
    print()

    confirm = input("  Type YES to confirm uninstall: ").strip()
    if confirm.upper() != "YES":
        print("\n  Cancelled.")
        input("\nPress Enter to exit...")
        return

    # ── Remove scheduled tasks ────────────────────────────────────────
    print("\n  Removing scheduled tasks...")
    for task in TASK_NAMES:
        result = subprocess.run(
            f'schtasks /delete /tn "{task}" /f',
            shell=True, capture_output=True, text=True
        )
        if result.returncode == 0:
            print(f"  [OK] Removed task: {task}")
        else:
            # Task may not exist — not an error
            print(f"  [--] Task not found (already removed): {task}")

    # ── Remove retry task if it exists ───────────────────────────────
    retry_tasks = subprocess.run(
        'schtasks /query /fo CSV',
        shell=True, capture_output=True, text=True
    )
    for line in retry_tasks.stdout.splitlines():
        if "KRA_Retry_" in line:
            task_name = line.split(",")[0].strip('"')
            subprocess.run(
                f'schtasks /delete /tn "{task_name}" /f',
                shell=True, capture_output=True
            )
            print(f"  [OK] Removed retry task: {task_name}")

    # ── Remove installation directory ─────────────────────────────────
    print(f"\n  Removing {INSTALL_DIR}...")
    try:
        shutil.rmtree(INSTALL_DIR)
        print(f"  [OK] Removed: {INSTALL_DIR}")
    except Exception as e:
        print(f"  [ERROR] Could not remove folder: {e}")
        print(f"  Please delete manually: {INSTALL_DIR}")

    # ── Remove base ops dir if empty ──────────────────────────────────
    try:
        if os.path.exists(BASE_OPS_DIR) and not os.listdir(BASE_OPS_DIR):
            os.rmdir(BASE_OPS_DIR)
            print(f"  [OK] Removed empty folder: {BASE_OPS_DIR}")
        elif os.path.exists(BASE_OPS_DIR):
            print(f"  [--] Kept {BASE_OPS_DIR} (other projects still present)")
    except Exception:
        pass

    banner("Uninstall Complete")
    print("  All scheduled tasks and files have been removed.")
    print("  Google Sheets data is untouched.")
    print()
    _self_delete()
    input("Press Enter to exit...")


if __name__ == "__main__":
    main()


def _self_delete():
    """Drop a bat file that deletes the uninstall exe after this process exits."""
    try:
        exe_path = os.path.abspath(
            sys.executable if getattr(sys, "frozen", False) else __file__
        )
        if not exe_path.endswith(".exe"):
            return  # running as .py — nothing to clean up
        bat_path = os.path.join(os.path.dirname(exe_path), "cleanup_uninstaller.bat")
        with open(bat_path, "w") as f:
            f.write("@echo off\r\n")
            f.write("timeout /t 2 >nul\r\n")
            f.write(f'del /f /q "{exe_path}"\r\n')
            f.write('del /f /q "%~f0"\r\n')
            f.write("exit\r\n")
        subprocess.Popen(
            ["cmd", "/c", bat_path],
            creationflags=subprocess.CREATE_NO_WINDOW
        )
    except Exception:
        pass
