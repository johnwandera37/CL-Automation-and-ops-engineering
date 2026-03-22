"""
Auto Updater
Checks the global config sheet for a newer version and, if found,
downloads the updated .exe from Google Drive and hot-swaps it via a
small .bat launcher — then exits so Windows restarts the new binary.

Called at the TOP of kra_auto_checker.py and heartbeat_monitor.py
BEFORE any main logic runs.
"""

import os
import sys
import logging
import subprocess

logger = logging.getLogger(__name__)


def check_and_update() -> None:
    """
    Entry point called by the two main executables.
    Fails silently on any error so it never blocks the main program.
    """
    try:
        _run_update()
    except Exception as e:
        logger.warning(f"[UPDATER] Skipped due to error: {e}")


# ──────────────────────────────────────────────────────────────────────────────

def _run_update() -> None:
    try:
        from packaging import version
    except ImportError:
        logger.warning("[UPDATER] 'packaging' not installed — skipping update check")
        return

    # Late import so auto_updater doesn't fail if config_loader has issues
    from config_loader import ConfigLoader

    config = ConfigLoader()

    exe_path = os.path.abspath(sys.argv[0])
    exe_name = os.path.basename(exe_path)

    # Determine which Drive file ID to use
    is_kra = "kra" in exe_name.lower()
    drive_id_key = "kra_checker_drive_id" if is_kra else "heartbeat_monitor_drive_id"
    drive_file_id = config.get(drive_id_key)

    if not drive_file_id:
        logger.info("[UPDATER] No Drive file ID configured — skipping")
        return

    local_version = config.get("current_version", "0.0.0")
    # remote_version comes from the Global Config sheet (already merged)
    remote_version = config.get("remote_version", local_version)

    logger.info(f"[UPDATER] Local: {local_version}  Remote: {remote_version}")

    if version.parse(remote_version) <= version.parse(local_version):
        logger.info("[UPDATER] Already up to date")
        return

    logger.info("[UPDATER] Newer version available — downloading...")
    _download_and_apply(drive_file_id, exe_path, exe_name)


def _download_and_apply(drive_file_id: str, exe_path: str, exe_name: str) -> None:
    import requests  # only needed if an update is actually available

    download_url = f"https://drive.google.com/uc?export=download&id={drive_file_id}"
    new_exe = exe_path + ".new"
    backup_exe = exe_path + ".old"

    response = requests.get(download_url, stream=True, timeout=60)
    response.raise_for_status()

    with open(new_exe, "wb") as f:
        for chunk in response.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)

    # Validate: must be at least 100 KB
    if not os.path.exists(new_exe) or os.path.getsize(new_exe) < 100 * 1024:
        logger.error("[UPDATER] Downloaded file is too small — aborting update")
        if os.path.exists(new_exe):
            os.remove(new_exe)
        return

    logger.info("[UPDATER] Download OK — preparing hot-swap")

    # Write a self-deleting batch file that replaces the exe and restarts it
    bat_path = os.path.join(os.path.dirname(exe_path), "apply_update.bat")
    bat_content = f"""@echo off
timeout /t 2 >nul
if exist "{backup_exe}" del /f /q "{backup_exe}"
if exist "{exe_path}"   ren "{exe_path}" "{os.path.basename(backup_exe)}"
ren "{new_exe}" "{exe_name}"
start "" "{exe_path}"
del /f /q "%~f0"
exit
"""
    with open(bat_path, "w") as f:
        f.write(bat_content)

    subprocess.Popen(
        ["cmd", "/c", bat_path],
        creationflags=subprocess.CREATE_NO_WINDOW,
    )

    logger.info("[UPDATER] Launcher started — exiting for restart")
    sys.exit(0)
