"""
backup_finder.py – Locate the latest ETIMS backup file.

Detection order:
  1. Use backup_path from config.json if explicitly set
  2. Try the default:  <home>/Documents/Backup
  3. Search recursively under <home> for any folder named 'Backup'
     that contains at least one ETIMS* file
  4. Give up and return None (caller will handle)
"""

from pathlib import Path
from typing import Optional
import logging

log = logging.getLogger(__name__)


def detect_backup_dir() -> Optional[Path]:
    """
    Auto-detect the backup directory without any config hint.
    Returns the Path if found, None otherwise.
    """
    home = Path.home()

    # 1. Default location
    default = home / "Documents" / "Backup"
    if _has_etims_files(default):
        log.info("Backup dir found at default location: %s", default)
        return default

    # 2. Search recursively under home for a folder called 'Backup'
    #    that actually contains ETIMS files
    log.info("Default backup path not found, searching under %s ...", home)
    try:
        for folder in sorted(home.rglob("Backup")):
            if folder.is_dir() and _has_etims_files(folder):
                log.info("Backup dir found by search: %s", folder)
                return folder
    except PermissionError:
        pass

    log.warning("Could not auto-detect a backup directory under %s", home)
    return None


def _has_etims_files(directory: Path) -> bool:
    """Return True if the directory exists and contains at least one ETIMS* file."""
    if not directory.exists():
        return False
    return any(
        f.is_file() and f.name.upper().startswith("ETIMS")
        for f in directory.iterdir()
    )


def find_latest_backup(backup_path_str: Optional[str] = None) -> tuple[Optional[Path], Optional[Path]]:
    """
    Resolve the backup directory and return (latest_file, backup_dir).
    backup_path_str can be:
      - None / empty  → auto-detect
      - Relative path → resolved from home dir
      - Absolute path → used as-is
    Returns (None, None) if nothing found.
    """
    backup_dir: Optional[Path] = None

    if backup_path_str:
        candidate = Path(backup_path_str)
        if not candidate.is_absolute():
            candidate = Path.home() / candidate
        if candidate.exists():
            backup_dir = candidate
        else:
            log.warning("Configured backup path not found: %s — falling back to auto-detect", candidate)

    if backup_dir is None:
        backup_dir = detect_backup_dir()

    if backup_dir is None:
        return None, None

    candidates = [
        f for f in backup_dir.iterdir()
        if f.is_file() and f.name.upper().startswith("ETIMS")
    ]

    if not candidates:
        log.warning("No ETIMS backup files found in %s", backup_dir)
        return None, backup_dir

    latest = max(candidates, key=lambda f: f.name)
    log.info("Backup candidates: %d  →  latest: %s", len(candidates), latest.name)
    return latest, backup_dir
