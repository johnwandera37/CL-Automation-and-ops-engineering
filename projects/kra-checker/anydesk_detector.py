"""
AnyDesk Detector
Reads the local AnyDesk ID from the Windows registry or config files.
Designed to run on Windows; gracefully skips on other platforms.
"""

import sys
import os

def get_anydesk_id() -> str | None:
    """
    Detect the AnyDesk ID for this PC.
    Tries three locations in order:
      1. Windows registry  (HKLM\\SOFTWARE\\WOW6432Node\\AnyDesk)
      2. User-level config (%APPDATA%\\AnyDesk\\system.conf)
      3. System-level config (C:\\ProgramData\\AnyDesk\\system.conf)
    Returns the ID as a string, or None if not found.
    """

    # ── Method 1: Windows registry ──────────────────────────────────
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\WOW6432Node\AnyDesk",
            0,
            winreg.KEY_READ,
        )
        anydesk_id, _ = winreg.QueryValueEx(key, "ad.anynet.id")
        winreg.CloseKey(key)
        print(f"✓ AnyDesk ID from registry: {anydesk_id}")
        return str(anydesk_id)
    except ImportError:
        pass  # Not on Windows — continue to file-based methods
    except Exception:
        pass  # Key not present — try next method

    # ── Method 2: User-level config (%APPDATA%\\AnyDesk\\system.conf) ─
    try:
        appdata = os.getenv("APPDATA")
        if appdata:
            config_path = os.path.join(appdata, "AnyDesk", "system.conf")
            result = _parse_anydesk_conf(config_path, "user-level config")
            if result:
                return result
    except Exception:
        pass

    # ── Method 3: System-level config (ProgramData) ──────────────────
    try:
        programdata = os.getenv("ProgramData", r"C:\ProgramData")
        config_path = os.path.join(programdata, "AnyDesk", "system.conf")
        result = _parse_anydesk_conf(config_path, "system-level config")
        if result:
            return result
    except Exception:
        pass

    print("✗ Could not detect AnyDesk ID")
    return None


def _parse_anydesk_conf(path: str, source_label: str) -> str | None:
    """Parse ad.anynet.id from an AnyDesk system.conf file."""
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r") as f:
            for line in f:
                if line.strip().startswith("ad.anynet.id"):
                    anydesk_id = line.split("=", 1)[1].strip()
                    print(f"✓ AnyDesk ID from {source_label}: {anydesk_id}")
                    return anydesk_id
    except Exception as e:
        print(f"  Could not read {path}: {e}")
    return None


if __name__ == "__main__":
    anydesk_id = get_anydesk_id()
    if anydesk_id:
        with open("detected_anydesk.txt", "w") as f:
            f.write(anydesk_id)
        print(f"Saved to detected_anydesk.txt")
    sys.exit(0 if anydesk_id else 1)
