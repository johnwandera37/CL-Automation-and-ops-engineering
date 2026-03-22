"""
Config Loader
Merges local config.json with global settings from the Automation Helper Google Sheet.
Global sheet values OVERRIDE local values (except station-specific fields).
"""

import json
import os
import sys
import logging
from typing import Dict, Any

logger = logging.getLogger(__name__)

# Station-specific keys that should NEVER be overridden by the global sheet
STATION_ONLY_KEYS = {
    "station_name",
    "anydesk_code",
    "sql_server",
    "sql_database",
    "sql_username",
    "sql_password",
    "service_account_file",
    "spreadsheet_id",
    "automation_helper_sheet_id",
}

try:
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
    GOOGLE_AVAILABLE = True
except ImportError:
    GOOGLE_AVAILABLE = False
    logger.warning("Google API libraries not installed. Global config unavailable.")

# Base directory is wherever the .exe (or script) lives
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))


class ConfigLoader:
    """
    Loads and merges configuration from:
      1. Local config.json  (station-specific, always present)
      2. Automation Helper Google Sheet → 'Global Config' tab  (operator-controlled, overrides)
    """

    def __init__(self, local_config_file: str = "config.json"):
        self.local_config_path = os.path.join(BASE_DIR, local_config_file)
        self.config = self._build_config()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get(self, key: str, default=None):
        return self.config.get(key, default)

    def __getattr__(self, name: str):
        if name.startswith("_") or name == "config":
            raise AttributeError(name)
        try:
            return self.config[name]
        except KeyError:
            raise AttributeError(f"Config has no key '{name}'")

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _build_config(self) -> Dict[str, Any]:
        local = self._load_local()
        global_cfg = self._load_global(local) if GOOGLE_AVAILABLE else {}

        # Global overrides local, EXCEPT for station-specific keys
        merged = dict(local)
        for key, value in global_cfg.items():
            if key not in STATION_ONLY_KEYS:
                merged[key] = value

        # Resolve paths relative to BASE_DIR
        for path_key in ("service_account_file", "retry_file"):
            if path_key in merged and not os.path.isabs(merged[path_key]):
                merged[path_key] = os.path.join(BASE_DIR, merged[path_key])

        return merged

    def _load_local(self) -> Dict[str, Any]:
        if not os.path.exists(self.local_config_path):
            logger.error(f"Local config not found: {self.local_config_path}")
            return {}
        try:
            with open(self.local_config_path, "r") as f:
                data = json.load(f)
            logger.info("Loaded local config.json")
            return data
        except json.JSONDecodeError as e:
            logger.error(f"config.json is invalid JSON: {e}")
            return {}

    def _load_global(self, local: Dict) -> Dict[str, Any]:
        """Pull the 'Global Config' sheet from the Automation Helper spreadsheet."""
        try:
            sa_file = os.path.join(
                BASE_DIR, local.get("service_account_file", "credentials.json")
            )
            helper_id = local.get("automation_helper_sheet_id")

            if not helper_id:
                logger.warning("automation_helper_sheet_id not set — skipping global config")
                return {}

            if not os.path.exists(sa_file):
                logger.warning(f"Service account file not found: {sa_file}")
                return {}

            creds = service_account.Credentials.from_service_account_file(
                sa_file,
                scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
            )
            service = build("sheets", "v4", credentials=creds)

            result = (
                service.spreadsheets()
                .values()
                .get(spreadsheetId=helper_id, range="Global Config!A:B")
                .execute()
            )

            rows = result.get("values", [])
            global_cfg: Dict[str, Any] = {}

            for row in rows[1:]:  # skip header row
                if len(row) < 2:
                    continue
                key = row[0].strip()
                raw = row[1].strip()

                # Type coercion
                if raw.lower() in ("true", "false"):
                    value: Any = raw.lower() == "true"
                elif raw.isdigit():
                    value = int(raw)
                elif key == "retry_hours" and "," in raw:
                    value = [int(x.strip()) for x in raw.split(",") if x.strip().isdigit()]
                else:
                    value = raw

                global_cfg[key] = value

            logger.info(f"Loaded {len(global_cfg)} keys from Global Config sheet")
            return global_cfg

        except Exception as e:
            logger.warning(f"Could not load global config from sheet: {e}")
            return {}
