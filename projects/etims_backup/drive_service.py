"""
drive_service.py – Wraps all Google Drive API operations.
Supports both regular Drive folders and Shared Drives (supportsAllDrives=True).
"""

import os
import io
from pathlib import Path
from typing import Optional

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]


class DriveService:
    def __init__(self, credentials_path: str):
        creds = service_account.Credentials.from_service_account_file(
            credentials_path, scopes=SCOPES
        )
        self._service = build("drive", "v3", credentials=creds, cache_discovery=False)

    # ── Folder helpers ────────────────────────────────────────────────────────

    def get_or_create_folder(self, name: str, parent_id: str) -> str:
        query = (
            f"name='{name}' "
            f"and mimeType='application/vnd.google-apps.folder' "
            f"and '{parent_id}' in parents "
            f"and trashed=false"
        )
        result = self._service.files().list(
            q=query,
            spaces="drive",
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()

        files = result.get("files", [])
        if files:
            return files[0]["id"]

        meta = {
            "name": name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id],
        }
        folder = self._service.files().create(
            body=meta,
            fields="id",
            supportsAllDrives=True,
        ).execute()
        return folder["id"]

    # ── File helpers ──────────────────────────────────────────────────────────

    def file_exists(self, filename: str, folder_id: str) -> bool:
        query = (
            f"name='{filename}' "
            f"and '{folder_id}' in parents "
            f"and trashed=false"
        )
        result = self._service.files().list(
            q=query,
            spaces="drive",
            fields="files(id)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        return bool(result.get("files"))

    def upload_file(self, local_path: str, folder_id: str) -> str:
        name  = Path(local_path).name
        meta  = {"name": name, "parents": [folder_id]}
        media = MediaFileUpload(local_path, resumable=True)
        uploaded = self._service.files().create(
            body=meta,
            media_body=media,
            fields="id",
            supportsAllDrives=True,
        ).execute()
        return uploaded["id"]

    def list_backup_files(self, folder_id: str) -> list[dict]:
        query = (
            f"name contains 'ETIMS' "
            f"and '{folder_id}' in parents "
            f"and trashed=false"
        )
        result = self._service.files().list(
            q=query,
            spaces="drive",
            fields="files(id, name, createdTime)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        return result.get("files", [])

    def delete_file(self, file_id: str):
        self._service.files().delete(
            fileId=file_id,
            supportsAllDrives=True,
        ).execute()

    # ── Spreadsheet helpers ───────────────────────────────────────────────────

    def find_file(self, name: str, folder_id: Optional[str] = None) -> Optional[str]:
        query = f"name='{name}' and trashed=false"
        if folder_id:
            query += f" and '{folder_id}' in parents"
        result = self._service.files().list(
            q=query,
            spaces="drive",
            fields="files(id)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files = result.get("files", [])
        return files[0]["id"] if files else None

    def download_file_bytes(self, file_id: str) -> bytes:
        request = self._service.files().get_media(fileId=file_id)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return buf.getvalue()

    def update_file_bytes(self, file_id: str, data: bytes,
                          mime: str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
        import tempfile
        suffix = ".xlsx" if "spreadsheetml" in mime else ".bin"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        media = MediaFileUpload(tmp_path, mimetype=mime, resumable=False)
        self._service.files().update(
            fileId=file_id,
            media_body=media,
            supportsAllDrives=True,
        ).execute()
        os.unlink(tmp_path)

    def create_file_from_bytes(self, name: str, data: bytes, folder_id: str,
                               mime: str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") -> str:
        import tempfile
        suffix = ".xlsx" if "spreadsheetml" in mime else ".bin"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        meta  = {"name": name, "parents": [folder_id]}
        media = MediaFileUpload(tmp_path, mimetype=mime, resumable=False)
        result = self._service.files().create(
            body=meta,
            media_body=media,
            fields="id",
            supportsAllDrives=True,
        ).execute()
        os.unlink(tmp_path)
        return result["id"]
