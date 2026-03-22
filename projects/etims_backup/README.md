# ETIMS Backup Uploader

Automatically uploads the latest ETIMS database backup from a petrol station mini PC to Google Shared Drive every day using a service account — no browser login required.

---

## How It Works

- Runs daily via Windows Task Scheduler (default 1:00 AM)
- Finds the latest `ETIMS*` file in the station's Backup folder
- Uploads it to the correct folder on Google Shared Drive, organized by station type and name
- Skips upload if the file is already on Drive (safe to re-run)
- If the PC was offline at the scheduled time, runs automatically when it comes back online
- Waits up to 2 hours for internet if the connection is down
- Retries 3 times on failure before giving up
- Logs critical failures to a shared Google Sheet (`ETIMSBackupLog`) in red text

---

## Drive Folder Structure

```
ETIMS Backups/              ← Shared Drive root
├── Rubis/
│   ├── Eldoret Town/
│   │   ├── ETIMS202603130000.zip
│   │   └── ETIMS202603140000.zip
│   └── Nakuru West/
│       └── ETIMS202603140000.zip
├── Shell/
├── Total/
├── Regnol/
└── ETIMSBackupLog          ← Google Sheet, critical errors only
```

---

## Project Structure

```
etims_backup/
├── main.py               # Entry point — orchestrates everything
├── drive_service.py      # All Google Drive API calls
├── backup_finder.py      # Auto-detects latest ETIMS backup file
├── error_logger.py       # Writes red error rows to Google Sheet
├── config.json           # Station settings (pre-filled before deployment)
├── requirements.txt      # Python dependencies
├── build.bat             # Builds main.exe using PyInstaller (run on Windows)
├── install.bat           # Interactive installer (run on each station)
├── uninstall.bat         # Removes program and scheduled task
└── register_task.ps1     # Registers Windows Task Scheduler job (called by install.bat)
```

---

## One-Time Setup (Administrator)

These steps are done once before deploying to any station.

### 1. Google Shared Drive

1. Go to [drive.google.com](https://drive.google.com) → **Shared drives** → **+ New**
2. Name it `ETIMS Backups`
3. Create station folders inside it: `Rubis`, `Shell`, `Total`, `Regnol`
4. Right-click the Shared Drive → **Manage members**
5. Add the service account as **Content Manager**:
   ```
   kra-checker-service@kra-auto-checker.iam.gserviceaccount.com
   ```
6. Create a Google Sheet in the root named `ETIMSBackupLog` and share it with the same service account as **Editor**

### 2. Get the IDs

**Drive root folder ID** — open the Shared Drive root in a browser and copy the ID from the URL:
```
https://drive.google.com/drive/folders/<FOLDER_ID_HERE>
```

**Sheet ID** — open ETIMSBackupLog and copy the ID from the URL:
```
https://docs.google.com/spreadsheets/d/<SHEET_ID_HERE>/edit
```

### 3. Fill in config.json

Open `config.json` and set the two IDs before packaging the deployment folder:

```json
{
  "service_account_file": "credentials.json",
  "service_account_email": "kra-checker-service@kra-auto-checker.iam.gserviceaccount.com",
  "station_type": "Rubis",
  "station_name": "Eldoret Town",
  "drive_root_folder_id": "YOUR_SHARED_DRIVE_FOLDER_ID",
  "log_sheet_id": "YOUR_SHEET_ID",
  "schedule_time": "01:00",
  "max_retries": 3,
  "retry_interval_seconds": 300,
  "max_wait_minutes": 120,
  "log_spreadsheet_name": "ETIMSBackupLog"
}
```

> `station_type` and `station_name` don't need to be correct here — `install.bat` overwrites them per station.

### 4. Build the Executable (Windows)

```bat
pip install -r requirements.txt
build.bat
```

Output: `dist\main.exe`

---

## Deploying to a Station

### Deployment package (6 files)

```
main.exe
credentials.json
config.json          ← drive_root_folder_id and log_sheet_id filled in
install.bat
uninstall.bat
register_task.ps1
```

### Installation steps

1. Copy the 6 files to any folder on the station PC
2. Double-click **`install.bat`**
3. Enter the station in the format `Type Name` e.g. `Rubis Eldoret Town`
4. Confirm the auto-detected backup path
5. Press Enter to accept the default run time (`01:00`) or enter a custom time
6. Click **Yes** on the UAC prompt
7. Done — the program installs to `C:\ProgramData\ETIMSBackup\`

`install.bat` automatically:
- Copies files to `C:\ProgramData\ETIMSBackup\`
- Writes `config.json` with the station name and backup path
- Registers the Task Scheduler job (runs as SYSTEM, no login required)
- Enables `StartWhenAvailable` — missed runs fire at next boot
- Tests the Google Drive connection

### Backup path auto-detection

The installer searches for the ETIMS backup folder in this order:

1. `Documents\Backup` — default for most stations
2. Recursive search under the user's home for any `Backup` folder containing `ETIMS*` files
3. If not found, asks for the full path manually

---

## Verifying the Installation

### Check the scheduled task

```
Win + R → taskschd.msc → Task Scheduler Library → ETIMSBackupUploader
```

Confirm the **Next Run Time** shows the correct scheduled time.

### Run manually

```cmd
"C:\ProgramData\ETIMSBackup\main.exe"
```

Expected output:
```
2026-03-15 01:00:01  INFO     === Attempt 1 / 3 ===
2026-03-15 01:00:01  INFO     Backup candidates: 1  →  latest: ETIMS202603140000.zip
2026-03-15 01:00:01  INFO     Using backup directory: C:\Users\Administrator\Documents\Backup
2026-03-15 01:00:04  INFO     Uploading ETIMS202603140000.zip ...
2026-03-15 01:00:07  INFO     Upload complete.
```

### Check the log file

```
C:\ProgramData\ETIMSBackup\etims_backup.log
```

The log rotates automatically at 2MB, keeping up to 4 files (~8MB total). No manual cleanup needed.

---

## CLI Reference

| Command | Description |
|---|---|
| `main.exe` | Normal upload (used by Task Scheduler) |
| `main.exe --test` | Test Google Drive connectivity |
| `main.exe --cleanup` | Delete all but the latest backup from Drive |
| `main.exe --cleanup --keep 3` | Keep 3 most recent, delete the rest |

---

## Uninstalling

```cmd
C:\ProgramData\ETIMSBackup\uninstall.bat
```

Removes the scheduled task and deletes all files from `C:\ProgramData\ETIMSBackup`. Asks for confirmation before proceeding.

---

## Error Log

The `ETIMSBackupLog` Google Sheet is shared across all stations — every station writes to the same sheet.

| Column | Content |
|---|---|
| Timestamp | Date and time of the failure |
| Station Type | e.g. Rubis |
| Station Name | e.g. Eldoret Town |
| Error Message | Full description of what went wrong |

- Only critical failures are logged — successful uploads are never written
- Error rows appear in **red bold text**
- Rows are permanent and never auto-deleted

---

## Changing Settings After Installation

### Change the run time

Edit `schedule_time` in `C:\ProgramData\ETIMSBackup\config.json`, then re-run `register_task.ps1` as Administrator.

### Change the service account

Replace `credentials.json` in `C:\ProgramData\ETIMSBackup\` and update `service_account_email` in `config.json`. Ensure the new account is added to the Shared Drive as **Content Manager**.

### Change the Drive root folder

Update `drive_root_folder_id` in `config.json`. No rebuild required — config is read fresh on every run.

---

## Requirements

- Python 3.11+
- Windows (for deployment)
- Google Cloud project with **Drive API** and **Sheets API** enabled
- Service account with a JSON key (`credentials.json`)
- Google Workspace account with Shared Drives enabled

### Python dependencies

```
google-auth
google-auth-oauthlib
google-api-python-client
openpyxl
pyinstaller
```

---

## Important Notes

- **Shared Drive required** — service accounts have no storage quota of their own. Files must be uploaded to a Shared Drive (Google Workspace feature). Personal Google accounts do not support Shared Drives.
- **Keep `credentials.json` private** — never commit it to version control. Add it to `.gitignore`.
- **Config is reloaded on every retry** — editing `config.json` while the program is running takes effect on the next attempt.

---

## .gitignore

```
credentials.json
config.json
*.log
dist/
build/
__pycache__/
*.spec
```
