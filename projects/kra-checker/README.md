# KRA Auto-Checker System v2.1

An automated ops system that monitors KRA eTIMS transaction submissions across petrol station mini PCs. Runs silently in the background via Windows Task Scheduler, reports to Google Sheets, and self-updates remotely.

---

## What It Does

Each station mini PC runs two background programs:

| Program | Schedule | Purpose |
|---|---|---|
| `kra_checker.exe` | Daily at configured time (default 7 PM) | Pulls two random ETIMS transaction from SQL Server, one Fuel Card and any other oayment mode transaction, checks their QR links against KRA portal, writes result to Google Sheets |
| `heartbeat_monitor.exe` | Every N minutes (default 30) | Checks internet, SQL, disk space — updates station status in Google Sheets. Also manages remote updates and config sync for both programs |

tasks run under SYSTEM
tasks run:
while locked
while logged out
after reboot
silently via VBS launcher

---

## System Architecture

```
Google Drive                    Automation Helper Sheet
(exe files)                     (Global Config + Station Mapping)
     │                                      │
     └──────────────┬───────────────────────┘
                    │
          heartbeat_monitor.exe  (runs every x min, (30 min recommneded to reduce API calls and not affect PC performance, but it's extremely light) on each station)
                    │
          ┌─────────┴──────────────────────────┐
          │                                    │
    Check for updates                   Sync config values
    (kra + heartbeat)                   (intervals, retry hours,
          │                              timeouts, check time)
          │                                    │
    Download & swap exe            Update Task Scheduler
    if remote > local              if schedule changed
          │                                    │
    Write Last Seen timestamp + interval     Write to Station Status sheet
    to Station Status sheet
                    │
          Google Apps Script (runs on Google servers every 1 min which forces Now() recalculation, Ensure your `Timezone` is selected to work correctly)
                    │
          Refreshes Last Seen + Forcing the new formula to transition correctly
          🟢 Online / 🟡 Stale / 🔴 Offline to Status column
                    │
          kra_checker.exe  (runs daily at configured time)
                    │
          Query SQL Server (ETPumpSales table)
          Expand time window if no transactions found
                    │
          Check QR link against KRA portal
                    │
          Write result to Report sheet
          Retry overnight if not submitted
```

### KRA Checker Logic
The checker:

Queries normal pump sales transactions
Queries fuel card transactions
Selects recent/random valid transactions
Verifies QR links against KRA portal
Writes results to Google Sheets

multiple transaction sources are supported
fallback logic exists
wider compatibility across station POS systems

| Source           | Table                 |
| ---------------- | ---------------------------- |
| Any payment mode POS sales | ETPumpSales                  |
| Fuel card sales  | ETPumpSales                  


### Retry & Recovery Mechanism

When a transaction cannot be verified successfully, the system saves it locally for automatic retry later.

This handles situations such as:

Temporary internet outages
KRA portal downtime
Connection resets
Timeouts
DNS failures
Slow station networks

Failed checks are written to a local JSON retry queue:
```
{
    "retry_count": 2,
    "saved_at": "2026-05-07 23:25:21",
    "transactions": [
        {
            "transaction_date": "2025-08-19 03:36:10",
            "qr_link": "https://etims.kra.go.ke/common/link/etims/receipt/indexEtimsReceiptData?Data=P052160332H01276VWTSUNMQTCBXZ",
            "check_date": "2025-08-19",
            "payment_mode": "CASH",
            "last_status": "ERROR"
        }
    ]
}
```

#### How retries work
Initial transaction check fails\
Transaction is saved locally into retry storage\
Heartbeat/global config defines retry hours\
KRA checker retries automatically later\
Successful retry removes the transaction from queue\
Retry count prevents infinite loops\
Supported transaction types

Retries work for both:

Normal POS fuel transactions
Fuel card transactions
Retry scheduling

#### Controlled remotely from the Global Config sheet:
| Key           | Example | Meaning                |
| ------------- | ------- | ---------------------- |
| `retry_hours` | `0,2,4` | `Retry after 0h, 2h, 4h` |

This allows failed submissions to recover automatically overnight without technician intervention.


### Task Scheduler Architecture

### Task Scheduler Design

The system uses Windows Task Scheduler for fully automated background execution.

SYSTEM account execution

Tasks run under the built-in `Windows SYSTEM account.`

#### Benefits:

No Windows password required during installation
Runs even when no user is logged in
Runs while workstation is locked
Survives reboot automatically
Higher reliability than user-bound tasks
Silent execution

Tasks launch through hidden VBScript (.vbs) wrappers:

`wscript.exe run_heartbeat.vbs`\
`wscript.exe run_kra_checker.vbs`

#### This prevents:

console windows\
flashing CMD prompts\
user disruption\

Programs run completely invisibly in the background.

#### Advanced scheduler settings

After task creation, PowerShell applies advanced settings:

Allow running on battery\
Do not stop on battery\
Run missed tasks automatically\
Wake machine if needed\
Ignore overlapping runs\
Unlimited execution time\
Automatic schedule updates

heartbeat_monitor.exe continuously syncs Task Scheduler settings from Google Sheets.

Changes made remotely to:

heartbeat_interval\
kra_check_time

are automatically applied on all stations without manual intervention.

Missed task recovery

If a station is:

powered off\
asleep\
disconnected\

scheduled tasks automatically run once the machine becomes available again.

---

## Google Sheets Structure

### Rubis Stations Monitoring Allocations (Results Spreadsheet)

| Sheet | Purpose | Written by |
|---|---|---|
| `Report` | KRA check results per station per day | kra_checker.exe |
| `Logs` | Detailed event log for all stations | both |
| `Station Status` | Live connectivity dashboard | heartbeat_monitor.exe + Apps Script |
| `Heartbeat Error Logs` | Error-level events only | heartbeat_monitor.exe |

Only the service account (via API) and `clear_sheets.py` for development can modify data.

**Date separator rows:** The Report, Logs, and Heartbeat Error Logs sheets automatically insert a merged date row whenever the date changes:
```
╔25/04/2026 ════════════════════════════╗   ←  merged
  TKDesk         🟢 SUCCESS    ...
  WincateDesk    🔴 NOT SUBMITTED    ...
╔ 26/04/2026 ════════════════════════════╗
  ...
```
Date separators are cached locally — only one API call per new day, not on every write.

### Automation Helper (Control Spreadsheet)

| Sheet | Purpose |
|---|---|
| `Global Config` | Settings that apply to all stations |
| `Station Mapping` | AnyDesk code → station name mapping + install/update status |

---

## Station Status — How Online/Offline Works

The Status column uses a **formula** written automatically by the program. It compares `Last Seen` against `NOW()` using the station's heartbeat interval.

**Why `NOW()` needs help:** Google Sheets `NOW()` only recalculates when the sheet is opened or edited. A Google Apps Script runs every minute on Google's servers and writes to a hidden helper cell (Z1), forcing `NOW()` to recalculate continuously — even when nobody has the sheet open.

**How it works end to end:**
- Heartbeat writes `Last Seen` as a timestamp to column C
- Heartbeat writes the configured interval (minutes) to column I
- The simplified formula in column D computes minutes elapsed since last seen
- Apps Script triggers `NOW()` recalculation every 1 minute
- Status updates automatically with no station involvement
- `Timezone` MUST be specified for script to work correctly, Got to File > Settings > Update Timezone 

**Status thresholds:**

| Status | Condition | Meaning |
|---|---|---|
| `🟢 Online` | Last seen within 1× interval | Station running normally |
| `🟡 Stale` | Last seen within 2× interval | Missed one cycle — may recover |
| `🔴 Offline` | Last seen > 2× interval ago | Station not running |
| `⚪ Never` | Column C empty | Installed but never ran |

**Example:** Heartbeat interval = 30 min. Station goes offline at 10:00.
- 10:01 → Apps Script fires → `🟢 Online` (only 1 min elapsed)
- 10:31 → Apps Script fires → `🟡 Stale` (31 min > 1× interval)
- 11:01 → Apps Script fires → `🔴 Offline` (61 min > 2× interval)

**Key insight:** A station can only report that it is **online**. If it goes offline, nothing on the station can send that report. The formula + Apps Script combination fills this gap entirely on the Google side — no station involvement needed.

---

## Apps Script Setup (One Time)

This is required so the Status column (Online/Stale/Offline) updates in real time without anyone having the sheet open.

**Why it is needed:** Google Sheets `NOW()` only recalculates when the sheet is opened or edited. Without this script, a station could go offline but still show 🟢 Online for hours. The script writes to a hidden helper cell every minute, which forces `NOW()` to recalculate and the Status formula to update.

**Steps:**
1. Open the **Rubis Stations Monitoring Allocations** spreadsheet
2. Click **Extensions → Apps Script**
3. Delete any existing code and paste the contents of `station_status_monitor.gs`
4. Click **Save**
5. Select `createTrigger` from the function dropdown
6. Click **Run** and accept the permissions prompt
7. Done — the trigger runs every minute automatically, forever
8. Ensure `Timezone` is set based on ones's time zone, e.g `(GMT+03:00) Nairobi` In Sheet go Files > Settings

**After setup:** Hide column Z in the Station Status sheet (right-click column Z → Hide column). This is the helper cell the script writes to.

**The Status formula** (written automatically by the program into column D):
```
formula = (
            f'=IF('
            f'C{row}="",'
            f'"⚪ Never",'
            f'IF('
            f'(NOW()-C{row})*1440>I{row}*2,'
            f'"🔴 Offline",'
            f'IF('
            f'(NOW()-C{row})*1440>I{row},'
            f'"🟡 Stale",'
            f'"🟢 Online"'
            f')'
            f')'
            f')'
        )
```
- `C{row number}` = Last Seen (stored as Sheets datetime serial) for each station
- `I2{row number}` = Heartbeat Interval in minutes for each station
- `NOW()-C2*1440>I{row}` = minutes elapsed since last heartbeat
- `Stale threshold` = 1× interval, Offline threshold = 2× interval

---

## Global Config Sheet Reference

Changes take effect on every station within one heartbeat cycle. No manual intervention on any station.

| Key | Default | Description |
|---|---|---|
| `max_retries` | 3 | KRA link check retry attempts |
| `retry_delay` | 30 | Seconds between retry attempts |
| `timeout` | 15 | HTTP request timeout in seconds |
| `retry_hours` | 0,2,4 | Hours to retry failed transactions overnight |
| `heartbeat_interval` | 30 | Minutes between heartbeat runs. Heartbeat automatically updates the Windows Task Scheduler trigger remotely |
| `kra_check_time` | 19:00 | Daily time to run KRA checker. Heartbeat automatically updates the Windows Task Scheduler trigger remotely |
| `remote_version_kra` | 1.0.0 | Bump to trigger kra_checker.exe update on all stations |
| `remote_version_heartbeat` | 1.0.0 | Bump to trigger heartbeat_monitor.exe update on all stations |
| `kra_checker_drive_id` | — | Google Drive file ID for kra_checker.exe |
| `heartbeat_monitor_drive_id` | — | Google Drive file ID for heartbeat_monitor.exe |
| `spreadsheet_id` | — | Results spreadsheet ID (can be switched remotely) |

---

## Station Mapping Sheet Reference

| Column | Content | Written by |
|---|---|---|
| A | Station name | install.py or manually |
| B | AnyDesk code | install.py or manually |
| C | Installed date | install.py on first install |
| D | KRA Checker version/status | install.py + heartbeat on update |
| E | KRA update timestamp | heartbeat on update |
| F | Heartbeat version/status | install.py + heartbeat on update |
| G | Heartbeat update timestamp | heartbeat on update |

---

## API Call Budget

The system is designed to stay well within Google Sheets API limits (60 reads/writes per minute per user) across 100+ stations.

| Operation | API calls | Frequency |
|---|---|---|
| Load Global Config | 1 read | Every run |
| Write station status | 1 write | Every heartbeat |
| Write KRA result | 1 write | Once daily |
| Write log entry | 1 write | Per event |
| Date separator | 2 calls | **Once per day** (flag file cached) |
| Task sync | 0 reads | Compares against config.json, no API |
| Auto-update check | Drive API | 60s thread timeout, fails safely |

**Key design decisions:**
- Date separators use a local `.lastdate_SheetName` flag file — 2 API calls once per day, then zero
- Task Scheduler sync compares against `config.json` cache, not `schtasks /query` — zero API calls
- Auto-updater runs in a background thread with 60 second timeout — never blocks the heartbeat

---

## Deployment

### Folder Structure

**Builder folder** (your development machine):
```
Builddet/
├── kra_auto_checker.py
├── heartbeat_monitor.py
├── config_loader.py
├── anydesk_detector.py
├── fetch_station_info.py
├── install.py
├── uninstall.py
├── config.json
├── credentials.json
├── clean_build.bat
├── task_scheduler
└── requirements.txt this should be ran when setting up the program for development, testing executables snd testing python source
```

**Installer folder** (copied to each station):
```
KRA_Installer/
├── install.exe
├── uninstall.exe
├── kra_checker.exe
├── heartbeat_monitor.exe
├── credentials.json
├── config.json
├── anydesk_detector.py
└── fetch_station_info.py
```

**Installed on each station:**
```
C:\Automation_and_ops_engineering\KRA_Checker\
├── kra_checker.exe
├── heartbeat_monitor.exe
├── uninstall.exe
├── run_kra_checker.vbs        ← Task Scheduler calls this (silent launch)
├── run_heartbeat.vbs          ← Task Scheduler calls this (silent launch)
├── credentials.json
├── config.json
├── anydesk_detector.py
├── fetch_station_info.py
├── .lastdate_Report           ← flag: last date separator written
├── .lastdate_Logs
└── .lastdate_Heartbeat_Error_Logs
```

---

## Silent Background Execution
Tasks run under the built-in Windows SYSTEM account. No Windows username or password is required during installation. This allows fully unattended execution even when no user is logged in and avoids password expiry or credential sync issues across stations.

Both programs run completely silently when triggered by Task Scheduler. No console window appears on screen.

**How it works:** Task Scheduler calls `wscript.exe run_heartbeat.vbs` instead of the exe directly. VBScript's `Run` with parameter `0` starts the process with no window at all — no flash, no minimize, nothing visible to station staff.

**For debugging:** Run the `.py` files directly with Python from CMD to see full console output:
```cmd
python kra_auto_checker.py
python heartbeat_monitor.py
```

---

## Installing on a Station

1. AnyDesk into the station
2. Copy the `KRA_Installer` folder to the Desktop
3. Double-click `install.exe` — UAC prompt handles elevation automatically
4. The installer will:
   - Detect the AnyDesk ID automatically (registry + config file)
   - Look up station name from Station Mapping sheet
   - If not found, prompt for name and **register it in the sheet automatically**
   - Ask for SQL Server `sa` password (only manual input required)
   - Copy all files to `C:\Automation_and_ops_engineering\KRA_Checker\`
   - Create two VBS launchers for silent execution
  - Create Task Scheduler tasks under SYSTEM account
  - Configure advanced scheduler settings automatically:
    - Run with highest privileges
    - Run on battery power
    - Do not stop on battery
    - Start missed tasks automatically
    - Ignore overlapping runs
  - Create silent VBS launchers for hidden execution
   - Mark station as installed in Station Mapping sheet

---

## Uninstalling from a Station

1. Run `C:\Automation_and_ops_engineering\KRA_Checker\uninstall.exe`
2. Type `YES` to confirm
3. Removes tasks, deletes install folder, self-deletes the exe
4. Leaves `C:\Automation_and_ops_engineering\` if other projects are present
5. Google Sheets data is untouched

---

## Pushing Updates to All Stations

### Config changes
Edit the Global Config sheet → all stations pick it up next heartbeat run.

### Exe updates
1. Build new exe
2. Upload to Google Drive — **replace existing file at same Drive ID**
3. Bump `remote_version_kra` or `remote_version_heartbeat` in Global Config sheet
4. Heartbeat downloads and hot-swaps on each station within one cycle
5. Check Station Mapping columns D-G for per-station update status

**Note:** Heartbeat manages updates for **both** exes. KRA checker is updated first (safe, not running), heartbeat updated last (causes one silent restart). Station shows `🟡 Stale` for at most one cycle during heartbeat update, then recovers.

### Credentials update
Manual — AnyDesk into each station and replace `credentials.json`.

---

## Clearing Sheets (Development)

(Development purpose) Run `clear_sheets.py` where there is a correct credentials.jsons file or key for fresh start of logs:
```cmd
python clear_sheets.py
```

This:
- Removes sheet protection(Previously protection existed in the sheet but not there anymore, it wont affect the running of this script)
- Clears all data rows (keeps headers)
- Deletes all local flag files (`.lastdate_*`) if exists
- After next program run, separators is re-applied fresh

---

## Building Exes

From the `Builddet` folder on Windows:

```cmd
pyinstaller --onefile --console --name kra_checker ^
  --hidden-import pyodbc ^
  --hidden-import googleapiclient ^
  --hidden-import googleapiclient.http ^
  --hidden-import google.auth ^
  --hidden-import google.oauth2.service_account ^
  --hidden-import config_loader ^
  --hidden-import packaging ^
  kra_auto_checker.py


pyinstaller --onefile --console --name heartbeat_monitor ^
  --hidden-import pyodbc ^
  --hidden-import googleapiclient ^
  --hidden-import googleapiclient.http ^
  --hidden-import google.auth ^
  --hidden-import google.oauth2.service_account ^
  --hidden-import config_loader ^
  --hidden-import packaging ^
  heartbeat_monitor.py

pyinstaller --onefile --console --name install --hidden-import task_scheduler --hidden-import anydesk_detector --hidden-import fetch_station_info --hidden-import config_loader install.py

pyinstaller --onefile --console --name uninstall ^
  --hidden-import task_scheduler ^
  uninstall.py
```

> Use `--console` for all builds. The programs call `ShowWindow(hwnd, 0)` on startup but Task Scheduler uses the VBS launcher which starts with no window at all. For debugging, run `.py` files directly with Python.

---

## KRA Check Result Codes

| Status | Meaning |
|---|---|
| `🟢 SUCCESS` | Transaction confirmed submitted to KRA |
| `🔴 NOT SUBMITTED` | Transaction not found on KRA portal |
| `🟡 ERROR` | Could not determine status (server error, timeout, connection reset) |
| `⚪ NO DATA` | No transactions found in SQL database for that day |

---

## Files Reference

| File                        | Purpose                                | Deploy to station?    |
| --------------------------- | -------------------------------------- | --------------------- |
| `kra_auto_checker.py`       | KRA checker source                     | No (build to exe)     |
| `heartbeat_monitor.py`      | Heartbeat source                       | No (build to exe)     |
| `config_loader.py`          | Loads + merges local/global config     | No (bundled into exe) |
| `task_scheduler.py`         | Task Scheduler helper utilities        | No (bundled into exe) |
| `anydesk_detector.py`       | Detects AnyDesk ID at install time     | Yes                   |
| `fetch_station_info.py`     | Looks up + registers station in sheet  | Yes                   |
| `install.py`                | Installer source                       | No (build to exe)     |
| `uninstall.py`              | Uninstaller source                     | No (build to exe)     |
| `clear_sheets.py`           | Dev utility — clears monitoring sheets | No                    |
| `clean_build.bat`           | Dev utility — Run once before a fresh build to remove all cached artifacts | No                    |
| `station_status_monitor.gs` | Apps Script — paste into Google Sheets | No                    |
| `config.json`               | Station config template                | Yes                   |
| `credentials.json`          | Google service account key             | Yes                   |
| `requirements.txt`          | Python dependencies                    | No                    |
| `hook-config_loader.py`     | PyInstaller hook                       | No                    |
