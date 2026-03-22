@echo off
setlocal EnableDelayedExpansion
title ETIMS Backup Uploader - Installation

echo.
echo  ============================================================
echo    ETIMS Backup Uploader  ^|  Installation
echo  ============================================================
echo.

REM ── Check required files ─────────────────────────────────────────────────
if not exist "%~dp0credentials.json" ( echo  [ERROR] credentials.json not found. & pause & exit /b 1 )
if not exist "%~dp0main.exe"         ( echo  [ERROR] main.exe not found.          & pause & exit /b 1 )
if not exist "%~dp0config.json"      ( echo  [ERROR] config.json not found.       & pause & exit /b 1 )

findstr /C:"REPLACE_WITH" "%~dp0config.json" >nul
if %errorlevel% equ 0 (
    echo  [ERROR] config.json still has placeholder values.
    echo          Fill in drive_root_folder_id and log_sheet_id before deploying.
    pause & exit /b 1
)

REM ── Station name ──────────────────────────────────────────────────────────
echo  Step 1 of 2 - Station
echo.
echo  Enter station type then name. First word = type, rest = name.
echo  Examples:  Rubis Eldoret Town   /   Shell Nakuru West
echo.
set /p STATION_FULL="  Enter station: "
if "!STATION_FULL!"=="" ( echo  Cannot be empty. & pause & exit /b 1 )

for /f "tokens=1,*" %%A in ("!STATION_FULL!") do (
    set STATION_TYPE=%%A
    set STATION_NAME=%%B
)
if "!STATION_NAME!"=="" (
    echo  [ERROR] Enter both type and name e.g. Rubis Eldoret Town
    pause & exit /b 1
)

echo.
echo  Station type : !STATION_TYPE!
echo  Station name : !STATION_NAME!
echo.
set /p CONFIRM="  Is this correct? (Y/N): "
if /i not "!CONFIRM!"=="Y" ( echo  Cancelled. & pause & exit /b 1 )

REM ── Auto-detect backup path ───────────────────────────────────────────────
echo.
echo  Step 2 of 2 - Backup Folder
echo.
echo  Searching for ETIMS backup folder...

set BACKUP_PATH=
for /f "delims=" %%A in ('"%~dp0main.exe" --detect-backup 2^>nul') do set DETECT_RESULT=%%A

REM Check if result starts with FOUND:
echo !DETECT_RESULT! | findstr /C:"FOUND:" >nul
if %errorlevel% equ 0 (
    REM Strip the FOUND: prefix
    set BACKUP_PATH=!DETECT_RESULT:FOUND:=!
    echo.
    echo  Backup folder found:
    echo    !BACKUP_PATH!
    echo.
    set /p CONFIRM2="  Use this path? (Y/N): "
    if /i not "!CONFIRM2!"=="Y" set BACKUP_PATH=
)

REM If not found or user said No, ask manually
if "!BACKUP_PATH!"=="" (
    echo.
    echo  Could not find backup folder automatically.
    echo  Please enter the full path to the backup folder:
    echo  Example: C:\Users\Administrator\Documents\ETIMS\SQLBackupAndFTP\Backup
    echo.
    set /p BACKUP_PATH="  Backup path: "
    if "!BACKUP_PATH!"=="" (
        echo  [ERROR] Backup path cannot be empty.
        pause & exit /b 1
    )
)

echo.
echo  Using backup path: !BACKUP_PATH!

REM ── Schedule time ─────────────────────────────────────────────────────────
echo.
echo  Optional - Daily run time (press Enter for default 01:00)
set /p SCHEDULE_TIME="  Enter time 24h format e.g. 08:50 [01:00]: "
if "!SCHEDULE_TIME!"=="" set SCHEDULE_TIME=01:00

REM ── Copy files ────────────────────────────────────────────────────────────
set INSTALL_DIR=C:\ProgramData\ETIMSBackup
echo.
echo  Installing to %INSTALL_DIR% ...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

copy /Y "%~dp0main.exe"          "%INSTALL_DIR%\main.exe"          >nul
copy /Y "%~dp0credentials.json"  "%INSTALL_DIR%\credentials.json"  >nul
copy /Y "%~dp0register_task.ps1" "%INSTALL_DIR%\register_task.ps1" >nul
copy /Y "%~dp0uninstall.bat"     "%INSTALL_DIR%\uninstall.bat"     >nul
echo  Files copied.

REM ── Read IDs from source config.json ─────────────────────────────────────
set DRIVE_FOLDER_ID=
set SHEET_ID=
set SA_EMAIL=kra-checker-service@kra-auto-checker.iam.gserviceaccount.com

for /f "tokens=2 delims=:, " %%A in ('findstr "drive_root_folder_id" "%~dp0config.json"') do set DRIVE_FOLDER_ID=%%~A
for /f "tokens=2 delims=:, " %%A in ('findstr "log_sheet_id" "%~dp0config.json"') do set SHEET_ID=%%~A

REM ── Write config.json ─────────────────────────────────────────────────────
echo  Writing config.json...
PowerShell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$config = [ordered]@{service_account_file='credentials.json'; service_account_email='!SA_EMAIL!'; station_type='!STATION_TYPE!'; station_name='!STATION_NAME!'; backup_path='!BACKUP_PATH!'; drive_root_folder_id='!DRIVE_FOLDER_ID!'; log_sheet_id='!SHEET_ID!'; schedule_time='!SCHEDULE_TIME!'; max_retries=3; retry_interval_seconds=300; max_wait_minutes=120; log_spreadsheet_name='ETIMSBackupLog'}; $json = $config | ConvertTo-Json; [System.IO.File]::WriteAllText('%INSTALL_DIR%\config.json', $json, [System.Text.UTF8Encoding]::new($false))"
echo  config.json saved.

REM ── Register scheduled task ───────────────────────────────────────────────
echo.
echo  Registering scheduled task (UAC prompt may appear - click Yes)...
PowerShell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%INSTALL_DIR%\register_task.ps1""' -Verb RunAs -Wait"

REM ── Connection test ───────────────────────────────────────────────────────
echo.
echo  Testing connection to Google Drive...
"%INSTALL_DIR%\main.exe" --test 2>&1
if %errorlevel% equ 0 (
    echo  [OK] Connected to Google Drive successfully.
) else (
    echo  [WARNING] Connection test failed. See %INSTALL_DIR%\etims_backup.log
)

REM ── Done ─────────────────────────────────────────────────────────────────
echo.
echo  ============================================================
echo    Installation complete!
echo.
echo    Station  : !STATION_TYPE! - !STATION_NAME!
echo    Backup   : !BACKUP_PATH!
echo    Schedule : Daily at !SCHEDULE_TIME!
echo    Location : %INSTALL_DIR%
echo    Log file : %INSTALL_DIR%\etims_backup.log
echo.
echo    To uninstall: run %INSTALL_DIR%\uninstall.bat
echo  ============================================================
echo.
pause
endlocal
