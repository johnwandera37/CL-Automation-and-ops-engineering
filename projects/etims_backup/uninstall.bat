@echo off
setlocal
title ETIMS Backup Uploader - Uninstall

echo.
echo  ============================================================
echo    ETIMS Backup Uploader  ^|  Uninstall
echo  ============================================================
echo.
echo  This will:
echo    - Remove the scheduled task
echo    - Delete all files from C:\ProgramData\ETIMSBackup
echo.
set /p CONFIRM="  Are you sure? (Y/N): "
if /i not "%CONFIRM%"=="Y" (
    echo  Cancelled.
    pause & exit /b 0
)

REM ── Remove scheduled task ────────────────────────────────────────────────
echo.
echo  Removing scheduled task...
PowerShell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -Command ""Unregister-ScheduledTask -TaskName ETIMSBackupUploader -Confirm:$false""' -Verb RunAs -Wait"

REM ── Delete install folder ────────────────────────────────────────────────
echo  Removing C:\ProgramData\ETIMSBackup ...
rd /s /q "C:\ProgramData\ETIMSBackup" 2>nul

echo.
echo  ============================================================
echo    Uninstall complete.
echo  ============================================================
echo.
pause
endlocal
