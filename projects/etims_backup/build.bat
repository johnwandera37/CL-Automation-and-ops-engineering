@echo off
REM build.bat  -  Run this on Windows to produce main.exe
REM Requires: Python 3.11+  and  pip install -r requirements.txt

echo Installing dependencies...
pip install -r requirements.txt

echo.
echo Building exe with PyInstaller...
pyinstaller ^
    --onefile ^
    --name main ^
    --hidden-import google.auth.transport.requests ^
    --hidden-import google.oauth2.service_account ^
    --hidden-import googleapiclient.discovery ^
    --hidden-import openpyxl ^
    main.py

echo.
echo Done. Executable is in .\dist\main.exe
echo.
echo === Deployment checklist ===
echo   Prepare a folder with these 6 files:
echo     1. dist\main.exe
echo     2. credentials.json
echo     3. config.json          (drive_root_folder_id and log_sheet_id must be filled in)
echo     4. install.bat
echo     5. uninstall.bat
echo     6. register_task.ps1
echo.
echo   Then on each station:
echo     - Double-click install.bat
echo     - Enter station name e.g. "Rubis Eldoret Town"
echo     - Confirm backup path
echo     - Set schedule time (default 01:00)
echo     - Done. install.bat handles everything else automatically.
pause
