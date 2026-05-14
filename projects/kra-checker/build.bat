@echo off
title KRA Checker Build System
color 0A

echo.
echo ====================================================
echo   KRA CHECKER - BUILD SYSTEM
echo ====================================================
echo.

REM =====================================================
REM Clean previous builds
REM =====================================================

echo [1/5] Cleaning old build files...

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

del /q *.spec >nul 2>&1

echo Done.
echo.

REM =====================================================
REM Build install.exe
REM =====================================================

echo [2/5] Building install.exe...

pyinstaller --onefile --console ^
--name install ^
--hidden-import task_scheduler ^
install.py

echo.
echo install.exe complete.
echo.

REM =====================================================
REM Build uninstall.exe
REM =====================================================

echo [3/5] Building uninstall.exe...

pyinstaller --onefile --console ^
--name uninstall ^
uninstall.py

echo.
echo uninstall.exe complete.
echo.

REM =====================================================
REM Build heartbeat_monitor.exe
REM =====================================================

echo [4/5] Building heartbeat_monitor.exe...

pyinstaller --onefile --console ^
--name heartbeat_monitor ^
--hidden-import pyodbc ^
--hidden-import googleapiclient ^
--hidden-import googleapiclient.http ^
--hidden-import google.auth ^
--hidden-import google.oauth2.service_account ^
--hidden-import config_loader ^
--hidden-import task_scheduler ^
--hidden-import packaging ^
heartbeat_monitor.py

echo.
echo heartbeat_monitor.exe complete.
echo.

REM =====================================================
REM Build kra_checker.exe
REM =====================================================

echo [5/5] Building kra_checker.exe...

pyinstaller --onefile --console ^
--name kra_checker ^
--hidden-import pyodbc ^
--hidden-import googleapiclient ^
--hidden-import googleapiclient.http ^
--hidden-import google.auth ^
--hidden-import google.oauth2.service_account ^
--hidden-import config_loader ^
--hidden-import task_scheduler ^
--hidden-import packaging ^
kra_auto_checker.py

echo.
echo kra_checker.exe complete.
echo.

echo ====================================================
echo   BUILD FINISHED
echo ====================================================
echo.

echo Executables available in:
echo dist
echo.

pause
