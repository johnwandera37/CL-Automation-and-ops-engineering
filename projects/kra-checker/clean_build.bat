@echo off
REM clean_build.bat
REM Run once before a fresh build to remove all cached artifacts

echo.
echo ========================================
echo   Clean Build - Removing Artifacts
echo ========================================
echo.

echo [1/3] Removing dist and build folders...
if exist dist        ( rd /s /q dist        && echo   Removed: dist        ) else ( echo   Not found: dist )
if exist build       ( rd /s /q build       && echo   Removed: build       ) else ( echo   Not found: build )

echo.
echo [2/3] Removing spec files...
del /f /q *.spec 2>nul && echo   Removed: *.spec

echo.
echo [3/3] Removing Python cache...
if exist __pycache__ ( rd /s /q __pycache__ && echo   Removed: __pycache__ ) else ( echo   Not found: __pycache__ )
for /r . %%F in (*.pyc) do ( del /f /q "%%F" )

echo.
echo ========================================
echo   Done. Ready for a clean build.
echo ========================================
echo.