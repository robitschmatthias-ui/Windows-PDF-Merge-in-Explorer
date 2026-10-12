@echo off
REM PDF Merge in Explorer - starts the uninstaller
setlocal
set "SCRIPT_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%uninstall.ps1"
echo.
pause
