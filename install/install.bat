@echo off
REM PDF Merge in Explorer - starts the installer (bypasses the PowerShell ExecutionPolicy)
setlocal
set "SCRIPT_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%install.ps1"
echo.
pause
