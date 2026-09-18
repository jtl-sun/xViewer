@echo off
setlocal
cd /d "%~dp0"
title xViewer Installer

echo.
echo ============================================================
echo            xViewer Installer
echo ============================================================
echo.

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_windows.ps1"
set "RC=%ERRORLEVEL%"

echo.
if not "%RC%"=="0" (
    echo Installation failed. The error is shown above.
) else (
    echo Installation complete.
    echo Use the xViewer icon on your Desktop.
)
echo.
pause
exit /b %RC%
