@echo off
setlocal
set "XVIEWER_PYTHON=%LOCALAPPDATA%\xViewer\venv\Scripts\python.exe"
if not exist "%XVIEWER_PYTHON%" (
    echo Please run install_xViewer.bat first.
    pause
    exit /b 1
)
"%XVIEWER_PYTHON%" "%~dp0tools\diagnose_excel.py" %*
set "XVIEWER_RESULT=%ERRORLEVEL%"
pause
exit /b %XVIEWER_RESULT%
