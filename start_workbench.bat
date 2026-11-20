@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

rem Standalone launcher for the IAV TV workbench.
rem Unlike run.bat this skips the menu and starts the API server directly.
rem Runs under the windowless pythonw so a logon-triggered start does not leave
rem a console window behind that can be closed by accident.
rem
rem Two log files, kept apart on purpose:
rem   output\workbench_launcher.log - when the launcher started/skipped/exited
rem   output\workbench_server.log  - the server's own stdout/stderr (tracebacks)
rem A running server holds the server log open, so launcher events must not be
rem written there or a duplicate-start attempt fails with "file is in use".

if not exist "%~dp0output" mkdir "%~dp0output"
set "LOG_LAUNCH=%~dp0output\workbench_launcher.log"
set "LOG_OUT=%~dp0output\workbench_server.log"
set "LOG_ERR=%~dp0output\workbench_server.err.log"

rem Do nothing when an instance already owns the port (manual start, repeated
rem logon trigger, ...). Two servers on 8765 would only make the second fail.
netstat -ano | findstr /R /C:":8765 .*LISTENING" >nul 2>&1
if not errorlevel 1 (
    echo ==== skipped %DATE% %TIME% : port 8765 already listening ====>> "%LOG_LAUNCH%"
    exit /b 0
)

rem Pick an interpreter that actually has the workbench dependencies.
set "PYTHON_RUNNER=pyw -3.12"
py -3.12 -c "import flask, openpyxl, pdfplumber, pypdfium2, PIL" >nul 2>&1
if errorlevel 1 (
    if exist "%~dp0.paddle-venv\Scripts\pythonw.exe" (
        set "PYTHON_RUNNER=%~dp0.paddle-venv\Scripts\pythonw.exe"
    ) else (
        set "PYTHON_RUNNER=python"
    )
)

echo ==== started %DATE% %TIME% via %PYTHON_RUNNER% ====>> "%LOG_LAUNCH%"
%PYTHON_RUNNER% api_server.py --port 8765>> "%LOG_OUT%" 2>> "%LOG_ERR%"
echo ==== exited %DATE% %TIME% (code %ERRORLEVEL%) ====>> "%LOG_LAUNCH%"
