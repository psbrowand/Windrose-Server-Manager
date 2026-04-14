@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0Windrose-Server-Manager.ps1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: The manager failed to launch. See details above.
    pause
)
