@echo off
cd /d "%~dp0"

REM Unblock downloaded files automatically (no Properties dialog needed).
powershell.exe -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%~dp0' -File | Unblock-File -ErrorAction SilentlyContinue" >nul 2>&1

REM STA is required for the folder-picker window.
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -Sta -File "%~dp0pub2pdf-gui.ps1"

if errorlevel 1 (
    echo.
    echo Something went wrong starting the converter.
    echo Press any key to close this window.
    pause >nul
)
