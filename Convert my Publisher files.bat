@echo off
cd /d "%~dp0"

REM === EDIT THESE TWO FOLDERS ===
set "SOURCE=PUT YOUR PUBLISHER FOLDER HERE"
set "OUTPUT=PUT YOUR PDF FOLDER HERE"

REM Optional: seconds to wait for export (600 = 10 minutes). Remove this switch to use the default (180).
set "TIMEOUT=600"

powershell.exe -ExecutionPolicy Bypass -File "%~dp0pub2pdf.ps1" -SourceRoot "%SOURCE%" -OutputRoot "%OUTPUT%" -Skip -ExportTimeoutSeconds %TIMEOUT%

echo.
echo Finished. Press any key to close this window.
pause >nul
