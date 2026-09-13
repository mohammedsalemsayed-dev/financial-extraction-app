@echo off
cd /d "%~dp0"
echo Starting the extraction app...
echo A browser tab will open. Close this window (or press Ctrl+C) to stop.
echo.
python -m financial_extract
pause
