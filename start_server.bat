@echo off
REM ============================================================
REM  AdrUs-Engine - Random US Address Scraper launcher
REM  Activates the local venv and runs the scraper forever
REM  (Ctrl+C to stop). Any extra args are passed through, e.g.
REM    start_server.bat --delay 3
REM ============================================================

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] Virtual environment not found at .venv
    echo Create it first:  python -m venv .venv ^&^& .venv\Scripts\pip install -r requirements.txt
    pause
    exit /b 1
)

echo Starting address scraper... (press Ctrl+C to stop)
".venv\Scripts\python.exe" scrape_addresses.py %*

pause
