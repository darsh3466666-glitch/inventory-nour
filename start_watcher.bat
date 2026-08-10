@echo off
REM ============================================================
REM  Inventory Watcher Launcher - النور لإدارة المخزون
REM  Runs the Python watcher script that updates inventory JSON
REM  Called by Windows Task Scheduler every 1 minute
REM ============================================================

cd /d "C:\Users\GoldenTech\.openclaw-autoclaw\workspace\.cluster\DELIVERY"

REM Run the watcher in --once mode (single update, then exit)
REM Task Scheduler will call this script every minute
python3 inventory_watcher.py --once

REM Exit code propagated to Task Scheduler
exit /b %ERRORLEVEL%
