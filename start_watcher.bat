@echo off
cd /d "C:\Users\GoldenTech\inventory-nour"
python3 inventory_watcher.py --once
exit /b %ERRORLEVEL%
