@echo off
cd /d "C:\Users\GoldenTech\inventory-nour"
"C:\Python314\python.exe" inventory_watcher.py --once
exit /b %ERRORLEVEL%
