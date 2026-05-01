@echo off
cd /d E:\Sauvegarde\draft club\mal1-fantasy
taskkill /F /IM pythonw.exe >nul 2>&1
taskkill /F /IM python.exe >nul 2>&1
timeout /t 1 /nobreak >nul
start "" cmd /c "python -u scripts\admin_server.py > admin_log.txt 2>&1"
timeout /t 1 /nobreak >nul
start http://localhost:8765
exit
