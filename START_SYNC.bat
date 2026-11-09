@echo off
cd /d "%~dp0"
py sync_to_cloud.py
if errorlevel 1 python sync_to_cloud.py
pause
