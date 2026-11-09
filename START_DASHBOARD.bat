@echo off
cd /d "%~dp0"
py -m pip show openpyxl >nul 2>&1
if errorlevel 1 (
  echo Dang cai thu vien openpyxl...
  py -m pip install openpyxl
)
py app.py
pause
