@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements-build.txt
if errorlevel 1 exit /b 1
python -m PyInstaller --noconfirm --clean --onefile --windowed --name BidNoticeCollector desktop_app.py
if errorlevel 1 exit /b 1
echo EXE created: dist\BidNoticeCollector.exe
pause
