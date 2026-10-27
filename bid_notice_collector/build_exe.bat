@echo off
REM 파일 역할: Windows에서 의존성을 설치하고 Tkinter 앱을 단일 EXE로 빌드합니다.
REM 마지막 업데이트: 2026-10-23
REM 작업자: SMI 장지수

setlocal
cd /d "%~dp0"
python -m pip install -r requirements-build.txt
if errorlevel 1 exit /b 1
python -m PyInstaller --noconfirm --clean --onefile --windowed --name BidNoticeCollector desktop_app.py
if errorlevel 1 exit /b 1
echo EXE created: dist\BidNoticeCollector.exe
pause
