@echo off
REM 파일 역할: Windows에서 의존성을 설치하고 Tkinter 앱을 단일 EXE로 빌드합니다.
REM 마지막 업데이트: 2026-11-09
REM 작업자: SMI 장지수

setlocal
cd /d "%~dp0"
python -m pip install -r requirements-build.txt
if errorlevel 1 exit /b 1
python -m PyInstaller --noconfirm --clean --onefile --windowed --name BidNoticeCollector desktop_app.py
if errorlevel 1 exit /b 1
REM exe는 옆에 있는 settings.json을 읽으므로 최신 설정을 함께 복사합니다.
REM 주의: exe 화면에서 저장한 dist\settings.json의 키워드는 프로젝트 설정으로 덮어써집니다.
copy /Y "settings.json" "dist\settings.json" >nul
if errorlevel 1 exit /b 1
echo EXE created: dist\BidNoticeCollector.exe
echo Settings copied: dist\settings.json
pause
