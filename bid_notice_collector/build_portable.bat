@echo off
REM 파일 역할: Python 미설치 PC용 포터블 배포본(dist\BidNoticeCollector_portable.zip)을 만듭니다.
REM 마지막 업데이트: 2026-11-03
REM 작업자: SMI 장지수

setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_portable.ps1"
if errorlevel 1 exit /b 1
pause
