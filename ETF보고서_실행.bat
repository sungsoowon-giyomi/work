@echo off
chcp 65001 > nul
title ETF 보고서 생성기
cd /d "%~dp0"
echo ======================================
echo   ETF 보고서 생성기
echo ======================================
echo.
echo config.txt 를 확인하세요.
echo 한국ETF 섹션에 코드 입력, 미국ETF 섹션에 티커 입력
echo.
echo 시작하려면 Enter...
pause > nul
python run_report.py
if errorlevel 1 echo 오류가 발생했습니다. 위 메시지를 확인하세요.
if errorlevel 1 pause
