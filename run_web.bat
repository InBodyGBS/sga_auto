@echo off
chcp 65001 > nul
title InBody SG&A 분석 - Web UI

echo.
echo ============================================================
echo  InBody SG&A 분석 - Web UI 시작
echo ============================================================
echo.

cd /d "%~dp0"

:: uv 경로 탐색
set UV=
if exist "%USERPROFILE%\.local\bin\uv.exe" set UV=%USERPROFILE%\.local\bin\uv.exe
if exist "%APPDATA%\uv\bin\uv.exe"         set UV=%APPDATA%\uv\bin\uv.exe

where uv >nul 2>&1
if %ERRORLEVEL% == 0 set UV=uv

if "%UV%"=="" (
    echo [오류] uv 를 찾을 수 없습니다.
    echo https://docs.astral.sh/uv/getting-started/installation 에서 설치하세요.
    pause
    exit /b 1
)

:: 기존 8501 포트 프로세스 종료
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr ":8501 "') do (
    taskkill /F /PID %%a >nul 2>&1
)

echo  서버를 시작합니다. 잠시 기다려주세요...
echo.

:: 서버를 백그라운드로 실행
start "SGA_Q Server" "%UV%" run --with streamlit --with openpyxl --with pandas streamlit run "%~dp0app.py" --server.headless true --browser.gatherUsageStats false --server.port 8501

:: 서버 기동 대기 (3초)
timeout /t 3 /nobreak > nul

:: 브라우저 오픈
echo  브라우저를 엽니다: http://localhost:8501
start http://localhost:8501

echo.
echo  서버가 실행 중입니다.
echo  브라우저 주소: http://localhost:8501
echo  서버를 종료하려면 이 창을 닫으세요.
echo.

:: 서버 프로세스가 살아있는 동안 이 창 유지
:loop
timeout /t 5 /nobreak > nul
tasklist /fi "windowtitle eq SGA_Q Server" 2>nul | find "cmd.exe" >nul
if %ERRORLEVEL% == 0 goto loop

echo  서버가 종료되었습니다.
pause
