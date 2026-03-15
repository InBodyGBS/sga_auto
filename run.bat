@echo off
chcp 65001 > nul
title InBody SG&A 분기별 분석

echo.
echo ============================================================
echo  InBody SG&A 분기별 분석 시작
echo ============================================================
echo.

:: Python 경로 자동 탐색
where python >nul 2>&1
if %ERRORLEVEL% == 0 (
    set PYTHON=python
    goto :run
)

where py >nul 2>&1
if %ERRORLEVEL% == 0 (
    set PYTHON=py
    goto :run
)

:: 가상환경 탐색 (현재 폴더 기준)
if exist ".venv\Scripts\python.exe" (
    set PYTHON=.venv\Scripts\python.exe
    goto :run
)

if exist "venv\Scripts\python.exe" (
    set PYTHON=venv\Scripts\python.exe
    goto :run
)

echo [오류] Python을 찾을 수 없습니다.
echo 다음 중 하나를 확인하세요:
echo   1. Python이 설치되어 있고 PATH에 등록되어 있는지
echo   2. 이 폴더에 .venv 또는 venv 가상환경이 있는지
echo.
pause
exit /b 1

:run
cd /d "%~dp0"
%PYTHON% main.py

:: main.py 내부에서 pause 처리하므로 여기서는 생략
