@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Smart Auditor - ?????? ?????

echo ============================================================
echo   Smart Auditor  ^|  ?????? ?????
echo   Address :  http://127.0.0.1:8000
echo   Press Ctrl+C to stop  /  ???? Ctrl+C ???????
echo ============================================================
echo.

REM --- Discover a usable Python interpreter (skip Windows App Installer shim) ---
set "PY="
python --version >nul 2>&1
if not errorlevel 1 set "PY=python"

if not defined PY (
    for /d %%D in ("%LocalAppData%\Programs\Python\Python3*") do (
        if exist "%%D\python.exe" if not defined PY set "PY=%%D\python.exe"
    )
)

if not defined PY (
    py --version >nul 2>&1
    if not errorlevel 1 set "PY=py"
)

if not defined PY (
    echo [!] Python was not found on this computer.
    echo     Please install Python 3, then run this file again.
    echo.
    pause
    exit /b 1
)

REM --- If a server is already listening on port 8000, just open it -------
netstat -ano | findstr ":8000" | findstr /i "LISTENING" >nul 2>nul
if not errorlevel 1 (
    echo [i] A server is already running on port 8000.
    echo     Opening the browser...  /  ??? ???????
    start "" http://127.0.0.1:8000
    exit /b 0
)

REM --- Open the browser once per launch (not on automatic restarts) ------
start "" http://127.0.0.1:8000

REM --- Run the server, restart automatically if it ever stops ------------
:loop
"%PY%" main.py < nul
echo.
echo [!] Server stopped / ???? ???????. Restarting in 3s... (Ctrl+C to exit)
timeout /t 3 /nobreak >nul
goto loop