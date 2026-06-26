@echo off
setlocal
chcp 65001 >nul 2>&1
title Office Agent Local
set "PORT=8787"
set "URL=http://127.0.0.1:%PORT%"

REM ── 打印 Logo / Banner ───────────────────────
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0logo.ps1" 2>nul
if errorlevel 1 (
    echo  Office Agent Local v0.1.0
    echo.
)

REM ── 检查端口是否已被占用 ──────────────────────
echo  Checking port %PORT%...

REM 先看是否是自己的服务已在运行
curl -s -o NUL -w "%%{http_code}" %URL%/health >"%TEMP%\oa-health.txt" 2>nul
set /p HEALTH=<"%TEMP%\oa-health.txt"
del "%TEMP%\oa-health.txt" >nul 2>&1

if "%HEALTH%"=="200" (
    echo  [OK] Service already running at %URL%
    echo  Opening Word add-in...
    goto :ready
)

REM 端口被其他程序占用
netstat -ano | findstr ":%PORT% " | findstr "LISTENING" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo  [!] Port %PORT% is occupied by another process.
    echo.
    for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":%PORT% " ^| findstr "LISTENING"') do (
        echo      PID: %%p
        for /f "tokens=1" %%n in ('tasklist /FI "PID eq %%p" /NH ^| findstr /V "^$"') do (
            echo      Process: %%n
        )
    )
    echo.
    choice /M "  Kill the process and free port %PORT%"
    if errorlevel 2 (
        echo  Aborted. Please free port %PORT% manually.
        pause
        exit /b 1
    )
    for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":%PORT% " ^| findstr "LISTENING"') do (
        taskkill /F /PID %%p >nul 2>&1
    )
    timeout /t 1 /nobreak >nul
    echo  Port freed.
)

REM ── 启动服务 ──────────────────────────────────
echo.
echo  Starting service at %URL%
echo  Press Ctrl+C to stop.
echo.
cd /d "%~dp0server"
node server.js
goto :eof

:ready
echo.
echo  Service is ready at %URL%
echo  Open Word and click the Word Agent button.
echo.
pause
