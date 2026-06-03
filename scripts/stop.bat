@echo off
setlocal
set "PORT=8787"

echo.
echo  Stopping Office Agent Local...
echo.

REM Check if our service is running
curl -s -o NUL -w "%%{http_code}" http://127.0.0.1:%PORT%/health >"%TEMP%\oa-check.txt" 2>nul
set /p STATUS=<"%TEMP%\oa-check.txt"
del "%TEMP%\oa-check.txt" >nul 2>&1

if not "%STATUS%"=="200" (
    echo  Service is not running on port %PORT%.
    pause
    exit /b 0
)

REM Find and kill the process on our port
for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":%PORT% " ^| findstr "LISTENING"') do (
    echo  Stopping PID %%p ...
    taskkill /F /PID %%p >nul 2>&1
)

timeout /t 1 /nobreak >nul

REM Verify stopped
curl -s -o NUL -w "%%{http_code}" http://127.0.0.1:%PORT%/health >"%TEMP%\oa-check.txt" 2>nul
set /p STATUS=<"%TEMP%\oa-check.txt"
del "%TEMP%\oa-check.txt" >nul 2>&1

if "%STATUS%"=="200" (
    echo  [!] Failed to stop. Try running as administrator.
) else (
    echo  [OK] Service stopped.
)

echo.
pause
