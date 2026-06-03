@echo off
setlocal
echo.
echo  正在卸载 Office Agent Local...
echo.

REM 停止服务
taskkill /F /IM node.exe /FI "WINDOWTITLE eq Office Agent Local*" >nul 2>&1

REM 删除 WEF 注册 (通过 PowerShell 获取 SID)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$sid = (New-Object System.Security.Principal.NTAccount('%USERNAME%')).Translate([System.Security.Principal.SecurityIdentifier]).Value; " ^
  "$dir = Join-Path (Join-Path (Join-Path (Join-Path $env:LOCALAPPDATA 'Microsoft\Office') '16.0\Wef') $sid) 'office-agent-local'; " ^
  "if (Test-Path $dir) { Remove-Item -Recurse -Force $dir; Write-Host '  Removed:' $dir } else { Write-Host '  Not found (already removed?)' }"

echo.
echo  卸载完成。如需删除程序文件，请手动删除整个文件夹。
echo.
pause
