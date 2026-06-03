@echo off
setlocal
chcp 65001 >nul 2>&1
echo.
echo  '╔══════════════════════════════════════╗'
echo  '║  Office Agent Local ─ 安装向导       ║'
echo  '╚══════════════════════════════════════╝'
echo.

REM ── 检查 Node.js ──────────────────────────────
where node >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [错误] 未检测到 Node.js，请先安装 Node.js ^(https://nodejs.org^)
    echo        建议版本: 18.x 或更高
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('node -v') do echo  Node.js 版本: %%v

REM ── 注册 Add-in Manifest (不需要启动服务) ──────
echo.
echo  [1/2] 正在注册 Word Add-in...

REM 使用 PowerShell 获取用户 SID 并写入 WEF 目录
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$sid = (New-Object System.Security.Principal.NTAccount('%USERNAME%')).Translate([System.Security.Principal.SecurityIdentifier]).Value; " ^
  "$dir = Join-Path (Join-Path (Join-Path (Join-Path $env:LOCALAPPDATA 'Microsoft\Office') '16.0\Wef') $sid) 'office-agent-local\Manifests'; " ^
  "New-Item -ItemType Directory -Path $dir -Force | Out-Null; " ^
  "Copy-Item '%~dp0manifest.xml' (Join-Path $dir 'manifest.xml') -Force; " ^
  "Write-Host '  Manifest registered at:' $dir"

if %ERRORLEVEL% neq 0 (
    echo  [警告] 自动注册失败，请手动注册:
    echo         Word ^> 文件 ^> 选项 ^> 信任中心 ^> 受信任的加载项目录
    echo         添加路径: %~dp0
)

REM ── 完成 ──────────────────────────────────────
echo.
echo  [2/2] 安装完成!
echo.
echo  ──────────────────────────────────────
echo   接下来请:
echo   1. 双击 start.bat 启动本地服务
echo   2. 打开 Microsoft Word
echo   3. 点击「开始」选项卡中的「Word Agent」按钮
echo   4. 首次使用需在侧边栏配置 API 地址和密钥
echo  ──────────────────────────────────────
echo.
echo  启动服务: 双击 start.bat
echo  卸载:     双击 uninstall.bat
echo.
pause
