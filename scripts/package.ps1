# Office Agent Local - Package Script
# Usage: npm run package

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$distDir = Join-Path $root "dist"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Office Agent Local - Packager" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Clean dist
Write-Host "[1/5] Cleaning dist..." -ForegroundColor Yellow
if (Test-Path $distDir) { Remove-Item -Recurse -Force $distDir }
New-Item -ItemType Directory -Path $distDir | Out-Null

# Step 2: Build word-addin (production)
Write-Host "[2/5] Building Word Add-in (production)..." -ForegroundColor Yellow
Push-Location (Join-Path $root "apps\word-addin")
npm run build:prod
if ($LASTEXITCODE -ne 0) { Write-Error "Word Add-in build failed"; exit 1 }
Pop-Location

# Step 3: Build local-agent
Write-Host "[3/5] Building Local Agent..." -ForegroundColor Yellow
Push-Location (Join-Path $root "apps\local-agent")
npm run build
if ($LASTEXITCODE -ne 0) { Write-Error "Local Agent build failed"; exit 1 }
Pop-Location

# Step 4: Assemble dist folder
Write-Host "[4/5] Assembling distribution..." -ForegroundColor Yellow

# Copy local-agent dist
Copy-Item -Recurse (Join-Path $root "apps\local-agent\dist") (Join-Path $distDir "server")
Copy-Item (Join-Path $root "apps\local-agent\package.json") (Join-Path $distDir "server\package.json")

# Copy built word-addin into server/public
$publicDir = Join-Path $distDir "server\public"
New-Item -ItemType Directory -Path $publicDir | Out-Null
Copy-Item -Recurse (Join-Path $root "apps\word-addin\dist\*") $publicDir

# Copy icons (if exist)
$iconDir = Join-Path $root "apps\word-addin\public"
if (Test-Path $iconDir) {
    Get-ChildItem $iconDir -Filter "*.png" | ForEach-Object {
        Copy-Item $_.FullName $publicDir
    }
}

# Copy manifest
Copy-Item (Join-Path $root "scripts\manifest-production.xml") (Join-Path $distDir "manifest.xml")

# Copy scripts
Copy-Item (Join-Path $root "scripts\install.bat") (Join-Path $distDir "install.bat")
Copy-Item (Join-Path $root "scripts\start.bat") (Join-Path $distDir "start.bat")
Copy-Item (Join-Path $root "scripts\stop.bat") (Join-Path $distDir "stop.bat")
Copy-Item (Join-Path $root "scripts\uninstall.bat") (Join-Path $distDir "uninstall.bat")
Copy-Item (Join-Path $root "scripts\logo.ps1") (Join-Path $distDir "logo.ps1")
# logo.ps1 contains box-drawing chars; PS 5.1 needs UTF-8 BOM to read them
$logoDst = Join-Path $distDir "logo.ps1"
$logoText = [System.IO.File]::ReadAllText($logoDst, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($logoDst, $logoText, (New-Object System.Text.UTF8Encoding($true)))

# Copy README
Copy-Item (Join-Path $root "scripts\README-dist.md") (Join-Path $distDir "README.md")

# Step 5: Install production dependencies
Write-Host "[5/5] Installing production dependencies..." -ForegroundColor Yellow
Push-Location (Join-Path $distDir "server")
npm install --omit=dev --production
if ($LASTEXITCODE -ne 0) {
    Write-Warning "Dependency install may be incomplete"
}
Pop-Location

# Summary
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Package complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Output: $distDir" -ForegroundColor Cyan
Write-Host ""
Write-Host "Structure:" -ForegroundColor Cyan
Write-Host "  dist/"
Write-Host "  +-- manifest.xml"
Write-Host "  +-- install.bat"
Write-Host "  +-- start.bat"
Write-Host "  +-- uninstall.bat"
Write-Host "  +-- README.md"
Write-Host "  +-- server/"
Write-Host "      +-- server.js"
Write-Host "      +-- package.json"
Write-Host "      +-- node_modules/"
Write-Host "      +-- public/"
Write-Host ""
Write-Host "Distribute by zipping the dist/ folder." -ForegroundColor Yellow
