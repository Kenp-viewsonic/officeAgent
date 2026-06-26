# Office Agent - ASCII Logo & Banner
# Auto-invoked by start.bat. Prints a true-color (24-bit) gradient banner
# plus version, build date, host and endpoint info.
#
# ANSI 24-bit colors (ESC[38;2;R;G;Bm) require Windows 10+ with a VT-enabled
# host (Windows Terminal, modern conhost). On legacy hosts the escape codes
# are ignored and the banner falls back to default foreground color.

# ---------------------------------------------------------------------------
# Encoding: ensure box-drawing / ASCII art renders correctly regardless of
# the system codepage or whether start.bat was run from cmd / explorer.
#   - chcp 65001 puts the console into UTF-8 (start.bat already does this,
#     but we redo it here for the case when logo.ps1 is invoked directly).
#   - [Console]::OutputEncoding = UTF8 makes PowerShell emit UTF-8 bytes
#     for Write-Host (Windows PowerShell 5.1 defaults to the OEM codepage,
#     e.g. gb2312 on zh-CN, which collides with chcp 65001 and produces
#     garbled characters like 鈻堚枅).
# ---------------------------------------------------------------------------
try { chcp 65001 | Out-Null } catch {}
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
try { $OutputEncoding          = [System.Text.Encoding]::UTF8 } catch {}

$ErrorActionPreference = 'SilentlyContinue'

# ---------------------------------------------------------------------------
# Resolve package.json (works for both source scripts/ and packaged dist/)
# ---------------------------------------------------------------------------
$pkgPath = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\package.json'))
$version = '1.2.0'
$name    = 'Office Agent'
if (Test-Path $pkgPath) {
    try {
        $pkg = Get-Content -Raw $pkgPath | ConvertFrom-Json
        if ($pkg.version) { $version = $pkg.version }
        if ($pkg.name)    { $name    = $pkg.name }
    } catch {}
}

$buildDate = Get-Date -Format 'yyyy-MM-dd'
$hostName  = $env:COMPUTERNAME
$userName  = $env:USERNAME
$pid       = $PID
$endpoint  = 'http://127.0.0.1:8787'

# ---------------------------------------------------------------------------
# ASCII art (figlet "ANSI Shadow", two-word layout: OFFICE  AGENT)
# ---------------------------------------------------------------------------
$banner = @'
 ██████╗ ███████╗███████╗██╗ ██████╗███████╗     █████╗  ██████╗ ███████╗███╗   ██╗████████╗
██╔═══██╗██╔════╝██╔════╝██║██╔════╝██╔════╝    ██╔══██╗██╔════╝ ██╔════╝████╗  ██║╚══██╔══╝
██║   ██║█████╗  █████╗  ██║██║     █████╗      ███████║██║  ███╗█████╗  ██╔██╗ ██║   ██║
██║   ██║██╔══╝  ██╔══╝  ██║██║     ██╔══╝      ██╔══██║██║   ██║██╔══╝  ██║╚██╗██║   ██║
╚██████╔╝██║     ██║     ██║╚██████╗███████╗    ██║  ██║╚██████╔╝███████╗██║ ╚████║   ██║
 ╚═════╝ ╚═╝     ╚═╝     ╚═╝ ╚═════╝╚══════╝    ╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═╝  ╚═══╝   ╚═╝
'@

# ---------------------------------------------------------------------------
# 4-stop horizontal gradient: cyan -> blue -> violet -> pink
# ---------------------------------------------------------------------------
$stops = @(
    @{ t = 0.00; r =   6; g = 182; b = 212 }   # #06B6D4  cyan-500
    @{ t = 0.33; r =  59; g = 130; b = 246 }   # #3B82F6  blue-500
    @{ t = 0.66; r = 139; g =  92; b = 246 }   # #8B5CF6  violet-500
    @{ t = 1.00; r = 236; g =  72; b = 153 }   # #EC4899  pink-500
)

function Get-GradientRGB([double]$t) {
    if ($t -le 0) { $s = $stops[0]  ; return $s.r, $s.g, $s.b }
    if ($t -ge 1) { $s = $stops[-1] ; return $s.r, $s.g, $s.b }
    for ($i = 0; $i -lt $stops.Count - 1; $i++) {
        $a = $stops[$i]; $b = $stops[$i + 1]
        if ($t -ge $a.t -and $t -le $b.t) {
            $lt = ($t - $a.t) / ($b.t - $a.t)
            $r  = [int][Math]::Round($a.r + ($b.r - $a.r) * $lt)
            $g  = [int][Math]::Round($a.g + ($b.g - $a.g) * $lt)
            $bl = [int][Math]::Round($a.b + ($b.b - $a.b) * $lt)
            return $r, $g, $bl
        }
    }
    $s = $stops[-1]; return $s.r, $s.g, $s.b
}

# ---------------------------------------------------------------------------
# Render banner line-by-line with per-character horizontal gradient
# ---------------------------------------------------------------------------
$esc   = [char]27
$reset = "$esc[0m"
$bold  = "$esc[1m"

Write-Host ''
$lines = $banner -split "`r?`n"
# Use the longest visible line as the gradient width reference so the
# gradient stays aligned even if some rows are shorter.
$maxLen = 0
foreach ($ln in $lines) {
    if ($ln.Length -gt $maxLen) { $maxLen = $ln.Length }
}

foreach ($line in $lines) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $sb = [System.Text.StringBuilder]::new(($line.Length * 14))
    [void]$sb.Append($bold)
    for ($i = 0; $i -lt $line.Length; $i++) {
        $ch = $line[$i]
        if ($ch -eq ' ') {
            [void]$sb.Append(' ')
            continue
        }
        $t  = if ($maxLen -gt 1) { [double]$i / ($maxLen - 1) } else { 0.0 }
        $r, $g, $b = Get-GradientRGB $t
        [void]$sb.AppendFormat("{0}[38;2;{1};{2};{3}m{4}", $esc, $r, $g, $b, $ch)
    }
    [void]$sb.Append($reset)
    Write-Host ('  ' + $sb.ToString())
}

# ---------------------------------------------------------------------------
# Subtitle + meta block (ANSI-styled; falls back gracefully on legacy hosts)
# ---------------------------------------------------------------------------
$magentaDim = "$esc[38;2;236;72;153m"
$gray       = "$esc[38;2;120;120;120m"
$yellow     = "$esc[38;2;250;204;21m"

Write-Host ''
Write-Host ('    {0}{1}{2}' -f $bold, $name, $reset) -NoNewline
Write-Host ('  ' + $magentaDim + 'Local Edition' + $reset)

Write-Host ('    v{0}' -f $version) -NoNewline
Write-Host ('  ' + $gray + 'build' + $reset + ' ') -NoNewline
Write-Host $buildDate

Write-Host ''
Write-Host ('    ' + $gray + '-' + $reset + (' {0}\{1}  pid:{2}' -f $hostName, $userName, $pid)) -ForegroundColor DarkGray
Write-Host ('    ' + $gray + '-' + $reset + (' endpoint: {0}' -f $endpoint)) -ForegroundColor DarkGray
Write-Host ''