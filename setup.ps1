# VoiceClaude Setup & Launch
# Checks all components, installs missing, then runs main.py

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Write-Step($msg) { Write-Host $msg -ForegroundColor Cyan }
function Write-OK($msg)   { Write-Host "  [OK] $msg" -ForegroundColor Green }
function Write-FIX($msg)  { Write-Host "  [INSTALLING] $msg" -ForegroundColor Yellow }
function Write-FAIL($msg) { Write-Host "  [ERROR] $msg" -ForegroundColor Red }

Write-Host ""
Write-Host "  VoiceClaude - Checking components..." -ForegroundColor White
Write-Host ""

# ── 1. Python ──────────────────────────────────────────────────────────────
Write-Step "1. Python"
$pythonExe = $null
$pythonwExe = $null

$candidates = @(
    "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python310\python.exe",
    "C:\Program Files\Python313\python.exe",
    "C:\Program Files\Python312\python.exe",
    "C:\Program Files\Python311\python.exe",
    "C:\Program Files\Python310\python.exe"
)

# Also check via where.exe, skip WindowsApps stub
$whereResult = where.exe python 2>$null
if ($whereResult) {
    foreach ($p in $whereResult) {
        if ($p -notlike "*WindowsApps*" -and (Test-Path $p)) {
            $pythonExe = $p
            break
        }
    }
}

# Check standard paths if not found via where
if (-not $pythonExe) {
    foreach ($c in $candidates) {
        if (Test-Path $c) { $pythonExe = $c; break }
    }
}

if (-not $pythonExe) {
    Write-FIX "Installing Python 3.11..."
    winget install Python.Python.3.11 --source winget -e --accept-source-agreements --accept-package-agreements
    # Refresh and search again
    foreach ($c in $candidates) {
        if (Test-Path $c) { $pythonExe = $c; break }
    }
}

if (-not $pythonExe) {
    Write-FAIL "Python not found after install. Please restart and try again."
    Read-Host "Press Enter to exit"
    exit 1
}

$pythonwExe = $pythonExe -replace "python\.exe$", "pythonw.exe"
if (-not (Test-Path $pythonwExe)) { $pythonwExe = $pythonExe }

$ver = & $pythonExe --version 2>&1
Write-OK "$ver  ($pythonExe)"

# ── 2. AutoHotkey ──────────────────────────────────────────────────────────
Write-Step "2. AutoHotkey v2"
$ahkPaths = @(
    "$env:LOCALAPPDATA\Programs\AutoHotkey\v2\AutoHotkey64.exe",
    "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe",
    "C:\Program Files (x86)\AutoHotkey\v2\AutoHotkey64.exe"
)
$ahkFound = $false
foreach ($p in $ahkPaths) {
    if (Test-Path $p) { $ahkFound = $true; Write-OK $p; break }
}

if (-not $ahkFound) {
    Write-FIX "Installing AutoHotkey v2..."
    winget install AutoHotkey.AutoHotkey --source winget -e --accept-source-agreements --accept-package-agreements
    foreach ($p in $ahkPaths) {
        if (Test-Path $p) { $ahkFound = $true; Write-OK $p; break }
    }
}

if (-not $ahkFound) {
    Write-FAIL "AutoHotkey not found. Text paste into VSCode will not work."
}

# ── 3. Python packages ─────────────────────────────────────────────────────
Write-Step "3. Python packages"
$packages = @("faster_whisper", "keyboard", "pyaudio")
$missing = @()

foreach ($pkg in $packages) {
    $check = & $pythonExe -c "import $pkg" 2>&1
    if ($LASTEXITCODE -ne 0) {
        $missing += $pkg -replace "_", "-"
    } else {
        Write-OK $pkg
    }
}

if ($missing.Count -gt 0) {
    Write-FIX "Installing: $($missing -join ', ')"
    & $pythonExe -m pip install --upgrade pip --quiet
    & $pythonExe -m pip install @($missing)
    if ($LASTEXITCODE -ne 0 -and $missing -contains "pyaudio") {
        Write-FIX "Trying pyaudio via pipwin..."
        & $pythonExe -m pip install pipwin --quiet
        & $pythonExe -m pipwin install pyaudio
    }
    # Verify after install
    $failed = @()
    foreach ($pkg in $packages) {
        $check = & $pythonExe -c "import $pkg" 2>&1
        if ($LASTEXITCODE -ne 0) { $failed += $pkg }
    }
    if ($failed.Count -gt 0) {
        Write-FAIL "Failed to install: $($failed -join ', ')"
        Read-Host "Press Enter to exit"
        exit 1
    }
    Write-OK "All packages installed"
}

# ── Launch ─────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  All OK - Starting VoiceClaude..." -ForegroundColor Green
Write-Host ""
Start-Sleep -Milliseconds 800

$mainScript = Join-Path $scriptDir "main.py"
Start-Process $pythonwExe -ArgumentList "`"$mainScript`"" -Verb RunAs
