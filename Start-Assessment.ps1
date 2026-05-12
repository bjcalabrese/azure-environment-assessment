# Azure Environment Assessment — Windows Launcher
# Run this first on a fresh Windows machine — it installs Python if needed,
# then launches the interactive setup wizard.
#
# Usage (from PowerShell):
#   .\Start-Assessment.ps1
#
# If your execution policy blocks scripts, run once with:
#   powershell -ExecutionPolicy Bypass -File .\Start-Assessment.ps1

$ErrorActionPreference = "Stop"

# ── helpers ────────────────────────────────────────────────────────────────────
function Write-Header {
    Write-Host ""
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host "   Azure Environment Assessment — Launcher" -ForegroundColor Cyan
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step { param($msg) Write-Host "[*] $msg" -ForegroundColor Cyan }
function Write-Ok   { param($msg) Write-Host "[OK] $msg" -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host "[!] $msg" -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host "[X] $msg" -ForegroundColor Red }

# ── find python ────────────────────────────────────────────────────────────────
function Find-Python {
    foreach ($cmd in @("python", "python3", "py")) {
        try {
            $ver = & $cmd --version 2>&1
            if ($ver -match "Python (\d+)\.(\d+)") {
                $maj = [int]$Matches[1]
                $min = [int]$Matches[2]
                if ($maj -gt 3 -or ($maj -eq 3 -and $min -ge 10)) {
                    return $cmd
                }
            }
        } catch { }
    }
    return $null
}

# ── install python via winget ──────────────────────────────────────────────────
function Install-PythonWinget {
    Write-Step "Installing Python 3.12 via winget..."
    try {
        winget install --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements
        # Refresh PATH so the new python is visible in this session
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
                    [System.Environment]::GetEnvironmentVariable("PATH", "User")
        return $true
    } catch {
        return $false
    }
}

# ── install python via Microsoft Store ────────────────────────────────────────
function Install-PythonStore {
    Write-Warn "Attempting to open Python in the Microsoft Store..."
    Start-Process "ms-windows-store://pdp/?ProductId=9NRWMJLIVE28"
    Write-Host ""
    Write-Host "  The Microsoft Store is opening. Install Python 3.12, then re-run this script." -ForegroundColor Yellow
    Write-Host ""
    return $false
}

# ── manual install fallback ───────────────────────────────────────────────────
function Show-ManualInstall {
    Write-Host ""
    Write-Host "  Automatic Python installation was not available on this machine." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Install Python manually (3.10 or later):" -ForegroundColor White
    Write-Host "    1. Go to https://www.python.org/downloads/windows/" -ForegroundColor White
    Write-Host "    2. Download the latest Python 3.x installer" -ForegroundColor White
    Write-Host "    3. Run it — check 'Add Python to PATH' on the first screen" -ForegroundColor White
    Write-Host "    4. Re-run this script after installation" -ForegroundColor White
    Write-Host ""
}

# ── main ──────────────────────────────────────────────────────────────────────
Write-Header

# Step 1: Check for Python
Write-Step "Looking for Python 3.10 or later..."
$PYTHON = Find-Python

if ($null -eq $PYTHON) {
    Write-Warn "Python 3.10+ not found."
    Write-Host ""

    # Try winget first (available on Windows 10 1709+ and Windows 11)
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        $installed = Install-PythonWinget
        if ($installed) {
            $PYTHON = Find-Python
        }
    }

    # winget failed or not available — try Microsoft Store
    if ($null -eq $PYTHON) {
        $choice = Read-Host "  Open Python in the Microsoft Store to install? [Y/n]"
        if ($choice -eq "" -or $choice -match "^[Yy]") {
            Install-PythonStore
            exit 0
        }
    }

    # Still nothing — show manual instructions
    if ($null -eq $PYTHON) {
        Show-ManualInstall
        exit 1
    }
}

$verStr = & $PYTHON --version 2>&1
Write-Ok "Found $verStr (command: $PYTHON)"

# Step 2: Confirm script directory
# $PSScriptRoot is always set correctly whether run via -File, dot-sourced, or Run with PowerShell
$scriptDir = $PSScriptRoot
$wizardPath = Join-Path $scriptDir "setup_wizard.py"

if (-not (Test-Path $wizardPath)) {
    Write-Fail "setup_wizard.py not found in $scriptDir"
    Write-Host "  Make sure you are running this script from the azure-environment-assessment folder." -ForegroundColor Yellow
    exit 1
}

# Step 3: Hand off to the wizard
Write-Ok "Launching setup wizard..."
Write-Host ""

Set-Location $scriptDir
& $PYTHON "$wizardPath" @args
