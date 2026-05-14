# Azure Environment Assessment — Windows Launcher
# Prerequisite: Python 3.10 or later must be installed and on PATH.
# Download from https://www.python.org/downloads/windows/
#
# Usage (from PowerShell):
#   powershell -ExecutionPolicy Bypass -File .\Start-Assessment.ps1

$ErrorActionPreference = "Stop"

function Write-Header {
    Write-Host ""
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host "   Azure Environment Assessment — Launcher" -ForegroundColor Cyan
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Ok   { param($msg) Write-Host "[OK] $msg" -ForegroundColor Green }
function Write-Fail { param($msg) Write-Host "[X] $msg" -ForegroundColor Red }

Write-Header

# Find Python 3.10+
$PYTHON = $null
foreach ($cmd in @("python", "python3", "py")) {
    try {
        $ver = & $cmd --version 2>&1
        if ($ver -match "Python (\d+)\.(\d+)") {
            if ([int]$Matches[1] -gt 3 -or ([int]$Matches[1] -eq 3 -and [int]$Matches[2] -ge 10)) {
                $PYTHON = $cmd
                break
            }
        }
    } catch { }
}

if ($null -eq $PYTHON) {
    Write-Fail "Python 3.10 or later not found."
    Write-Host ""
    Write-Host "  Install Python from https://www.python.org/downloads/windows/" -ForegroundColor Yellow
    Write-Host "  Check 'Add Python to PATH' during installation, then re-run this script." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

$verStr = & $PYTHON --version 2>&1
Write-Ok "Found $verStr"

$scriptDir = $PSScriptRoot
$wizardPath = Join-Path $scriptDir "setup_wizard.py"

if (-not (Test-Path $wizardPath)) {
    Write-Fail "setup_wizard.py not found in $scriptDir"
    Write-Host "  Run this script from the azure-environment-assessment folder." -ForegroundColor Yellow
    exit 1
}

Write-Ok "Launching setup wizard..."
Write-Host ""

Set-Location $scriptDir
& $PYTHON "$wizardPath" @args
