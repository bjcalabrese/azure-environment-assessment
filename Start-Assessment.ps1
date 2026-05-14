# Azure Environment Assessment — Windows Launcher
# If Python 3.10+ is not found, prompts to download and install it automatically.
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

function Write-Step { param($msg) Write-Host "[*] $msg" -ForegroundColor Cyan }
function Write-Ok   { param($msg) Write-Host "[OK] $msg" -ForegroundColor Green }
function Write-Fail { param($msg) Write-Host "[X] $msg" -ForegroundColor Red }
function Write-Warn { param($msg) Write-Host "[!] $msg" -ForegroundColor Yellow }

function Find-Python {
    foreach ($cmd in @("python", "python3", "py")) {
        try {
            $ver = & $cmd --version 2>&1
            if ($ver -match "Python (\d+)\.(\d+)") {
                if ([int]$Matches[1] -gt 3 -or ([int]$Matches[1] -eq 3 -and [int]$Matches[2] -ge 10)) {
                    return $cmd
                }
            }
        } catch { }
    }
    return $null
}

function Install-Python {
    Write-Step "Downloading Python 3.12 from python.org..."
    $installer = "$env:TEMP\python-3.12.9-amd64.exe"
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri "https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe" `
            -OutFile $installer -UseBasicParsing
        Write-Step "Installing Python 3.12 (this may take a minute)..."
        Start-Process $installer -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0" -Wait
        # Refresh PATH in this session so python is immediately available
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
                    [System.Environment]::GetEnvironmentVariable("PATH", "User")
        return $true
    } catch {
        Write-Warn "Download or install failed: $_"
        return $false
    }
}

Write-Header

$PYTHON = Find-Python

if ($null -eq $PYTHON) {
    Write-Warn "Python 3.10 or later was not found on this machine."
    Write-Host ""
    $choice = Read-Host "  Download and install Python 3.12 automatically? [Y/n]"
    if ($choice -eq "" -or $choice -match "^[Yy]") {
        $ok = Install-Python
        if ($ok) {
            $PYTHON = Find-Python
        }
    }

    if ($null -eq $PYTHON) {
        Write-Host ""
        Write-Fail "Python not available. Install manually from https://www.python.org/downloads/windows/"
        Write-Host "  Check 'Add Python to PATH' during installation, then re-run this script." -ForegroundColor Yellow
        Write-Host ""
        exit 1
    }
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
