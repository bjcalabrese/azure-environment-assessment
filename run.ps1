#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Azure Environment Assessment Tool — interactive launcher

.DESCRIPTION
    Checks prerequisites, guides you through subscription selection,
    runs the Python assessment, and opens the resulting Excel workbook.

.PARAMETER Subscription
    One or more subscription IDs. Skips the interactive menu.

.PARAMETER AllSubscriptions
    Scan every accessible subscription.

.PARAMETER Output
    Output .xlsx filename (default: azure_assessment_YYYYMMDD.xlsx).

.PARAMETER SkipSnapshots
    Skip disk snapshot enumeration (much faster on large subscriptions).

.PARAMETER Workers
    Parallel subscription workers (default: 4).

.PARAMETER Verbose
    Show detailed per-service logging.

.PARAMETER DryRun
    Show the command that would run without executing it.

.EXAMPLE
    .\run.ps1
    .\run.ps1 -AllSubscriptions -SkipSnapshots
    .\run.ps1 -Subscription "00000000-...", "11111111-..."
#>

[CmdletBinding()]
param(
    [string[]]$Subscription,
    [switch]$AllSubscriptions,
    [string]$Output,
    [switch]$SkipSnapshots,
    [int]$Workers = 4,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# ── Colour helpers ────────────────────────────────────────────────────────────
function Write-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Blue
    Write-Host "  ║   ☁  Azure Environment Assessment Tool              ║" -ForegroundColor Blue
    Write-Host "  ║   Read-only inventory  →  Excel workbook            ║" -ForegroundColor Blue
    Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Blue
    Write-Host ""
}

function Write-Step  { param([string]$N, [string]$T) Write-Host "`n  ── Step $N — $T" -ForegroundColor Cyan }
function Write-Ok    { param([string]$M) Write-Host "  ✔  $M" -ForegroundColor Green  }
function Write-Warn  { param([string]$M) Write-Host "  ⚠  $M" -ForegroundColor Yellow }
function Write-Fail  { param([string]$M) Write-Host "  ✖  $M" -ForegroundColor Red    }
function Write-Info  { param([string]$M) Write-Host "     $M" -ForegroundColor Gray   }

function Prompt-Default {
    param([string]$Prompt, [string]$Default)
    $val = Read-Host "  $Prompt [default: $Default]"
    if ([string]::IsNullOrWhiteSpace($val)) { $Default } else { $val.Trim() }
}

# ── STEP 1 — Python ───────────────────────────────────────────────────────────
Write-Banner
Write-Step 1 "Python"

$python = $null
foreach ($candidate in @("python3", "python", "py")) {
    try {
        $ver = & $candidate --version 2>&1
        if ($ver -match "Python 3\.(\d+)") {
            if ([int]$Matches[1] -ge 8) {
                $python = $candidate
                Write-Ok "$ver  ($candidate)"
                break
            } else {
                Write-Warn "$ver found — Python 3.8+ required"
            }
        }
    } catch {}
}

if (-not $python) {
    Write-Fail "Python 3.8+ not found."
    if ($IsWindows) { Write-Info "Install: winget install Python.Python.3  or  https://python.org" }
    elseif ($IsMacOS) { Write-Info "Install: brew install python" }
    else { Write-Info "Install: sudo apt install python3  or  https://python.org" }
    exit 1
}

# ── STEP 2 — Dependencies ─────────────────────────────────────────────────────
Write-Step 2 "Dependencies"

$reqFile = Join-Path $ScriptDir "requirements.txt"
if (-not (Test-Path $reqFile)) {
    Write-Fail "requirements.txt not found at $reqFile"
    exit 1
}

Write-Info "Installing from requirements.txt …"
& $python -m pip install -r $reqFile --quiet --upgrade
if ($LASTEXITCODE -ne 0) { Write-Fail "pip install failed"; exit 1 }
Write-Ok "All dependencies ready"

# ── STEP 3 — Azure authentication ────────────────────────────────────────────
Write-Step 3 "Azure Authentication"

$availableSubs = @()
$authMethod   = ""

# Try az CLI
try {
    $azRaw = & az account list --output json 2>$null
    if ($LASTEXITCODE -eq 0 -and $azRaw) {
        $parsed = $azRaw | ConvertFrom-Json
        $enabled = @($parsed | Where-Object { $_.state -eq "Enabled" })
        if ($enabled.Count -gt 0) {
            $availableSubs = $enabled | ForEach-Object {
                [PSCustomObject]@{ Id = $_.id; Name = $_.name }
            }
            $azUser = (& az account show --query "user.name" -o tsv 2>$null)
            $authMethod = "Azure CLI"
            Write-Ok "Azure CLI — signed in as $azUser"
        }
    }
} catch {}

# Fall back to DefaultAzureCredential via Python
if ($availableSubs.Count -eq 0) {
    Write-Info "az CLI not available — trying DefaultAzureCredential …"
    $pyCode = @'
from azure.identity import DefaultAzureCredential
from azure.mgmt.resource import SubscriptionClient
import json, sys
try:
    c = DefaultAzureCredential()
    sc = SubscriptionClient(c)
    subs = [s for s in sc.subscriptions.list() if str(s.state).lower() == "enabled"]
    print(json.dumps([{"Id": s.subscription_id, "Name": s.display_name} for s in subs]))
except Exception as e:
    sys.exit(1)
'@
    try {
        $sdkRaw = & $python -c $pyCode 2>$null
        if ($LASTEXITCODE -eq 0 -and $sdkRaw) {
            $parsed = $sdkRaw | ConvertFrom-Json
            $availableSubs = @($parsed | ForEach-Object {
                [PSCustomObject]@{ Id = $_.Id; Name = $_.Name }
            })
            if ($availableSubs.Count -gt 0) {
                $authMethod = "DefaultAzureCredential"
                Write-Ok "Authenticated via DefaultAzureCredential"
            }
        }
    } catch {}
}

if ($availableSubs.Count -eq 0) {
    Write-Fail "Not authenticated."
    Write-Host ""
    Write-Host "  Run one of these in your terminal, then try again:" -ForegroundColor White
    Write-Host "    az login" -ForegroundColor Yellow
    Write-Host "    az login --use-device-code   (MFA / headless)" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

Write-Info "$($availableSubs.Count) subscription(s) accessible via $authMethod"

# ── STEP 4 — Subscription selection ──────────────────────────────────────────
Write-Step 4 "Select Subscriptions"

$selectedIds = @()
$scanAll     = $false

if ($AllSubscriptions) {
    $selectedIds = @($availableSubs | ForEach-Object { $_.Id })
    $scanAll     = $true
    Write-Ok "All $($selectedIds.Count) subscription(s) selected"

} elseif ($Subscription) {
    $selectedIds = @($Subscription)
    Write-Ok "$($selectedIds.Count) subscription ID(s) from parameters"

} else {
    # Interactive
    Write-Host ""
    Write-Host "  How would you like to select subscriptions?" -ForegroundColor White
    Write-Host ""
    Write-Host "    [1]  Current / default  (auto-detect from environment)" -ForegroundColor Gray
    Write-Host "    [2]  Pick from list" -ForegroundColor Gray
    Write-Host "    [3]  All subscriptions  ($($availableSubs.Count) found)" -ForegroundColor Gray
    Write-Host ""
    $choice = Prompt-Default "Choice (1/2/3)" "1"

    switch ($choice) {
        "3" {
            $selectedIds = @($availableSubs | ForEach-Object { $_.Id })
            $scanAll     = $true
            Write-Ok "All $($selectedIds.Count) subscription(s) selected"
        }
        "2" {
            Write-Host ""
            for ($i = 0; $i -lt $availableSubs.Count; $i++) {
                $s = $availableSubs[$i]
                Write-Host ("    [{0,2}]  {1,-42}  {2}" -f ($i + 1), $s.Name, $s.Id) -ForegroundColor Gray
            }
            Write-Host ""
            $picks = Prompt-Default "Numbers separated by commas (e.g. 1,3)" "1"
            foreach ($p in ($picks -split ",")) {
                $idx = [int]$p.Trim() - 1
                if ($idx -ge 0 -and $idx -lt $availableSubs.Count) {
                    $selectedIds += $availableSubs[$idx].Id
                }
            }
            Write-Ok "$($selectedIds.Count) subscription(s) selected"
        }
        default {
            Write-Ok "Using default subscription (auto-detect)"
        }
    }
}

# ── STEP 5 — Options ──────────────────────────────────────────────────────────
Write-Step 5 "Options"
Write-Host ""

$dateStr = Get-Date -Format "yyyyMMdd"

if ([string]::IsNullOrWhiteSpace($Output)) {
    $Output = Prompt-Default "Output filename" "azure_assessment_$dateStr.xlsx"
}
if (-not $Output.EndsWith(".xlsx")) { $Output += ".xlsx" }
$outputPath = Join-Path $ScriptDir $Output

if (-not $SkipSnapshots -and -not $PSBoundParameters.ContainsKey('SkipSnapshots')) {
    $snapAns = Prompt-Default "Skip disk snapshot scan? Faster on large subscriptions (y/N)" "N"
    if ($snapAns -match "^[yY]") { $SkipSnapshots = $true }
}

if (-not $PSBoundParameters.ContainsKey('Workers')) {
    $wInput = Prompt-Default "Parallel workers" "$Workers"
    if ($wInput -match "^\d+$") { $Workers = [int]$wInput }
}

Write-Host ""
Write-Ok "Output   : $Output"
Write-Ok "Workers  : $Workers"
Write-Ok "Snapshots: $(if ($SkipSnapshots) { 'skipped' } else { 'included' })"

# ── STEP 6 — Run ──────────────────────────────────────────────────────────────
Write-Step 6 "Running Assessment"
Write-Host ""
Write-Info "Read-only scan — no changes will be made to your Azure environment."
Write-Host ""

$assessScript = Join-Path $ScriptDir "azure_assessment.py"
if (-not (Test-Path $assessScript)) {
    Write-Fail "azure_assessment.py not found at $assessScript"
    exit 1
}

# Build argument list
$pyArgs = @($assessScript, "--output", $outputPath, "--workers", "$Workers")

if ($scanAll) {
    $pyArgs += "--all-subscriptions"
} elseif ($selectedIds.Count -gt 0) {
    $pyArgs += @("--subscription") + $selectedIds
}

if ($SkipSnapshots)        { $pyArgs += "--skip-snapshots" }
if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"]) { $pyArgs += "--verbose" }

if ($DryRun) {
    Write-Host ""
    Write-Host "  DRY RUN — command that would execute:" -ForegroundColor Magenta
    Write-Host "  $python $($pyArgs -join ' ')" -ForegroundColor White
    Write-Host ""
    exit 0
}

& $python @pyArgs
$exitCode = $LASTEXITCODE

Write-Host ""

# ── STEP 7 — Done ─────────────────────────────────────────────────────────────
Write-Step 7 "Done"
Write-Host ""

if ($exitCode -eq 0 -and (Test-Path $outputPath)) {
    $sizeMB = [math]::Round((Get-Item $outputPath).Length / 1MB, 2)
    Write-Ok "Assessment complete!  →  $Output  ($sizeMB MB)"
    Write-Host ""

    try {
        if     ($IsWindows) { Start-Process $outputPath }
        elseif ($IsMacOS)   { & open $outputPath }
        else                { & xdg-open $outputPath 2>$null }
        Write-Info "Opening $Output …"
    } catch {
        Write-Info "Saved to: $outputPath"
    }
} else {
    Write-Fail "Assessment failed (exit $exitCode) — see output above."
    exit 1
}

Write-Host ""
