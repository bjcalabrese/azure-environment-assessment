# Quick Start Guide

> **Disclaimer:** This is a community sample script provided without support guarantees. It is not an official product and is not covered by any support agreement. Use at your own risk.

This guide gets you from zero to a completed Azure assessment in under 10 minutes.

---

## Step 1 — Install Python dependencies

You need Python 3.10 or later. Check with `python --version`.

```bash
pip install -r requirements.txt
```

This installs the Azure SDK management libraries, `openpyxl` (Excel writer), and `tqdm` (progress bars).

---

## Step 2 — Set up Azure credentials

The tool uses `DefaultAzureCredential` — the same credential chain used by the Azure CLI and Azure SDKs. Pick one of the following:

### Option A: Azure CLI (simplest for interactive use)

```bash
az login
az account show   # confirm the active subscription
```

If that returns your subscription name and ID, skip to Step 3.

### Option B: Target a specific subscription

```bash
az login
az account set --subscription "your-subscription-name-or-id"
az account show   # confirm it switched
```

### Option C: Service principal (for customer accounts or automation)

```bash
export AZURE_CLIENT_ID="your-app-registration-client-id"
export AZURE_CLIENT_SECRET="your-client-secret"
export AZURE_TENANT_ID="customer-tenant-id"
export AZURE_SUBSCRIPTION_ID="customer-subscription-id"
```

Windows PowerShell:
```powershell
$env:AZURE_CLIENT_ID     = "your-app-registration-client-id"
$env:AZURE_CLIENT_SECRET = "your-client-secret"
$env:AZURE_TENANT_ID     = "customer-tenant-id"
$env:AZURE_SUBSCRIPTION_ID = "customer-subscription-id"
```

---

## Step 3 — Run the assessment

```bash
# Scan your current active subscription
python azure_assessment.py

# Scan a specific subscription by ID
python azure_assessment.py --subscription 00000000-0000-0000-0000-000000000000

# Scan every accessible subscription
python azure_assessment.py --all-subscriptions

# Custom output filename
python azure_assessment.py --all-subscriptions --output "Customer_$(date +%Y%m%d).xlsx"
```

The `.xlsx` file is saved in the directory you run the script from.

---

## Step 4 — Open the workbook

Open the file in Excel or Google Sheets. Start with the **Summary** sheet — it gives you the full picture without needing to look at individual tabs.

---

## All command-line options

```
python azure_assessment.py [OPTIONS]

Options:
  --subscription SUB_ID [SUB_ID ...]   Subscription(s) to scan (default: current)
  --all-subscriptions                  Scan all accessible enabled subscriptions
  --tenant TENANT_ID                   Azure tenant ID (optional)
  --output FILENAME                    Output .xlsx filename
  --workers N                          Parallel subscription workers, default 4
  --skip-snapshots                     Skip disk snapshot enumeration (faster)
  --verbose                            Show detailed logging
```

---

## Example commands

### First run — single subscription with verbose output

```bash
python azure_assessment.py \
  --subscription 00000000-0000-0000-0000-000000000000 \
  --verbose
```

### Full account scan with a date-stamped file

```bash
# macOS / Linux
python azure_assessment.py --all-subscriptions --output "Assessment_$(date +%Y%m%d).xlsx"

# Windows PowerShell
$date = Get-Date -Format "yyyyMMdd"
python azure_assessment.py --all-subscriptions --output "Assessment_$date.xlsx"

# Windows Command Prompt
python azure_assessment.py --all-subscriptions --output "Assessment.xlsx"
```

### Customer account — complete workflow using service principal

```bash
# 1. Customer creates a service principal with Reader role (they run this)
az ad sp create-for-rbac \
  --name "AzureAssessmentReadOnly" \
  --role "Reader" \
  --scopes /subscriptions/<their-subscription-id>
#
# Output:
# {
#   "appId":    "aaaaaaaa-...",   ← AZURE_CLIENT_ID
#   "password": "xxxxxxxx-...",   ← AZURE_CLIENT_SECRET
#   "tenant":   "bbbbbbbb-...",   ← AZURE_TENANT_ID
# }

# 2. You set the credentials on your machine
export AZURE_CLIENT_ID="aaaaaaaa-..."
export AZURE_CLIENT_SECRET="xxxxxxxx-..."
export AZURE_TENANT_ID="bbbbbbbb-..."
export AZURE_SUBSCRIPTION_ID="<their-subscription-id>"

# 3. Verify access
az login --service-principal \
  --username $AZURE_CLIENT_ID \
  --password $AZURE_CLIENT_SECRET \
  --tenant $AZURE_TENANT_ID

# 4. Run the assessment
python azure_assessment.py \
  --output "Customer_Assessment_$(date +%Y%m%d).xlsx"
```

### Customer account — interactive login to their tenant

```bash
az login --tenant their-tenant-id
# (browser opens for authentication)

python azure_assessment.py \
  --tenant their-tenant-id \
  --all-subscriptions \
  --output "Customer_$(date +%Y%m%d).xlsx"
```

### Multiple subscriptions

```bash
python azure_assessment.py \
  --subscription 00000000-0000-0000-0000-000000000000 \
                 11111111-1111-1111-1111-111111111111 \
                 22222222-2222-2222-2222-222222222222 \
  --output "MultiSub_Assessment_$(date +%Y%m%d).xlsx"
```

### Large environment — fastest scan

```bash
python azure_assessment.py \
  --all-subscriptions \
  --skip-snapshots \
  --workers 8 \
  --output "LargeEnv_$(date +%Y%m%d).xlsx"
```

---

## IAM permissions

The tool is read-only. It never creates, modifies, or deletes anything.

**Simplest option** — assign the built-in Reader role:

```bash
# On a single subscription
az role assignment create \
  --assignee <service-principal-app-id> \
  --role "Reader" \
  --scope /subscriptions/<subscription-id>

# On a management group (covers all child subscriptions)
az role assignment create \
  --assignee <service-principal-app-id> \
  --role "Reader" \
  --scope /providers/Microsoft.Management/managementGroups/<management-group-id>
```

The Reader role includes all `*/read` actions needed by this tool. No custom policy is required.

---

## Tips for large environments

| Situation | Recommended flags |
|---|---|
| Subscription with 1,000+ disk snapshots | `--skip-snapshots` |
| Scanning 5+ subscriptions | `--workers 6` |
| Both of the above | `--all-subscriptions --skip-snapshots --workers 6` |
| First run / debugging | `--verbose` |

---

## Understanding the output

### Summary sheet layout

```
┌─────────────────────────────────────────────────┐
│         Azure Environment Assessment            │  ← Title + subscription/date
├──────────┬──────────┬──────────┬──────────┬─────┤
│  TOTAL   │ STORAGE  │  VMs     │  VMs     │ ... │  ← KPI tiles
│RESOURCES │  (TiB)   │ Running  │ Stopped  │     │
├──────────┴──────────┴──────────┴──────────┴─────┤
│                                                  │
│  Workload Inventory    │  Risk & Findings        │
│  (left column)         │  (right column)         │
│                        │                         │
│  Virtual Machines  45  │  CRITICAL  Public Blob  │
│  Managed Disks     80  │  HIGH      No HTTPS     │
│  Azure SQL         12  │  MEDIUM    No Backup    │
│  Storage Accounts  28  │                         │
│  ...               ..  │  Backup Infrastructure  │
│                        │  Region Distribution    │
│                        │  Storage by Service     │
└────────────────────────┴─────────────────────────┘
```

### Colour coding

| Colour | Meaning | Examples |
|---|---|---|
| Red | Critical risk | Public blob access, SQL public endpoint, non-SSL Redis |
| Yellow | Warning | Deallocated VM, unattached disk, old snapshot |
| Green | Compliant / protected | HTTPS-only storage, SSL-only Redis, backup vault with items |

---

## Troubleshooting

**`DefaultAzureCredential failed to retrieve a token`**
No credentials are configured. Run `az login` or set the `AZURE_CLIENT_ID` / `AZURE_CLIENT_SECRET` / `AZURE_TENANT_ID` environment variables.

**`AuthorizationFailed` on a specific service**
The credential doesn't have Reader access on that subscription or resource group. Verify the role assignment with:
```bash
az role assignment list --assignee <service-principal-app-id> --all
```

**Storage sizes show `N/A`**
Azure Monitor storage metrics update once every 24 hours. For newly created storage accounts the metric history won't exist yet — values will appear the following day.

**Scan is slow**
Add `--skip-snapshots` (snapshot enumeration is slow in subscriptions with many disks) and increase `--workers` to match the number of subscriptions you're scanning.

**`ModuleNotFoundError`**
Run `pip install -r requirements.txt` to install all dependencies.
