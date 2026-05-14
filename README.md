# Azure Environment Assessment Tool

> **Disclaimer:** This is a community sample script provided without support guarantees. It is not an official product and is not covered by any support agreement. Use at your own risk. Review the code before running it in any environment.

A read-only Azure inventory tool that scans your subscription and produces a single Excel workbook covering every major workload type. The Azure equivalent of RVTools.

The output is a colour-coded, multi-sheet spreadsheet your team can use to understand what's running, what's at risk, and how much storage needs to be protected.

---

## How it works

1. You authenticate using your existing Azure credentials (`az login` or a service principal)
2. It scans every enabled subscription you specify (or all accessible ones)
3. It writes a single `.xlsx` file with one sheet per service type plus a summary dashboard

The script is **100% read-only** — it only calls `List*`, `Get*`, and `Describe*` equivalent APIs. It makes no changes to your environment.

---

## Prerequisites

**Python 3.10 or later** — the OS launchers will prompt to install this automatically if missing. To install manually: [python.org/downloads](https://www.python.org/downloads/)

**Azure CLI** — required for authentication.

```bash
az login
az account show   # confirm which subscription is active
```

If you can run `az account show` and see your subscription, you're ready.

**Python packages**

```bash
pip install -r requirements.txt
```

---

## Setup wizard (recommended for first-time use)

Use the launcher for your OS — it checks for Python, installs it if missing, then walks you through everything interactively.

**Windows** — open PowerShell in the project folder:
```powershell
powershell -ExecutionPolicy Bypass -File .\Start-Assessment.ps1
```
If Python 3.10+ is not found, the launcher will offer to download and install it automatically from python.org.

**macOS / Linux** — open Terminal in the project folder:
```bash
./start-assessment.sh
```
If the script isn't executable yet:
```bash
chmod +x start-assessment.sh && ./start-assessment.sh
```
If Python is missing, the launcher installs it via Homebrew (macOS) or your system package manager (`apt`, `dnf`, `yum`, `pacman`) on Linux.

The wizard walks you through 7 steps:
1. Python version check with OS-specific upgrade instructions
2. Azure CLI check with install instructions per OS
3. `pip install -r requirements.txt` with live output
4. Authentication — interactive login, service principal, or existing env vars
5. Subscription selection — all, specific, or current
6. Scan options — skip snapshots, workers, output filename, anonymize, Scenario Builder export
7. Run the assessment and open the workbook

---

## Quickstart

```bash
# Scan your current active subscription
python azure_assessment.py

# Scan a specific subscription by ID
python azure_assessment.py --subscription 00000000-0000-0000-0000-000000000000

# Scan every accessible subscription
python azure_assessment.py --all-subscriptions
```

The output file is saved in the current directory:
```
azure_assessment_<date>.xlsx
```

Open it in Excel or Google Sheets.

---

## All options

| Flag | Description | Default |
|---|---|---|
| `--subscription` | One or more subscription IDs to scan | Current active subscription |
| `--all-subscriptions` | Scan every enabled subscription accessible with current credentials | — |
| `--tenant` | Azure tenant ID (for multi-tenant environments) | Default tenant |
| `--output` | Output `.xlsx` filename | `azure_assessment_<date>.xlsx` |
| `--workers` | Number of subscriptions scanned in parallel | `4` |
| `--skip-snapshots` | Skip disk snapshot enumeration | — |
| `--anonymize` | Replace all resource names with opaque codes; saves a reversible mapping CSV alongside the workbook | — |
| `--scenario-builder` | Write a second file in Veeam Scenario Builder (CAzureWrapper) import format | — |
| `--verbose` | Print detailed per-service logging | — |

---

## Examples

### Basic scans

```bash
# Scan your current active subscription
python azure_assessment.py

# Scan a specific subscription
python azure_assessment.py --subscription 00000000-0000-0000-0000-000000000000

# Scan multiple specific subscriptions
python azure_assessment.py \
  --subscription 00000000-0000-0000-0000-000000000000 \
                 11111111-1111-1111-1111-111111111111

# Scan every accessible subscription
python azure_assessment.py --all-subscriptions
```

### Customer accounts

```bash
# Option 1: Customer logs you in directly with az login
az login
python azure_assessment.py --all-subscriptions \
  --output "Customer_Assessment_$(date +%Y%m%d).xlsx"
```

```bash
# Option 2: Service principal credentials
export AZURE_CLIENT_ID="your-app-id"
export AZURE_CLIENT_SECRET="your-secret"
export AZURE_TENANT_ID="their-tenant-id"
export AZURE_SUBSCRIPTION_ID="their-subscription-id"

python azure_assessment.py \
  --output "Customer_Assessment_$(date +%Y%m%d).xlsx"
```

```bash
# Option 3: Specific tenant with interactive login
az login --tenant their-tenant-id
python azure_assessment.py \
  --tenant their-tenant-id \
  --all-subscriptions \
  --output "Customer_Assessment_$(date +%Y%m%d).xlsx"
```

### Large or complex environments

```bash
# Skip snapshots — saves time on subscriptions with thousands of disks
python azure_assessment.py --all-subscriptions --skip-snapshots

# Increase parallel workers for faster multi-subscription scans
python azure_assessment.py --all-subscriptions --workers 8

# Both together — fastest possible full scan
python azure_assessment.py \
  --all-subscriptions \
  --skip-snapshots \
  --workers 8 \
  --output "LargeAccount_$(date +%Y%m%d).xlsx"
```

### Anonymize output (for sharing without exposing resource names)

```bash
python azure_assessment.py --all-subscriptions --anonymize \
  --output "Customer_$(date +%Y%m%d).xlsx"
# Produces Customer_<date>.xlsx  +  Customer_<date>_mapping.csv
```

### Veeam Scenario Builder export

```bash
python azure_assessment.py --all-subscriptions --scenario-builder \
  --output "Customer_$(date +%Y%m%d).xlsx"
# Produces Customer_<date>.xlsx  +  Customer_<date>_scenario_builder.xlsx
# Import the _scenario_builder.xlsx at veeam.com/calculators/scenario/build/cloud/azure
```

### Targeted scans

```bash
# Single subscription with verbose logging (useful for first run or debugging)
python azure_assessment.py \
  --subscription 00000000-0000-0000-0000-000000000000 \
  --verbose
```

### Windows users (PowerShell)

```powershell
# Basic scan
python azure_assessment.py

# Date-stamped output
$date = Get-Date -Format "yyyyMMdd"
python azure_assessment.py --all-subscriptions --output "Assessment_$date.xlsx"
```

### Windows users (Command Prompt)

```cmd
rem Basic scan
python azure_assessment.py

rem Specific subscription
python azure_assessment.py --subscription 00000000-0000-0000-0000-000000000000 --output "Assessment.xlsx"
```

---

## Authentication

The tool uses `DefaultAzureCredential` from the Azure SDK, which automatically tries the following in order:

| Method | How to set up |
|---|---|
| **Azure CLI** | Run `az login` — the simplest option for interactive use |
| **Environment variables** | Set `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, `AZURE_TENANT_ID` |
| **Managed Identity** | Automatic when running on an Azure VM or container |
| **Azure Developer CLI** | Run `azd auth login` |
| **Interactive browser** | Automatic fallback if nothing else is configured |

For customer assessments, the recommended approach is a service principal with read-only permissions.

---

## Setting up a service principal for a customer account

```bash
# Step 1: Create a service principal in the customer's tenant (they run this)
az ad sp create-for-rbac \
  --name "AzureAssessmentReadOnly" \
  --role "Reader" \
  --scopes /subscriptions/<subscription-id>

# Output:
# {
#   "appId":       "xxxxxxxx-...",   ← AZURE_CLIENT_ID
#   "password":    "xxxxxxxx-...",   ← AZURE_CLIENT_SECRET
#   "tenant":      "xxxxxxxx-...",   ← AZURE_TENANT_ID
# }

# Step 2: You set the environment variables on your machine
export AZURE_CLIENT_ID="appId from above"
export AZURE_CLIENT_SECRET="password from above"
export AZURE_TENANT_ID="tenant from above"
export AZURE_SUBSCRIPTION_ID="subscription-id"

# Step 3: Run the assessment
python azure_assessment.py --output "Customer_$(date +%Y%m%d).xlsx"
```

For multiple subscriptions, assign the Reader role at the management group or tenant level instead of per-subscription.

---

## IAM permissions required

The tool only needs read-only access. The built-in **Reader** role covers everything:

```bash
# Assign Reader role to a service principal on a subscription
az role assignment create \
  --assignee <service-principal-app-id> \
  --role "Reader" \
  --scope /subscriptions/<subscription-id>

# Or at management group level (covers all child subscriptions)
az role assignment create \
  --assignee <service-principal-app-id> \
  --role "Reader" \
  --scope /providers/Microsoft.Management/managementGroups/<mg-id>
```

The **Reader** role grants access to all `*/read` actions across every resource type — no custom policy needed.

---

## What's in the workbook

### Summary dashboard (first sheet)

| Section | What it shows |
|---|---|
| **KPI tiles** | Total resources, total storage (TiB), VMs running/stopped, SQL databases, storage accounts, VM backup coverage (protected/total %), current month cloud spend |
| **Workload inventory** | Every service type with resource count and storage in GiB/TiB — including SQL MI, Elastic Pools, and SQL Server VMs |
| **Risk & Findings** | Colour-coded CRITICAL / HIGH / MEDIUM findings — public blob access, SQL with public access, storage without HTTPS-only, unattached disks, Redis with non-SSL port, SQL Server VMs without backup, VMs without backup, no soft delete |
| **Azure Backup infrastructure** | Vault count, VM protected items count, SQL protected items count |
| **Region distribution** | How many resources are in each Azure region |
| **Storage by service** | Which services consume the most storage, ranked |
| **Backup sizing summary** | Protectable storage by service type (managed disks, blob, files, ANF, SQL MI) with suggested backup method |
| **Disk snapshot coverage** | Breakdown of disks by snapshot age — current, aging, stale, no snapshot |
| **Monthly spend by service** | Top 10 Azure services by current-month cost |

### Detail sheets (one per service)

| Sheet | What you get |
|---|---|
| **Virtual Machines** | Name, size (SKU), OS type, power state, OS disk, data disks, total storage, zones, tags, MSSQL-INSTALLED flag, backup policy name, backup protected status |
| **Managed Disks** | SKU, size, IOPS, throughput, encryption type, disk state (attached/unattached), attached VM |
| **Disk Snapshots** | Source disk, size, encryption, creation date, age in days |
| **Azure SQL** | Server, database, SKU, tier, max/allocated/used storage, elastic pool, PITR days, LTR weekly/monthly/yearly retention, backup redundancy, public access, TDE status |
| **SQL MI Databases** | Managed Instance, database name, license type, vCores, storage, status, collation, earliest restore point, PITR days, LTR weekly/monthly/yearly retention |
| **SQL Elastic Pools** | Pool name, SKU, tier, eDTUs/vCores, max/allocated/used storage, database count, zone redundancy |
| **SQL Server VMs** | VMs with SQL Server extension: image offer/SKU, license type, patching day, backup enabled |
| **Storage Accounts** | SKU, kind, HTTPS-only, public blob access, encryption key source, blob tier breakdown (hot/cool/cold/archive), file size, total size |
| **Azure File Shares** | Share name, protocol, access tier, quota, used size, UNC/mount path |
| **Azure NetApp Files** | Account, pool, volume, service level, quota, used size, throughput, protocols, mount path, snapshot details |
| **Cosmos DB** | API kind, consistency level, multi-region write, backup mode and retention, public access, live data/index size (GiB), document count, partition count |
| **Synapse Analytics** | Workspace, SQL pool, SKU, status, geo-backup |
| **AKS** | Cluster version, node pools, current/max node count, autoscale flag, node VM sizes, OS disk sizes per pool, network plugin, RBAC |
| **Container Instances** | Container group, OS type, CPU, memory, state, IP address |
| **Function Apps** | Runtime, OS type, app service plan, state |
| **Azure Virtual Desktop** | Host pool type, load balancer, max sessions, session host count |
| **Redis Cache** | SKU, capacity, Redis version, TLS settings, geo-replication |
| **Backup Vaults** | Redundancy type, protected items count |
| **Backup Protected Items** | Vault, item name, item type, protection status, last backup time, policy name |
| **Backup SQL Items** | SQL Server and SQL database protected items (AzureWorkload) across all vaults: server/instance, database, workload type, protection status, last backup status, policy name |
| **Backup Policies** | All VM and SQL backup policies across all vaults: schedule frequency, daily/weekly/monthly/yearly retention |
| **Backup Costs** | Recovery Services Vault spend over the last 30 days per vault, with totals |
| **Monthly Cloud Spend** | Azure spend by service category for the current month and previous month, with month-over-month change; top services also surfaced on the Summary dashboard |

---

## Colour coding

| Colour | Meaning |
|---|---|
| 🔴 Red | Critical gap — public access enabled, no HTTPS, non-SSL Redis port, zero backup coverage |
| 🟡 Yellow | Warning — deallocated/stopped VM, unattached disk, old snapshot |
| 🟢 Green | Protected / compliant |

---

## Security

### The tool is 100% read-only

Every Azure API call made by this script is a read operation. There are no `Create`, `Put`, `Update`, `Delete`, or `Patch` calls anywhere in the code. It is **not possible** for this tool to create, change, or delete any resource in your subscription.

You can verify this yourself:
```bash
grep -E "(create_|delete_|update_|patch_|begin_create|begin_delete)" azure_assessment.py
# Returns nothing
```

### Credentials stay on your machine

- Azure credentials are passed directly to the `DefaultAzureCredential` from the Azure SDK
- Credentials are **never** printed, logged, written to files, or transmitted anywhere other than to Azure API endpoints (`*.azure.com`, `*.microsoft.com`)
- The script has no knowledge of your credentials; it only calls `DefaultAzureCredential()` and the SDK handles everything else

### No data leaves your machine

- The only output is the `.xlsx` file written locally to your current directory
- The script makes **no HTTP calls** to any server other than official Azure API endpoints
- There is no telemetry, no analytics, no call-home behaviour of any kind
- No third-party libraries with network capability are used — only `azure-*` (Microsoft), `openpyxl` (local Excel writing), and `tqdm` (local progress bar)

### What's in the output file

The `.xlsx` file contains only resource **metadata** — the same information visible in the Azure Portal:

- Resource names, types, sizes, and SKUs
- Configuration flags (HTTPS-only: yes/no, public access: yes/no, TDE: yes/no)
- Regions and resource groups
- Tags you've applied to resources
- Counts and storage totals

It does **not** contain:
- Azure credentials, client secrets, or certificates
- Storage account access keys or connection strings
- Database passwords or connection strings
- Any data stored inside your resources (no blob contents, no database rows)

### Open source

The full source code is in this repository. There are no compiled binaries, no obfuscated code, and no external dependencies beyond the packages in `requirements.txt`.

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `DefaultAzureCredential failed` | No credentials found | Run `az login` or set service principal env vars |
| `AuthorizationFailed` on a service | Missing Reader role on that scope | Ensure Reader role is assigned at subscription level |
| `ResourceNotFoundError` on a service | Service not enabled in this subscription | Normal — the script skips it and continues |
| Storage sizes show `N/A` | Azure Monitor metrics update once daily for new accounts | Metrics will populate within 24 hours |
| Scan takes a long time | Large subscription or many subscriptions | Add `--skip-snapshots` and increase `--workers` |
| `ModuleNotFoundError` | Missing Azure SDK package | Run `pip install -r requirements.txt` |

---

## Files

| File | Purpose |
|---|---|
| `azure_assessment.py` | The assessment script |
| `setup_wizard.py` | Interactive setup wizard — detects OS, installs prerequisites, guides auth and launches the scan |
| `Start-Assessment.ps1` | Windows launcher — checks for Python 3.10+, offers to install it automatically if missing, then runs the wizard |
| `start-assessment.sh` | macOS / Linux launcher — installs Python via Homebrew or system package manager if missing, then runs the wizard |
| `veeam_scenario_builder_template.xlsx` | Template used by `--scenario-builder` to produce a Veeam-compatible import file |
| `requirements.txt` | Python dependencies |
| `QUICKSTART.md` | Step-by-step setup guide |
