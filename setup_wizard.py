#!/usr/bin/env python3
"""
Azure Environment Assessment — Setup Wizard

Detects your OS, verifies prerequisites, guides Azure authentication,
lets you choose subscriptions and options, then launches the assessment.

Usage:
    python setup_wizard.py
"""

import sys
import os
import platform
import subprocess
import shutil
import json
import datetime
import textwrap

# ── Minimum Python to run the wizard itself ───────────────────────────────────
if sys.version_info < (3, 8):
    print(f"Python 3.8 or later is required to run this wizard (you have {sys.version}).")
    sys.exit(1)

OS  = platform.system()   # "Windows", "Darwin", "Linux"
PY  = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
DIR = os.path.dirname(os.path.abspath(__file__))

# ── ANSI colour support ───────────────────────────────────────────────────────
def _enable_color():
    if OS == "Windows":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleMode(
                ctypes.windll.kernel32.GetStdHandle(-11), 7)
            return True
        except Exception:
            return False
    return hasattr(sys.stdout, "isatty") and sys.stdout.isatty()

_CLR = _enable_color()

def _c(t, code): return f"\033[{code}m{t}\033[0m" if _CLR else str(t)
def bold(t):    return _c(t, "1")
def green(t):   return _c(t, "92")
def yellow(t):  return _c(t, "93")
def red(t):     return _c(t, "91")
def blue(t):    return _c(t, "94")
def cyan(t):    return _c(t, "96")
def dim(t):     return _c(t, "2")
def white(t):   return _c(t, "97")

# ── UI helpers ────────────────────────────────────────────────────────────────
def hr(ch="─", w=68):  print(dim(ch * w))
def blank():           print()

def banner():
    blank()
    hr("═")
    print(bold(white("  Azure Environment Assessment  ·  Setup Wizard")))
    os_icon = {"Windows": "Windows", "Darwin": "macOS", "Linux": "Linux"}.get(OS, OS)
    print(dim(f"  Platform: {os_icon}   Python: {PY}"))
    hr("═")
    blank()

def step_header(n, total, title, subtitle=""):
    blank()
    print(f"  {bold(cyan(f'Step {n} / {total}'))}   {bold(white(title))}")
    if subtitle:
        print(f"  {dim(subtitle)}")
    hr()

def ok(msg):    print(f"  {green('✓')}  {msg}")
def warn(msg):  print(f"  {yellow('⚠')}  {msg}")
def fail(msg):  print(f"  {red('✗')}  {msg}")
def info(msg):  print(f"  {blue('·')}  {msg}")
def note(msg):  print(f"  {dim(msg)}")

def indent(text, prefix="      "):
    for line in text.strip().splitlines():
        print(prefix + line)

def ask(prompt, default=None, secret=False, required=False):
    suffix = f" {dim('[' + str(default) + ']')}" if default is not None else ""
    try:
        while True:
            if secret:
                import getpass
                val = getpass.getpass(f"\n  {bold(cyan('?'))} {prompt}{suffix}  ")
            else:
                val = input(f"\n  {bold(cyan('?'))} {prompt}{suffix}  ").strip()
            if not val and default is not None:
                return default
            if val:
                return val
            if not required:
                return ""
            print(f"  {red('A value is required.')}")
    except (KeyboardInterrupt, EOFError):
        print("\n\n  Wizard cancelled.")
        sys.exit(0)

def ask_yn(prompt, default="y"):
    hint = f"{bold('Y')}/n" if default.lower() == "y" else f"y/{bold('N')}"
    val = ask(f"{prompt} {dim('(' + hint + ')')}", default=default)
    return str(val).strip().lower() in ("y", "yes", "1", "true")

def ask_choice(prompt, choices, default=1):
    """Show a numbered menu, return (value, label) of the chosen item."""
    blank()
    print(f"  {bold(prompt)}")
    blank()
    for i, (label, _) in enumerate(choices, 1):
        marker = green(" ◀ default") if i == default else ""
        print(f"    {bold(cyan(str(i)))}  {label}{marker}")
    blank()
    while True:
        raw = ask("Enter number", default=str(default))
        try:
            idx = int(str(raw).strip()) - 1
            if 0 <= idx < len(choices):
                label, value = choices[idx]
                return value, label
        except (ValueError, TypeError):
            pass
        print(f"  {red('Please enter a number between 1 and ' + str(len(choices)))}")

def run_cmd(cmd, capture=True):
    try:
        r = subprocess.run(cmd, shell=True, capture_output=capture, text=True)
        return r.returncode, r.stdout.strip(), r.stderr.strip()
    except Exception as e:
        return -1, "", str(e)

def run_live(cmd):
    return subprocess.run(cmd, shell=True).returncode

def pause(msg="Press Enter to continue..."):
    try:
        input(f"\n  {dim(msg)}")
    except (KeyboardInterrupt, EOFError):
        sys.exit(0)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — Python version
# ─────────────────────────────────────────────────────────────────────────────
def check_python(total):
    step_header(1, total, "Python version",
                "This tool requires Python 3.10 or later.")

    if sys.version_info >= (3, 10):
        ok(f"Python {PY}  ✓")
        return

    warn(f"Python {PY} detected — version 3.10 or later is required.")
    blank()
    print(bold("  How to upgrade:"))
    blank()

    if OS == "Windows":
        indent(f"""
Option 1 — winget (Windows 10 / 11):
  {cyan('winget install Python.Python.3.12')}

Option 2 — Download installer:
  {cyan('https://python.org/downloads/')}
  Tick "Add Python to PATH" during install.
""")
    elif OS == "Darwin":
        path_cmd = 'echo \'export PATH="/opt/homebrew/bin:$PATH"\' >> ~/.zshrc'
        indent(f"""
Option 1 — Homebrew (recommended):
  {cyan('brew install python@3.12')}
  {cyan(path_cmd)}
  {cyan('source ~/.zshrc')}

Option 2 — Download installer:
  {cyan('https://python.org/downloads/')}
""")
    else:
        indent(f"""
Ubuntu / Debian:
  {cyan('sudo apt update && sudo apt install python3.12 python3.12-venv python3-pip')}

Fedora / RHEL 9+:
  {cyan('sudo dnf install python3.12')}

Arch Linux:
  {cyan('sudo pacman -S python')}
""")

    blank()
    fail("Please upgrade Python, then re-run this wizard.")
    sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — Azure CLI
# ─────────────────────────────────────────────────────────────────────────────
def check_az_cli(total):
    step_header(2, total, "Azure CLI",
                "The Azure CLI handles authentication and subscription discovery.")

    # Augment PATH with common Homebrew locations so az is found even when
    # the shell doesn't source ~/.zshrc (e.g. when launched via a launcher script)
    extra = ["/opt/homebrew/bin", "/usr/local/bin", "/home/linuxbrew/.linuxbrew/bin"]
    for p in extra:
        if p not in os.environ.get("PATH", ""):
            os.environ["PATH"] = p + os.pathsep + os.environ.get("PATH", "")

    rc, out, _ = run_cmd("az --version")
    if rc == 0:
        line = out.splitlines()[0] if out else "found"
        ok(f"Azure CLI found — {dim(line)}")
        return

    warn("Azure CLI not found on your PATH.")
    blank()
    note("The Azure CLI is free, open-source, and takes about 2 minutes to install.")
    blank()
    print(bold("  Install instructions:"))
    blank()

    if OS == "Windows":
        indent(f"""
Option 1 — winget (simplest):
  {cyan('winget install Microsoft.AzureCLI')}

Option 2 — MSI installer (download and double-click):
  {cyan('https://aka.ms/installazurecliwindows')}

Option 3 — PowerShell one-liner:
  {cyan('$ProgressPreference = "SilentlyContinue"')}
  {cyan('Invoke-WebRequest -Uri https://aka.ms/installazurecliwindows -OutFile AzureCLI.msi')}
  {cyan('Start-Process msiexec.exe -Wait -ArgumentList "/I AzureCLI.msi /quiet"')}

After installing, close and reopen your terminal, then re-run this wizard.
""")
    elif OS == "Darwin":
        indent(f"""
Option 1 — Homebrew (recommended):
  {cyan('brew update && brew install azure-cli')}

Option 2 — Script installer:
  {cyan('curl -L https://aka.ms/InstallAzureCli | bash')}

After installing, open a new terminal tab and re-run this wizard.
""")
    else:
        indent(f"""
Ubuntu / Debian (official Microsoft repo):
  {cyan('curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash')}

RHEL / CentOS / Fedora:
  {cyan('sudo rpm --import https://packages.microsoft.com/keys/microsoft.asc')}
  {cyan('sudo dnf install azure-cli')}

Any Linux (script installer):
  {cyan('curl -L https://aka.ms/InstallAzureCli | bash')}
""")

    if ask_yn("Have you installed the Azure CLI and want to retry?", default="n"):
        rc2, out2, _ = run_cmd("az --version")
        if rc2 == 0:
            ok("Azure CLI found — continuing")
            return
        fail("Azure CLI still not found. Open a new terminal after installing, then re-run.")
        sys.exit(1)

    info("Install the Azure CLI, then re-run this wizard.")
    sys.exit(0)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 — Python packages
# ─────────────────────────────────────────────────────────────────────────────
def install_packages(total):
    step_header(3, total, "Python packages",
                "Installing the Azure SDK management libraries and Excel/progress dependencies.")

    req = os.path.join(DIR, "requirements.txt")
    if not os.path.exists(req):
        fail(f"requirements.txt not found in {DIR}")
        fail("Make sure you are running this wizard from the assessment folder.")
        sys.exit(1)

    # Quick check — validate all required packages before offering to skip
    rc, _, _ = run_cmd(
        f'"{sys.executable}" -c "'
        "import openpyxl, tqdm, azure.identity, azure.mgmt.resource, azure.mgmt.compute, "
        "azure.mgmt.storage, azure.mgmt.sql, azure.mgmt.containerservice, "
        "azure.mgmt.web, azure.mgmt.cosmosdb, azure.mgmt.recoveryservices, "
        "azure.mgmt.recoveryservicesbackup, azure.mgmt.redis, azure.mgmt.netapp, "
        "azure.mgmt.desktopvirtualization, azure.mgmt.synapse, "
        "azure.mgmt.containerinstance, azure.mgmt.monitor, "
        'azure.mgmt.sqlvirtualmachine, azure.mgmt.costmanagement"',
        capture=True,
    )
    if rc == 0:
        ok("All required packages already installed")
        if not ask_yn("Re-install / upgrade packages anyway?", default="n"):
            return

    blank()
    info(f"Running:  pip install -r requirements.txt  (this may take 1–2 minutes)")
    blank()
    hr()
    rc = run_live(f'"{sys.executable}" -m pip install -r "{req}" --upgrade')
    hr()
    blank()

    if rc != 0:
        # PEP 668: Homebrew/system Python blocks pip — retry with --break-system-packages
        rc2, err2, _ = run_cmd(
            f'"{sys.executable}" -m pip install -r "{req}" --upgrade --break-system-packages 2>&1',
            capture=True,
        )
        if rc2 == 0:
            rc = 0
            ok("Packages installed with --break-system-packages (PEP 668 environment)")

    if rc == 0:
        ok("All packages installed successfully")
    else:
        warn("pip reported one or more errors (see above).")
        note("Common fixes:")
        note("  • Try:  python -m pip install --upgrade pip  then re-run")
        note("  • On Linux you may need:  sudo apt install python3-dev")
        if not ask_yn("Continue anyway?", default="n"):
            sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 — Authentication
# ─────────────────────────────────────────────────────────────────────────────
def setup_auth(total):
    step_header(4, total, "Azure authentication",
                "The tool is 100% read-only — it only calls List/Get APIs.")

    blank()
    note("Your credentials are passed directly to the Azure SDK.")
    note("They are never printed, logged, or sent anywhere except Azure API endpoints.")
    blank()

    # Already logged in?
    rc, out, _ = run_cmd("az account show --output json")
    if rc == 0:
        try:
            acct = json.loads(out)
            ok(f"Already signed in as:  {bold(acct.get('user', {}).get('name', '?'))}")
            ok(f"Active subscription:   {bold(acct.get('name', '?'))}  {dim('(' + acct.get('id','') + ')')}")
            blank()
            if ask_yn("Use this account and continue?"):
                return "cli"
        except Exception:
            pass

    choice, label = ask_choice(
        "How would you like to authenticate?",
        [
            ("Interactive browser login  —  simplest for your own account", "cli"),
            ("Service principal          —  for customer / automation environments", "sp"),
            ("Environment variables already set  (AZURE_CLIENT_ID etc.)", "env"),
        ],
        default=1,
    )

    blank()
    if choice == "cli":
        _auth_interactive()
    elif choice == "sp":
        _auth_service_principal()
    else:
        _auth_env_vars()

    return choice


def _auth_interactive():
    info("Opening your browser for Azure login...")
    note("If a browser does not open automatically, copy the device code shown")
    note("and visit:  https://microsoft.com/devicelogin")
    blank()
    rc = run_live("az login")
    blank()
    if rc != 0:
        fail("Login failed. Please try again or choose a different auth method.")
        sys.exit(1)
    ok("Signed in successfully")


def _auth_service_principal():
    blank()
    print(bold("  What is a service principal?"))
    blank()
    indent(f"""
A service principal is an app identity in Azure Active Directory — like a
service account. Assigning it the built-in {bold('Reader')} role on a subscription
gives it read-only access to everything the assessment needs.

{bold('The customer runs this once to create one:')}

  {cyan('az ad sp create-for-rbac')}
  {cyan('  --name "AzureAssessmentReadOnly"')}
  {cyan('  --role "Reader"')}
  {cyan('  --scopes /subscriptions/<their-subscription-id>')}

The output looks like:
  {dim('{')}
  {dim('  "appId":    "aaaaaaaa-...",')}   {green('← AZURE_CLIENT_ID')}
  {dim('  "password": "xxxxxxxx-...",')}   {green('← AZURE_CLIENT_SECRET')}
  {dim('  "tenant":   "bbbbbbbb-..."')}    {green('← AZURE_TENANT_ID')}
  {dim('}')}

For multiple subscriptions, assign Reader at the management group instead:

  {cyan('az role assignment create')}
  {cyan('  --assignee <appId>')}
  {cyan('  --role "Reader"')}
  {cyan('  --scope /providers/Microsoft.Management/managementGroups/<mg-id>')}
""")

    blank()
    info("Enter the credentials below. They are set as session environment")
    info("variables and are not written to disk.")
    blank()

    client_id     = ask("AZURE_CLIENT_ID     (appId)",    required=True)
    client_secret = ask("AZURE_CLIENT_SECRET (password)", required=True, secret=True)
    tenant_id     = ask("AZURE_TENANT_ID     (tenant)",   required=True)
    sub_id        = ask("AZURE_SUBSCRIPTION_ID (optional — leave blank to scan all accessible)")

    os.environ["AZURE_CLIENT_ID"]     = client_id
    os.environ["AZURE_CLIENT_SECRET"] = client_secret
    os.environ["AZURE_TENANT_ID"]     = tenant_id
    if sub_id:
        os.environ["AZURE_SUBSCRIPTION_ID"] = sub_id

    blank()
    info("Verifying credentials...")
    rc, out, _ = run_cmd("az account show --output json")
    if rc == 0:
        try:
            acct = json.loads(out)
            ok(f"Verified — subscription: {bold(acct.get('name','?'))}")
            return
        except Exception:
            pass
    warn("Could not verify via 'az account show' — credentials set, will attempt scan anyway.")
    note("If the scan fails with AuthorizationFailed, double-check the credentials")
    note("and confirm the Reader role is assigned on the target subscription.")


def _auth_env_vars():
    needed = ["AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET", "AZURE_TENANT_ID"]
    missing = [v for v in needed if not os.environ.get(v)]
    if missing:
        warn(f"Missing environment variables: {', '.join(missing)}")
        blank()
        note("Set them in your shell before running, for example:")
        blank()
        if OS == "Windows":
            for v in missing:
                indent(f'$env:{v} = "your-value-here"', prefix="    ")
        else:
            for v in missing:
                indent(f'export {v}="your-value-here"', prefix="    ")
        blank()
        fail("Please set the variables and re-run the wizard.")
        sys.exit(1)
    ok(f"AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID — all set")
    if os.environ.get("AZURE_SUBSCRIPTION_ID"):
        ok(f"AZURE_SUBSCRIPTION_ID — set")


# ─────────────────────────────────────────────────────────────────────────────
# STEP 5 — Subscription selection
# ─────────────────────────────────────────────────────────────────────────────
def select_subscriptions(total):
    step_header(5, total, "Subscription selection",
                "Choose which Azure subscriptions to include in the assessment.")

    blank()
    info("Fetching your accessible subscriptions...")
    rc, out, _ = run_cmd("az account list --output json --all")

    subs = []
    if rc == 0:
        try:
            subs = [s for s in json.loads(out) if s.get("state") == "Enabled"]
        except Exception:
            pass

    if not subs:
        warn("Could not list subscriptions — will use the current active subscription.")
        return [], False

    blank()
    print(bold(f"  Found {len(subs)} enabled subscription(s):"))
    blank()
    name_w = min(52, max((len(s["name"]) for s in subs), default=20))
    for i, s in enumerate(subs, 1):
        print(f"    {dim(str(i).rjust(3))}.  {s['name'][:name_w]:<{name_w}}  {dim(s['id'])}")

    choice, _ = ask_choice(
        "Which subscriptions should be scanned?",
        [
            (f"All {len(subs)} enabled subscriptions", "all"),
            ("Choose specific subscriptions by number", "pick"),
            ("Current active subscription only", "current"),
        ],
        default=1,
    )

    if choice == "all":
        blank()
        ok(f"Will scan all {len(subs)} subscriptions")
        return [], True

    if choice == "current":
        rc2, out2, _ = run_cmd("az account show --output json")
        try:
            cur = json.loads(out2)
            ok(f"Will scan: {bold(cur.get('name','current subscription'))}")
        except Exception:
            ok("Will scan the current active subscription")
        return [], False

    # pick
    blank()
    note("Enter numbers separated by commas, e.g.  1, 3, 5")
    raw = ask("Subscription numbers", required=True)
    picked_ids = []
    for part in raw.split(","):
        try:
            idx = int(part.strip()) - 1
            if 0 <= idx < len(subs):
                picked_ids.append(subs[idx]["id"])
        except ValueError:
            pass

    if not picked_ids:
        warn("No valid selections — defaulting to current active subscription")
        return [], False

    blank()
    for sid in picked_ids:
        name = next((s["name"] for s in subs if s["id"] == sid), sid)
        ok(f"Selected: {bold(name)}")
    return picked_ids, False


# ─────────────────────────────────────────────────────────────────────────────
# STEP 6 — Scan options
# ─────────────────────────────────────────────────────────────────────────────
def select_options(total):
    step_header(6, total, "Scan options",
                "Press Enter to accept the default shown in [brackets].")

    blank()

    # ── Performance ──────────────────────────────────────────────────────────
    print(f"  {bold(cyan('── Performance ──────────────────────────────────────'))}")
    blank()

    print(f"  {bold('Skip disk snapshots?')}")
    note("  Snapshot enumeration is slow on subscriptions with thousands of disks.")
    note("  Recommended: Yes for large environments or quick inventories.")
    skip_snap = ask_yn("  Skip snapshots", default="n")
    blank()

    print(f"  {bold('Parallel workers')}")
    note("  Subscriptions scanned simultaneously. Match to number of subscriptions (max 8).")
    workers_raw = ask("  Workers", default="4")
    try:
        workers = max(1, min(16, int(workers_raw)))
    except ValueError:
        workers = 4
    blank()

    # ── Output ───────────────────────────────────────────────────────────────
    print(f"  {bold(cyan('── Output ───────────────────────────────────────────'))}")
    blank()

    print(f"  {bold('Output filename')}")
    note("  Full assessment workbook saved to the current directory.")
    date_str = datetime.datetime.now().strftime("%Y%m%d")
    default_out = f"azure_assessment_{date_str}.xlsx"
    output = ask("  Filename", default=default_out)
    if not output.endswith(".xlsx"):
        output += ".xlsx"
    blank()

    print(f"  {bold('Anonymize resource names?')}")
    note("  Replaces all names (subscriptions, resource groups, VMs, etc.) with")
    note("  opaque codes — e.g. SUB-0001, RG-0001, VM-0001.")
    note("  A reversible mapping CSV is saved alongside the workbook.")
    anonymize = ask_yn("  Anonymize", default="n")
    blank()

    print(f"  {bold('Generate Veeam Scenario Builder file?')}")
    note("  Writes a second file in the CAzureWrapper import format for")
    note("  veeam.com/calculators/scenario/build/cloud/azure")
    note(f"  Output: {output.replace('.xlsx', '_scenario_builder.xlsx')}")
    scenario_builder = ask_yn("  Scenario Builder", default="n")
    blank()

    # ── Troubleshooting ───────────────────────────────────────────────────────
    print(f"  {bold(cyan('── Troubleshooting ──────────────────────────────────'))}")
    blank()

    print(f"  {bold('Verbose logging?')}")
    note("  Prints detailed per-service log lines — useful for first runs or debugging.")
    verbose = ask_yn("  Verbose", default="n")
    blank()

    # ── Summary ───────────────────────────────────────────────────────────────
    hr()
    blank()
    print(bold("  Your settings:"))
    blank()
    ok(f"Skip snapshots     :  {'Yes' if skip_snap else 'No'}")
    ok(f"Workers            :  {workers}")
    ok(f"Output file        :  {output}")
    ok(f"Anonymize          :  {'Yes' if anonymize else 'No'}")
    ok(f"Scenario Builder   :  {'Yes  →  ' + output.replace('.xlsx', '_scenario_builder.xlsx') if scenario_builder else 'No'}")
    ok(f"Verbose logging    :  {'Yes' if verbose else 'No'}")

    return {
        "skip_snapshots":   skip_snap,
        "workers":          workers,
        "output":           output,
        "anonymize":        anonymize,
        "scenario_builder": scenario_builder,
        "verbose":          verbose,
    }


# ─────────────────────────────────────────────────────────────────────────────
# STEP 7 — Run
# ─────────────────────────────────────────────────────────────────────────────
def run_assessment(sub_ids, all_subs, opts, total):
    step_header(7, total, "Run the assessment",
                "Everything is ready. The scan is 100% read-only.")

    script = os.path.join(DIR, "azure_assessment.py")
    if not os.path.exists(script):
        fail(f"azure_assessment.py not found in {DIR}")
        fail("Make sure this wizard is in the same folder as the assessment script.")
        sys.exit(1)

    # Build arg list — never shell=True so user-supplied filenames can't inject commands
    argv = [sys.executable, script]

    if all_subs:
        argv.append("--all-subscriptions")
    elif sub_ids:
        argv += ["--subscription"] + list(sub_ids)

    if opts["skip_snapshots"]:
        argv.append("--skip-snapshots")
    argv += ["--workers", str(opts["workers"])]
    argv += ["--output", opts["output"]]
    if opts["anonymize"]:
        argv.append("--anonymize")
    if opts.get("scenario_builder"):
        argv.append("--scenario-builder")
    if opts["verbose"]:
        argv.append("--verbose")

    # Display string for the user (cosmetic only — not passed to shell)
    cmd = " ".join(
        f'"{a}"' if " " in a else a
        for a in argv
    )

    blank()
    print(bold("  Command that will be run:"))
    blank()
    indent(cyan(cmd))
    blank()
    note("This may take a few minutes depending on the size of the environment.")
    note("A progress bar will appear for each subscription being scanned.")
    blank()

    if not ask_yn("Start the scan now?"):
        blank()
        info("Run it manually when ready:")
        blank()
        indent(cyan(cmd))
        blank()
        sys.exit(0)

    blank()
    hr("─")
    blank()
    rc = subprocess.run(argv).returncode
    blank()
    hr("─")
    blank()

    if rc == 0:
        ok(bold("Assessment complete!"))
        ok(f"Output saved to:  {bold(opts['output'])}")
        blank()
        _open_output(opts["output"])
    else:
        fail("The assessment exited with errors — review the output above.")
        blank()
        print(bold("  Common fixes:"))
        blank()
        indent(f"""
{yellow('AuthorizationFailed')}
  The credential does not have Reader access on that subscription.
  Ask the customer to run:
    {cyan('az role assignment create --assignee <appId> --role Reader --scope /subscriptions/<id>')}

{yellow('DefaultAzureCredential failed')}
  No credentials configured. Run  {cyan('az login')}  or set service principal env vars.

{yellow('ResourceNotFoundError')}
  A service is not enabled in this subscription — normal, the scan skips it.

{yellow('ModuleNotFoundError')}
  A package is missing. Re-run the wizard and let it install packages again.

Use  {cyan('--verbose')}  to see detailed per-service error messages.
""")


def _open_output(path):
    if not os.path.exists(path):
        return
    blank()
    if ask_yn(f"Open {path} now?"):
        if OS == "Windows":
            os.startfile(path)
        elif OS == "Darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────
def main():
    TOTAL = 7
    banner()

    print(bold("  What this wizard does:"))
    blank()
    indent(f"""
{cyan('Step 1')}  Confirm Python 3.10+ is installed
{cyan('Step 2')}  Check for the Azure CLI (with install instructions if missing)
{cyan('Step 3')}  Install Python package dependencies from requirements.txt
{cyan('Step 4')}  Set up Azure authentication (interactive, service principal, or env vars)
{cyan('Step 5')}  Choose which subscriptions to scan
{cyan('Step 6')}  Configure scan options (skip snapshots, workers, output name, anonymize, scenario builder)
{cyan('Step 7')}  Run the assessment and open the Excel workbook
""")

    print(bold("  What it does NOT do:"))
    blank()
    indent(f"""
{green('✓')}  The assessment is 100% read-only — no resources are created, modified, or deleted
{green('✓')}  Credentials are passed directly to the Azure SDK and never written to disk
{green('✓')}  No data leaves your machine except API calls to official Azure endpoints
{green('✓')}  No telemetry, no call-home, no third-party network traffic
""")

    if not ask_yn("Ready to begin?"):
        sys.exit(0)

    check_python(TOTAL)
    check_az_cli(TOTAL)
    install_packages(TOTAL)
    setup_auth(TOTAL)
    sub_ids, all_subs = select_subscriptions(TOTAL)
    opts = select_options(TOTAL)
    run_assessment(sub_ids, all_subs, opts, TOTAL)


if __name__ == "__main__":
    main()
