#!/usr/bin/env bash
# Azure Environment Assessment — macOS / Linux Launcher
# Installs Python if needed, then launches the interactive setup wizard.
#
# Usage:
#   chmod +x start-assessment.sh
#   ./start-assessment.sh

set -euo pipefail

# ── colours ───────────────────────────────────────────────────────────────────
if [ -t 1 ]; then
    CYAN='\033[96m' GREEN='\033[92m' YELLOW='\033[93m' RED='\033[91m' RESET='\033[0m'
else
    CYAN='' GREEN='' YELLOW='' RED='' RESET=''
fi

header() { echo -e "\n${CYAN}=================================================${RESET}"; \
           echo -e "${CYAN}   Azure Environment Assessment — Launcher${RESET}"; \
           echo -e "${CYAN}=================================================${RESET}\n"; }
step()   { echo -e "${CYAN}[*] $*${RESET}"; }
ok()     { echo -e "${GREEN}[OK] $*${RESET}"; }
warn()   { echo -e "${YELLOW}[!] $*${RESET}"; }
fail()   { echo -e "${RED}[X] $*${RESET}"; }

# ── detect OS ─────────────────────────────────────────────────────────────────
OS="$(uname -s)"   # Darwin | Linux

# ── find python 3.10+ ─────────────────────────────────────────────────────────
find_python() {
    for cmd in python3 python python3.13 python3.12 python3.11 python3.10; do
        if command -v "$cmd" &>/dev/null; then
            ver="$($cmd --version 2>&1 | sed -E 's/Python ([0-9]+\.[0-9]+).*/\1/')"
            maj="${ver%%.*}"
            min="${ver##*.}"
            if [ "$maj" -gt 3 ] 2>/dev/null || { [ "$maj" -eq 3 ] 2>/dev/null && [ "$min" -ge 10 ] 2>/dev/null; }; then
                echo "$cmd"
                return 0
            fi
        fi
    done
    return 1
}

# ── install python on macOS ───────────────────────────────────────────────────
install_macos() {
    # Try Homebrew first
    if command -v brew &>/dev/null; then
        warn "Python 3.10+ not found. Installing via Homebrew..."
        brew install python@3.12
        return 0
    fi

    # No Homebrew — offer to install it
    warn "Python 3.10+ not found and Homebrew is not installed."
    echo ""
    echo "  Option A — Install Homebrew (recommended, installs Python too):"
    echo "    Paste this in your terminal:"
    echo "    /bin/bash -c \"\$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\""
    echo "    Then run:  brew install python@3.12"
    echo ""
    echo "  Option B — Download the official Python installer:"
    echo "    https://www.python.org/downloads/macos/"
    echo "    Download and run the .pkg file, then re-run this script."
    echo ""
    read -rp "  Press Enter to open the Python download page in your browser, or Ctrl-C to cancel: "
    open "https://www.python.org/downloads/macos/" 2>/dev/null || true
    exit 1
}

# ── install python on Linux ───────────────────────────────────────────────────
install_linux() {
    warn "Python 3.10+ not found. Attempting to install..."
    echo ""

    if command -v apt-get &>/dev/null; then
        echo "  Detected Debian/Ubuntu — running: sudo apt-get install -y python3 python3-pip"
        sudo apt-get update -qq && sudo apt-get install -y python3 python3-pip
    elif command -v dnf &>/dev/null; then
        echo "  Detected Fedora/RHEL — running: sudo dnf install -y python3 python3-pip"
        sudo dnf install -y python3 python3-pip
    elif command -v yum &>/dev/null; then
        echo "  Detected CentOS/RHEL (yum) — running: sudo yum install -y python3 python3-pip"
        sudo yum install -y python3 python3-pip
    elif command -v zypper &>/dev/null; then
        echo "  Detected openSUSE — running: sudo zypper install -y python3 python3-pip"
        sudo zypper install -y python3 python3-pip
    elif command -v pacman &>/dev/null; then
        echo "  Detected Arch Linux — running: sudo pacman -S --noconfirm python python-pip"
        sudo pacman -S --noconfirm python python-pip
    else
        fail "Could not detect a supported package manager (apt, dnf, yum, zypper, pacman)."
        echo ""
        echo "  Install Python 3.10 or later manually from https://www.python.org/downloads/"
        echo "  then re-run this script."
        exit 1
    fi
}

# ── main ──────────────────────────────────────────────────────────────────────
header

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Step 1: Find Python
step "Looking for Python 3.10 or later..."
PYTHON=""
if PYTHON="$(find_python)"; then
    VER="$($PYTHON --version 2>&1)"
    ok "Found $VER (command: $PYTHON)"
else
    # Try to install
    if [ "$OS" = "Darwin" ]; then
        install_macos
    else
        install_linux
    fi

    # Try again after install
    if PYTHON="$(find_python)"; then
        VER="$($PYTHON --version 2>&1)"
        ok "Installed and found $VER"
    else
        fail "Python installation did not complete. Please install Python 3.10+ manually and re-run."
        exit 1
    fi
fi

# Step 2: Confirm wizard exists
WIZARD="$SCRIPT_DIR/setup_wizard.py"
if [ ! -f "$WIZARD" ]; then
    fail "setup_wizard.py not found in $SCRIPT_DIR"
    echo "  Make sure you are running this script from the azure-environment-assessment directory."
    exit 1
fi

# Step 3: Hand off to wizard
ok "Launching setup wizard..."
echo ""
cd "$SCRIPT_DIR"
exec "$PYTHON" "$WIZARD" "$@"
