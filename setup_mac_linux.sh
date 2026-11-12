#!/usr/bin/env bash
# jobpilot setup for macOS and Linux
set -e
cd "$(dirname "$0")"
echo "============================================"
echo "  jobpilot setup for macOS / Linux"
echo "============================================"

PY="$(command -v python3 || command -v python || true)"
if [ -z "$PY" ]; then
  echo "Python was not found. Install Python 3.10+ (https://www.python.org/downloads/ or 'brew install python')."
  exit 1
fi
"$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' || {
  echo "Python 3.10 or newer is required (found: $("$PY" --version))."; exit 1; }

echo "[1/3] Creating a private Python environment in this folder (.venv)..."
[ -x .venv/bin/python ] || "$PY" -m venv .venv
echo "[2/3] Installing the libraries JobPilot needs (about a minute)..."
.venv/bin/python -m pip install --upgrade pip --quiet
.venv/bin/python -m pip install -r requirements.txt --quiet
chmod +x jobpilot
echo "[3/3] Guided setup"
echo
.venv/bin/python jobpilot.py init
echo
echo "Setup finished. From this folder run, for example:"
echo "   ./jobpilot tailor \"https://www.linkedin.com/jobs/view/1234567890/\""
