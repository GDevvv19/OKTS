#!/bin/bash
# Double-clickable launcher for Ship Report Builder (runs natively on any Mac).
# First run sets up a small local Python environment; later runs start instantly.
cd "$(dirname "$0")" || exit 1

echo "Starting Ship Report Builder..."

if ! command -v python3 >/dev/null 2>&1; then
  echo ""
  echo "Python 3 is required but was not found."
  echo "Install it from https://www.python.org/downloads/ (or run: xcode-select --install), then try again."
  read -r -p "Press Enter to close."
  exit 1
fi

if [ ! -d .venv ]; then
  echo "First-time setup (creating a local environment; needs internet once)..."
  python3 -m venv .venv || { echo "Could not create environment."; read -r -p "Press Enter to close."; exit 1; }
  ./.venv/bin/pip install --quiet --upgrade pip && ./.venv/bin/pip install --quiet -r requirements.txt || {
    echo "Setup failed - an internet connection is required for the first run.";
    read -r -p "Press Enter to close."; exit 1; }
fi

exec ./.venv/bin/python src/webapp.py
