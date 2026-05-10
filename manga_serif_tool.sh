#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"

VENV=.venv_manga_serif
PYEXE="$VENV/bin/python"

if [ ! -f "$PYEXE" ]; then
    echo "[setup] Creating venv ($VENV)..."
    if ! python3 -m venv "$VENV"; then
        echo ""
        echo "ERROR: Python 3 not found or venv creation failed."
        echo "Install Python 3 from https://www.python.org/"
        exit 1
    fi
fi

if ! "$PYEXE" -c "import PyQt6" 2>/dev/null; then
    echo "[setup] Installing PyQt6 (first run only)..."
    "$PYEXE" -m pip install --upgrade pip
    if ! "$PYEXE" -m pip install PyQt6; then
        echo ""
        echo "ERROR: PyQt6 install failed."
        exit 1
    fi
fi

echo "[run] Launching manga serif tool..."
if ! "$PYEXE" "$(dirname "$0")/manga_serif_tool.py"; then
    EXITCODE=$?
    echo ""
    echo "ERROR: tool exited with code $EXITCODE"
    echo "See log: $(dirname "$0")/manga_serif_tool.log"
    exit $EXITCODE
fi
