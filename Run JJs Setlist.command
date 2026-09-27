#!/bin/bash
# Run JJ's Musicians Setlist Organiser straight from jjs_setlist.py - no build needed.
# Double-click this file. Keep this Terminal window open while using the app
# (closing it closes the app); any error messages appear here.

cd "$(dirname "$0")" || exit 1

pause_and_exit() {
    echo
    read -n 1 -s -r -p "Press any key to close..."
    echo
    exit 1
}

if ! command -v python3 >/dev/null 2>&1; then
    echo "Python 3 was not found."
    echo "Install it from https://www.python.org/downloads/macos/"
    pause_and_exit
fi
if ! python3 -c "import tkinter" >/dev/null 2>&1; then
    echo "This Python has no Tk. Install Python from python.org (it includes Tk)."
    pause_and_exit
fi
if ! python3 -c "import openpyxl, certifi" >/dev/null 2>&1; then
    echo "First run: installing openpyxl and certifi (Excel files, secure downloads)..."
    python3 -m pip install --quiet openpyxl certifi || pause_and_exit
fi

echo "JJ's Musicians Setlist Organiser is running."
echo "Keep this window open while you use it - closing it closes the app."
python3 jjs_setlist.py || pause_and_exit
