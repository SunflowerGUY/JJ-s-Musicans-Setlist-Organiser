#!/usr/bin/env bash
# Run JJ's Musicians Setlist Organiser straight from this folder - no build needed.
#
#   ./run.sh           start the app (the first run sets up what it needs)
#   ./run.sh --setup   only check and set up what it needs
#
# Needs Python 3.10+ with Tk, plus openpyxl and certifi. If those two aren't
# installed system-wide, they're installed into a private ".venv" folder here
# (modern distributions don't allow pip to install into the system Python).

set -u
cd "$(dirname "$(readlink -f "$0")")" || exit 1

APP="JJ's Musicians Setlist Organiser"
VENV=".venv"
LOG="${XDG_CACHE_HOME:-$HOME/.cache}/jjs-setlist.log"

# Show an error in the terminal, or in a pop-up when started from the menu.
fail() {
    printf '\n%b\n\n' "$1" >&2
    if [ ! -t 2 ]; then
        if command -v zenity >/dev/null 2>&1; then
            zenity --error --no-markup --title="$APP" --text="$(printf '%b' "$1")"
        elif command -v kdialog >/dev/null 2>&1; then
            kdialog --title "$APP" --error "$(printf '%b' "$1")"
        elif command -v notify-send >/dev/null 2>&1; then
            notify-send "$APP" "$(printf '%b' "$1")"
        fi
    fi
    exit 1
}

# The install command for a package on this distribution.
pkg_cmd() {
    local apt dnf pac zyp
    case "$1" in
        python)   apt=python3          dnf=python3          pac=python          zyp=python3 ;;
        tk)       apt=python3-tk       dnf=python3-tkinter  pac=tk              zyp=python3-tk ;;
        venv)     apt=python3-venv     dnf=python3          pac=python          zyp=python3 ;;
        openpyxl) apt="python3-openpyxl python3-certifi"
                  dnf="python3-openpyxl python3-certifi"
                  pac="python-openpyxl python-certifi"
                  zyp="python3-openpyxl python3-certifi" ;;
    esac
    if   command -v apt-get >/dev/null 2>&1; then echo "sudo apt install $apt"
    elif command -v dnf     >/dev/null 2>&1; then echo "sudo dnf install $dnf"
    elif command -v pacman  >/dev/null 2>&1; then echo "sudo pacman -S $pac"
    elif command -v zypper  >/dev/null 2>&1; then echo "sudo zypper install $zyp"
    else echo "install your distribution's '$apt' package"
    fi
}

# Pick the Python to run with (sets PY), setting it up the first time.
choose_python() {
    command -v python3 >/dev/null 2>&1 ||
        fail "Python 3 was not found. Install it with:\n    $(pkg_cmd python)"
    python3 -c 'import sys; sys.exit(sys.version_info < (3, 10))' ||
        fail "$APP needs Python 3.10 or newer (this is $(python3 -V 2>&1))."
    python3 -c 'import tkinter' >/dev/null 2>&1 ||
        fail "Python's window toolkit (Tk) is missing. Install it with:\n    $(pkg_cmd tk)"

    if python3 -c 'import openpyxl, certifi' >/dev/null 2>&1; then
        PY=python3
        return
    fi
    if ! "$VENV/bin/python" -c 'import openpyxl, certifi' >/dev/null 2>&1; then
        echo "First run: installing openpyxl and certifi into $VENV (needs internet)..."
        rm -rf "$VENV"
        python3 -m venv "$VENV" >/dev/null 2>&1 || {
            rm -rf "$VENV"
            fail "Could not set up the add-ons. Either install them with:\n    $(pkg_cmd openpyxl)\nor allow a private copy with:\n    $(pkg_cmd venv)\nthen run this again."
        }
        "$VENV/bin/python" -m pip install --quiet --disable-pip-version-check \
            -r requirements.txt ||
            fail "Could not download openpyxl and certifi - check the internet connection, or install them with:\n    $(pkg_cmd openpyxl)"
    fi
    PY="$VENV/bin/python"
}

choose_python
if [ "${1:-}" = "--setup" ]; then
    echo "All set - Python, Tk, openpyxl and certifi are ready."
    exit 0
fi

if [ -t 2 ]; then
    "$PY" jjs_setlist.py || fail "The app closed with an error - see the messages above."
else
    # Started from the applications menu: keep any error messages in a log.
    mkdir -p "$(dirname "$LOG")"
    "$PY" jjs_setlist.py 2>"$LOG" ||
        fail "The app closed with an error. The details are in:\n    $LOG"
fi
