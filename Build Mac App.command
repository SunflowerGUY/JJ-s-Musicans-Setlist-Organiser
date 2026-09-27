#!/bin/bash
# Build "JJs Setlist.app" on a Mac.
# Double-click this file, or in Terminal:   bash "Build Mac App.command"

cd "$(dirname "$0")" || exit 1
APPNAME="JJs Setlist"
WORK="${TMPDIR:-/tmp}/setlist_build"

fail() {
    echo
    echo "*** Build FAILED - $1 ***"
    echo
    read -n 1 -s -r -p "Press any key to close..."
    echo
    exit 1
}

is_running() { pgrep -f "$APPNAME.app/Contents/MacOS" >/dev/null 2>&1; }

echo "=================================================="
echo "  JJ's Musicians Setlist Organiser  -  build \"$APPNAME.app\""
echo "=================================================="
echo

command -v python3 >/dev/null 2>&1 ||
    fail "Python 3 was not found. Install it from https://www.python.org/downloads/macos/"
python3 -c "import tkinter" >/dev/null 2>&1 ||
    fail "this Python has no Tk. Install Python from python.org (it includes Tk)."

# --- Close the app if it's running (it can't be replaced while open) ---
if is_running; then
    echo "\"$APPNAME\" is running - closing it..."
    echo "(If it asks to save your setlist, answer within 20 seconds.)"
    osascript -e "tell application \"$APPNAME\" to quit" >/dev/null 2>&1
    for i in $(seq 1 20); do
        is_running || break
        sleep 1
    done
    if is_running; then
        echo "Still running - forcing it to close..."
        pkill -9 -f "$APPNAME.app/Contents/MacOS"
        sleep 2
    fi
    echo "      Closed."
    echo
fi

echo "[1/3] Checking build tools (first time: downloads PyInstaller, openpyxl, Pillow, certifi)..."
python3 -m pip install --quiet --upgrade pyinstaller openpyxl pillow certifi ||
    fail "could not install the build tools - see the messages above."

echo "[2/3] Compiling (this takes a minute or two)..."
ICON=()
if [ -f "ARTWORK/JJ SETLIST Icon Source.png" ]; then
    python3 make_icons.py || fail "could not make the icons."
    ICON=(--icon "$PWD/ARTWORK/JJ SETLIST Icon Source.png")
    echo "      Using icon: ARTWORK/JJ SETLIST Icon Source.png"
fi
rm -rf "$WORK"
python3 -m PyInstaller --noconfirm --clean --log-level WARN \
    --windowed --name "$APPNAME" \
    --osx-bundle-identifier "com.jjssetlist.app" \
    "${ICON[@]}" --hidden-import openpyxl --hidden-import certifi \
    --distpath "$WORK/dist" --workpath "$WORK/build" --specpath "$WORK" \
    "$PWD/jjs_setlist.py" || fail "PyInstaller reported an error - see the messages above."

echo "[3/3] Finishing..."
rm -rf "$APPNAME.app"
mv "$WORK/dist/$APPNAME.app" . || fail "could not move the finished app into this folder."
rm -rf "$WORK"
xattr -cr "$APPNAME.app" 2>/dev/null

echo
echo "=================================================="
echo "  Done!  \"$APPNAME.app\" is in:"
echo "  $PWD"
echo
echo "  Drag it into your Applications folder (optional)."
echo "  Your settings and saved setlists are kept in:"
echo "  Documents/JJs Setlist"
echo "=================================================="
echo
read -n 1 -s -r -p "Press any key to close..."
echo
