#!/usr/bin/env bash
# Optional: build a single standalone program, "JJs Setlist", that runs
# without Python installed. Most people don't need this - run.sh is enough.
#
#   ./build.sh
#
# It must be built on Linux (a Windows PC can't build it), and a program built
# on one distribution runs on that one and newer ones - build on the oldest
# system you want to support.

set -u
cd "$(dirname "$(readlink -f "$0")")" || exit 1
APPNAME="JJs Setlist"
WORK="${TMPDIR:-/tmp}/setlist_build"
BUILD_VENV=".venv-build"

fail() {
    printf '\n*** Build FAILED - %b ***\n\n' "$1" >&2
    exit 1
}

echo "=================================================="
echo "  JJ's Musicians Setlist Organiser - build \"$APPNAME\""
echo "=================================================="
echo

echo "[1/3] Checking build tools (first time: downloads PyInstaller, openpyxl, certifi)..."
bash ./run.sh --setup >/dev/null || fail "run ./run.sh first and fix what it reports."
if [ ! -x "$BUILD_VENV/bin/python" ]; then
    python3 -m venv "$BUILD_VENV" || fail "could not make a build environment (see README: python3-venv)."
fi
"$BUILD_VENV/bin/python" -m pip install --quiet --disable-pip-version-check \
    pyinstaller -r requirements.txt || fail "could not install the build tools."

echo "[2/3] Compiling (this takes a minute or two)..."
rm -rf "$WORK"
"$BUILD_VENV/bin/python" -m PyInstaller --noconfirm --clean --log-level WARN \
    --onefile --windowed --name "$APPNAME" \
    --hidden-import openpyxl --hidden-import certifi \
    --distpath "$WORK/dist" --workpath "$WORK/build" --specpath "$WORK" \
    "$PWD/jjs_setlist.py" || fail "PyInstaller reported an error - see the messages above."

echo "[3/3] Finishing..."
mv -f "$WORK/dist/$APPNAME" . || fail "could not move the finished program into this folder."
rm -rf "$WORK"

echo
echo "=================================================="
echo "  Done!  \"$APPNAME\" is in:"
echo "  $PWD"
echo
echo "  Run it with:  ./\"$APPNAME\""
echo "  Your settings and saved setlists are kept in:"
echo "  ~/Documents/JJs Setlist"
echo "=================================================="
