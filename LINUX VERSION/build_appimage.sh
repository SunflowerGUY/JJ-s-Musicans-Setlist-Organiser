#!/usr/bin/env bash
# Optional: build "JJs_Setlist-x86_64.AppImage" - one file that runs on most
# Linux systems with nothing to install. Copy it anywhere, mark it executable
# (right-click > Properties > Permissions, or chmod +x) and double-click it.
#
#   ./build_appimage.sh
#
# The finished AppImage can also put itself in the applications menu:
#
#   ./JJs_Setlist-x86_64.AppImage --install    copy to ~/Applications + menu entry
#   ./JJs_Setlist-x86_64.AppImage --uninstall  remove both again
#
# It must be built on Linux, and runs on the distribution it was built on and
# newer ones, never older - so for sharing, build on an older system such as
# Ubuntu 22.04. The first build downloads appimagetool (about 10 MB) from
# github.com/AppImage/appimagetool and keeps it in ~/.cache/jjs-setlist-build.
#
# Only the app goes in: your setlists, spreadsheet, database_link.txt and
# config.json are never included.

set -u
cd "$(dirname "$(readlink -f "$0")")" || exit 1
APPNAME="JJs_Setlist"
OUTPUT="$APPNAME-x86_64.AppImage"
WORK="${TMPDIR:-/tmp}/setlist_appimage"
APPDIR="$WORK/AppDir"
BUILD_VENV=".venv-build"
TOOLS="${XDG_CACHE_HOME:-$HOME/.cache}/jjs-setlist-build"
TOOL_URL="https://github.com/AppImage/appimagetool/releases/download/continuous/appimagetool-x86_64.AppImage"

fail() {
    printf '\n*** Build FAILED - %b ***\n\n' "$1" >&2
    exit 1
}

echo "=================================================="
echo "  JJ's Musicians Setlist Organiser - build the AppImage"
echo "=================================================="
echo

[ "$(uname -m)" = "x86_64" ] || fail "this builds a 64-bit PC (x86_64) AppImage, but this machine is $(uname -m)."

echo "[1/4] Checking build tools (first time: downloads PyInstaller, openpyxl, certifi, appimagetool)..."
bash ./run.sh --setup >/dev/null || fail "run ./run.sh first and fix what it reports."
if [ ! -x "$BUILD_VENV/bin/python" ]; then
    python3 -m venv "$BUILD_VENV" || fail "could not make a build environment (see README: python3-venv)."
fi
"$BUILD_VENV/bin/python" -m pip install --quiet --disable-pip-version-check \
    pyinstaller -r requirements.txt || fail "could not install the build tools."
if command -v appimagetool >/dev/null 2>&1; then
    TOOL="$(command -v appimagetool)"
else
    TOOL="$TOOLS/appimagetool-x86_64.AppImage"
    if [ ! -x "$TOOL" ]; then
        echo "      Downloading appimagetool from github.com/AppImage/appimagetool..."
        mkdir -p "$TOOLS"
        if command -v curl >/dev/null 2>&1; then
            curl -fL --progress-bar -o "$TOOL.part" "$TOOL_URL"
        else
            wget -q --show-progress -O "$TOOL.part" "$TOOL_URL"
        fi || { rm -f "$TOOL.part"; fail "could not download appimagetool - check the internet connection."; }
        mv -f "$TOOL.part" "$TOOL" && chmod +x "$TOOL"
    fi
fi

echo "[2/4] Compiling (this takes a minute or two)..."
rm -rf "$WORK"
mkdir -p "$WORK" || fail "could not create $WORK."
# The compiled program points LD_LIBRARY_PATH at its own libraries. Undo that
# for the programs it starts (lp for printing, xdg-open, the PDF viewer), so
# they use the system's libraries and don't crash on other distributions.
cat > "$WORK/restore_env.py" <<'EOF'
import os
_orig = os.environ.pop("LD_LIBRARY_PATH_ORIG", None)
if _orig:
    os.environ["LD_LIBRARY_PATH"] = _orig
else:
    os.environ.pop("LD_LIBRARY_PATH", None)
EOF
# --onedir, not --onefile: the AppImage is already one compressed file, and a
# onefile program inside it would unpack itself a second time at every start.
# The program is called jjs_setlist so its window class is "Jjs_setlist", the
# same as when run from source, which keeps the taskbar icon matched up.
"$BUILD_VENV/bin/python" -m PyInstaller --noconfirm --clean --log-level WARN \
    --onedir --windowed --name jjs_setlist \
    --hidden-import openpyxl --hidden-import certifi \
    --runtime-hook "$WORK/restore_env.py" \
    --distpath "$WORK/dist" --workpath "$WORK/build" --specpath "$WORK" \
    "$PWD/jjs_setlist.py" || fail "PyInstaller reported an error - see the messages above."

echo "[3/4] Assembling the AppDir..."
mkdir -p "$APPDIR/usr/lib" || fail "could not create $APPDIR."
mv "$WORK/dist/jjs_setlist" "$APPDIR/usr/lib/jjs-setlist" || fail "could not move the compiled program."
cp assets/jjs-setlist.png "$APPDIR/jjs-setlist.png" || fail "assets/jjs-setlist.png is missing."
ln -s jjs-setlist.png "$APPDIR/.DirIcon"
cat > "$APPDIR/jjs-setlist.desktop" <<'EOF'
[Desktop Entry]
Type=Application
Name=JJ's Musicians Setlist Organiser
GenericName=Setlist Organiser
Comment=Build, save and print gig setlists from your song spreadsheet
Exec=jjs-setlist
Icon=jjs-setlist
Terminal=false
Categories=AudioVideo;Audio;Music;
Keywords=setlist;songs;gig;band;music;
StartupWMClass=Jjs_setlist
EOF
cat > "$APPDIR/AppRun" <<'EOF'
#!/bin/sh
# Start JJ's Musicians Setlist Organiser from inside the AppImage, or with
# --install / --uninstall, add it to (or remove it from) the applications menu.
# It uses the same menu entry as install.sh, so the two replace each other.
HERE="$(dirname "$(readlink -f "$0")")"
DEST="$HOME/Applications/JJs_Setlist-x86_64.AppImage"
APPS="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
ENTRY="$APPS/jjs-setlist.desktop"
ICON="${XDG_DATA_HOME:-$HOME/.local/share}/icons/jjs-setlist.png"

fail() {
    printf '\n%s\n\n' "$1" >&2
    exit 1
}

case "${1:-}" in
--install)
    [ -n "${APPIMAGE:-}" ] && [ -f "$APPIMAGE" ] ||
        fail "--install only works when run as the .AppImage file."
    echo "[1/2] Copying the app to ~/Applications..."
    mkdir -p "$(dirname "$DEST")" || fail "Could not create ~/Applications."
    if [ "$(readlink -f "$APPIMAGE")" != "$(readlink -f "$DEST")" ]; then
        # Copy then rename, so an older copy that's still running is replaced safely.
        cp -f "$APPIMAGE" "$DEST.new" && chmod +x "$DEST.new" && mv -f "$DEST.new" "$DEST" ||
            { rm -f "$DEST.new"; fail "Could not copy the app to ~/Applications."; }
    fi
    echo "[2/2] Adding it to the applications menu..."
    mkdir -p "$APPS" "$(dirname "$ICON")"
    cp -f "$HERE/jjs-setlist.png" "$ICON" || fail "Could not copy the app's icon."
    # In a .desktop file, quotes and $ in paths must be escaped.
    esc() { printf '%s' "$1" | sed -e 's/[\\"`$]/\\&/g'; }
    sed -e "s|^Exec=.*|Exec=\"$(esc "$DEST")\"|" -e "s|^Icon=.*|Icon=$ICON|" \
        "$HERE/jjs-setlist.desktop" > "$ENTRY" || fail "Could not add the menu entry."
    update-desktop-database "$APPS" >/dev/null 2>&1
    echo
    echo "Done! Find \"JJ's Musicians Setlist Organiser\" in your applications menu"
    echo "(it may take a few seconds to appear, or a log out and back in)."
    echo "The app itself is in ~/Applications - you can delete the downloaded copy."
    exit 0 ;;
--uninstall)
    rm -f "$ENTRY" "$ICON" "$DEST"
    update-desktop-database "$APPS" >/dev/null 2>&1
    echo "Removed the app from ~/Applications and the applications menu."
    echo "Your setlists and settings in ~/Documents/JJs Setlist are still there."
    exit 0 ;;
esac
exec "$HERE/usr/lib/jjs-setlist/jjs_setlist" "$@"
EOF
chmod +x "$APPDIR/AppRun"

echo "[4/4] Packing it into one file..."
rm -f "$OUTPUT"
# --appimage-extract-and-run lets appimagetool itself run without FUSE.
ARCH=x86_64 "$TOOL" --appimage-extract-and-run --no-appstream "$APPDIR" "$OUTPUT" >"$WORK/appimagetool.log" 2>&1 ||
    { cat "$WORK/appimagetool.log" >&2; fail "appimagetool reported an error - see the messages above."; }
chmod +x "$OUTPUT"
rm -rf "$WORK"

echo
echo "=================================================="
echo "  Done!  $OUTPUT is in:"
echo "  $PWD"
echo "  ($(du -h "$OUTPUT" | cut -f1), built on $(. /etc/os-release && echo "$PRETTY_NAME"))"
echo
echo "  Run it with:  ./$OUTPUT"
echo "  Add it to the applications menu with:  ./$OUTPUT --install"
echo "  Your settings and saved setlists are kept in:"
echo "  ~/Documents/JJs Setlist"
echo "=================================================="
