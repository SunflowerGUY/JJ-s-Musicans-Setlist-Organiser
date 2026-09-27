#!/usr/bin/env bash
# Add JJ's Musicians Setlist Organiser to your applications menu.
#
#   ./install.sh              set up what the app needs, then add the menu entry
#   ./install.sh --uninstall  remove the menu entry again
#
# The menu entry runs run.sh from this folder, so keep the folder where it is
# (if you move it, run ./install.sh again). Your setlists and settings live in
# ~/Documents/JJs Setlist and are never touched by this script.

set -u
HERE="$(dirname "$(readlink -f "$0")")"
APPS="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
ENTRY="$APPS/jjs-setlist.desktop"

if [ "${1:-}" = "--uninstall" ]; then
    rm -f "$ENTRY"
    update-desktop-database "$APPS" >/dev/null 2>&1
    echo "Removed from the applications menu."
    echo "Your setlists and settings in ~/Documents/JJs Setlist are still there."
    exit 0
fi

echo "[1/2] Checking what the app needs..."
bash "$HERE/run.sh" --setup || exit 1

echo "[2/2] Adding it to the applications menu..."
chmod +x "$HERE/run.sh" "$HERE/install.sh" "$HERE/build.sh" 2>/dev/null
mkdir -p "$APPS"
# In a .desktop file, quotes and $ in paths must be escaped.
esc() { printf '%s' "$1" | sed -e 's/[\\"`$]/\\&/g'; }
cat > "$ENTRY" <<EOF
[Desktop Entry]
Type=Application
Name=JJ's Musicians Setlist Organiser
GenericName=Setlist Organiser
Comment=Build, save and print gig setlists from your song spreadsheet
Exec=bash "$(esc "$HERE/run.sh")"
Icon=$HERE/assets/jjs-setlist.png
Terminal=false
Categories=AudioVideo;Audio;Music;
Keywords=setlist;songs;gig;band;music;
StartupWMClass=Jjs_setlist
EOF
update-desktop-database "$APPS" >/dev/null 2>&1

echo
echo "Done! Find \"JJ's Musicians Setlist Organiser\" in your applications menu"
echo "(it may take a few seconds to appear, or a log out and back in)."
