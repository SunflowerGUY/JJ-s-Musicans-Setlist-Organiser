#!/usr/bin/env bash
# Optional: build the AppImage inside an Ubuntu 22.04 container, so it runs on
# Ubuntu 22.04 / Mint 21 / Debian 12 and everything newer - the version to share.
#
#   ./build_appimage_2204.sh
#
# Needs Podman (Ubuntu/Mint: sudo apt install podman). The first run downloads
# Ubuntu 22.04 and the build tools (a few hundred MB, kept by Podman for next
# time), and later runs reuse them. Only the app files go into the container -
# never your setlists, spreadsheet, database_link.txt or config.json.
# The finished JJs_Setlist-x86_64.AppImage replaces the one in this folder.

set -u
cd "$(dirname "$(readlink -f "$0")")" || exit 1
IMAGE="jjs-setlist-build:22.04"
OUTPUT="JJs_Setlist-x86_64.AppImage"
OUT="$(mktemp -d "${TMPDIR:-/tmp}/setlist_2204.XXXXXX")" || exit 1
trap 'rm -rf "$OUT"' EXIT

fail() {
    printf '\n*** Build FAILED - %b ***\n\n' "$1" >&2
    exit 1
}

echo "=================================================="
echo "  JJ's Musicians Setlist Organiser - AppImage on Ubuntu 22.04"
echo "=================================================="
echo

echo "[1/3] Preparing Ubuntu 22.04 (first time: a few minutes of downloading)..."
command -v podman >/dev/null 2>&1 || fail "Podman is not installed. Install it with:\n    sudo apt install podman"
podman build --quiet --tag "$IMAGE" - >/dev/null <<'EOF' || fail "could not prepare the Ubuntu 22.04 container - see the messages above."
FROM docker.io/library/ubuntu:22.04
ENV DEBIAN_FRONTEND=noninteractive
RUN apt-get update && apt-get install -y --no-install-recommends \
        python3 python3-venv python3-tk python3-pip libpython3.10 binutils curl ca-certificates file \
    && rm -rf /var/lib/apt/lists/*
EOF

echo "[2/3] Building inside the container (a few minutes)..."
# Send in just the app files; the AppImage comes back through $OUT. The
# jjs-setlist-cache volume keeps appimagetool and the downloaded Python add-ons
# between builds (remove it with: podman volume rm jjs-setlist-cache).
tar -c jjs_setlist.py logo_data.py requirements.txt run.sh build_appimage.sh assets/jjs-setlist.png |
    podman run --rm -i --volume "$OUT:/out" --volume jjs-setlist-cache:/root/.cache "$IMAGE" bash -c '
        mkdir /work && cd /work && tar -x &&
        ./build_appimage.sh && cp JJs_Setlist-x86_64.AppImage /out/' ||
    fail "the build inside the container reported an error - see the messages above."

echo "[3/3] Finishing..."
mv -f "$OUT/$OUTPUT" "$OUTPUT" && chmod +x "$OUTPUT" ||
    fail "could not move the finished AppImage into this folder."

echo
echo "=================================================="
echo "  Done!  $OUTPUT ($(du -h "$OUTPUT" | cut -f1)) is in:"
echo "  $PWD"
echo "  Built on Ubuntu 22.04: runs on Ubuntu 22.04, Mint 21, Debian 12 and newer."
echo "=================================================="
