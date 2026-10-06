"""
Refresh the "LINUX VERSION" folder with the latest app files and zip it up
for sending to a Linux user.  Run:  python make_linux_package.py
(or Make Linux Package.bat)

"LINUX VERSION" is a small repository of its own. The Linux-only files in it
(README.md, run.sh, install.sh, build.sh, requirements.txt, .gitignore) are
edited there directly and kept; everything else is copied in from here.
"""

import json
import shutil
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "LINUX VERSION"
ZIP = OUT / "JJs Setlist - Linux.zip"

# Kept in LINUX VERSION itself.
LINUX_FILES = ["README.md", "run.sh", "install.sh", "build.sh",
               "build_appimage.sh", "build_appimage_2204.sh",
               "requirements.txt", ".gitignore"]
# Copied in from the main project:  source -> destination in LINUX VERSION.
COPIED = {
    "jjs_setlist.py": "jjs_setlist.py",
    "logo_data.py": "logo_data.py",
    "LICENSE": "LICENSE",
    "ARTWORK/JJ SETLIST Icon Source.png": "assets/jjs-setlist.png",
}
# Files that must be executable on Linux.
EXECUTABLE = {"run.sh", "install.sh", "build.sh",
              "build_appimage.sh", "build_appimage_2204.sh"}
TEXT = (".py", ".sh", ".txt", ".md", ".gitignore")


def read_config():
    try:
        return json.loads((HERE / "config.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def current_database(cfg):
    """The song spreadsheet the Windows app is using, if it's in this folder."""
    db = cfg.get("database")
    if not db:
        return None
    path = (HERE / db).resolve()
    try:
        return str(path.relative_to(HERE)).replace("\\", "/") if path.exists() else None
    except ValueError:
        return None


def copy(src, dest):
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.name.endswith(TEXT):
        # Linux line endings (LF) - a .sh with Windows endings won't run.
        text = src.read_text(encoding="utf-8").replace("\r\n", "\n")
        dest.write_text(text, encoding="utf-8", newline="\n")
    else:
        shutil.copy2(src, dest)


def main():
    missing = [f for f in LINUX_FILES if not (OUT / f).exists()]
    missing += [f for f in COPIED if not (HERE / f).exists()]
    if missing:
        print("Missing:", ", ".join(missing))
        return 1

    cfg = read_config()
    copied = dict(COPIED)
    db = current_database(cfg)
    if db:
        copied[db] = Path(db).name

    # Start the copied parts afresh, so renamed or deleted setlists and old
    # spreadsheets don't linger. The Linux-only files are left alone.
    shutil.rmtree(OUT / "setlists", ignore_errors=True)
    for old in list(OUT.glob("*.xls[xm]")) + [OUT / "database_link.txt", ZIP]:
        old.unlink(missing_ok=True)

    for name in LINUX_FILES:                    # make sure they have LF endings
        copy(OUT / name, OUT / name)
    for src, dest in copied.items():
        copy(HERE / src, OUT / dest)
        print("  copied ", src)
    for src in sorted((HERE / "setlists").glob("*.json")):
        copy(src, OUT / "setlists" / src.name)
        print("  copied ", f"setlists/{src.name}")

    # The Google Drive link of the song database, if set (the app picks it
    # up as its backup / Drive copy on first run).
    extra = []
    if cfg.get("backup_url"):
        (OUT / "database_link.txt").write_text(cfg["backup_url"] + "\n",
                                               encoding="utf-8", newline="\n")
        extra.append("database_link.txt")
        print("  added   database_link.txt (Google Drive copy of the song database)")

    files = (LINUX_FILES + list(copied.values()) + extra
             + [f"setlists/{p.name}" for p in sorted((OUT / "setlists").glob("*.json"))])

    # The zip records Unix permissions, so the .sh files stay runnable.
    with zipfile.ZipFile(ZIP, "w", zipfile.ZIP_DEFLATED) as z:
        for name in files:
            info = zipfile.ZipInfo.from_file(OUT / name, f"JJs Setlist/{name}")
            mode = 0o755 if Path(name).name in EXECUTABLE else 0o644
            info.external_attr = (0o100000 | mode) << 16
            info.create_system = 3                      # 3 = Unix
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, (OUT / name).read_bytes())
    print(f"\n  Zip for sending:  {ZIP}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
