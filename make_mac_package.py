"""
Copy the latest files into "APPLE MAC VERSION" and zip them up for sending
to a Mac user.  Run:  python make_mac_package.py   (or Make Mac Package.bat)
"""

import json
import shutil
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "APPLE MAC VERSION"
ZIP = OUT / "JJs Setlist - Mac.zip"

FILES = [
    "jjs_setlist.py",
    "logo_data.py",
    "make_icons.py",
    "Run JJs Setlist.command",
    "Build Mac App.command",
    "READ ME FIRST - Mac.txt",
    "LICENSE",
    "ARTWORK/JJ SETLIST Icon Source.png",
    "ARTWORK/JJ SETLIST Logo Source.png",
    "ARTWORK/TIP LIght Bulb #1.png",        # make_icons.py bakes it into logo_data.py
]
# Files that must be executable on the Mac (double-clickable).
EXECUTABLE = {"Build Mac App.command", "Run JJs Setlist.command"}


def current_database():
    """The song spreadsheet the Windows app is using, if it's in this folder."""
    try:
        db = json.loads((HERE / "config.json").read_text(encoding="utf-8"))["database"]
    except (OSError, ValueError, KeyError):
        return None
    path = (HERE / db).resolve()
    try:
        return str(path.relative_to(HERE)).replace("\\", "/") if path.exists() else None
    except ValueError:
        return None


def main():
    files = list(FILES)
    db = current_database()
    if db:
        files.append(db)
    # The saved setlists (the Mac app imports new ones automatically).
    files += sorted(f"setlists/{f.name}" for f in (HERE / "setlists").glob("*.json"))
    missing = [f for f in files if not (HERE / f).exists()]
    if missing:
        print("Missing:", ", ".join(missing))
        return 1

    OUT.mkdir(exist_ok=True)
    # The folder is output only: start afresh, so renamed or deleted files
    # (and old setlists) don't linger.
    for item in OUT.iterdir():
        if item.is_dir():
            shutil.rmtree(item, ignore_errors=True)
        else:
            item.unlink()
    for name in files:
        dest = OUT / name
        dest.parent.mkdir(parents=True, exist_ok=True)
        if name.endswith((".command", ".py", ".txt")):
            # Mac line endings (LF) - a .command with Windows endings won't run.
            text = (HERE / name).read_text(encoding="utf-8").replace("\r\n", "\n")
            dest.write_text(text, encoding="utf-8", newline="\n")
        else:
            shutil.copy2(HERE / name, dest)
        print("  copied ", name)

    # The Google Drive link of the song database, if set (the Mac app picks
    # it up as its backup / Drive copy).
    try:
        link = json.loads((HERE / "config.json").read_text(encoding="utf-8")).get("backup_url")
    except (OSError, ValueError):
        link = None
    if link:
        (OUT / "database_link.txt").write_text(link + "\n", encoding="utf-8", newline="\n")
        files.append("database_link.txt")
        print("  added   database_link.txt (Google Drive copy of the song database)")

    # The zip records Unix permissions, so the .command stays double-clickable.
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
