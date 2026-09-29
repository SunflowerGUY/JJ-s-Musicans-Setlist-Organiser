"""
JJ's Setlist
------------
Build up to 4 sets of up to 16 songs each from a song database
(CSV, or Excel .xlsx if openpyxl is installed), reorder them by
drag-and-drop or buttons, and save / reload named setlists.

Run:  python jjs_setlist.py
"""

import csv
import datetime as dt
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from urllib.parse import unquote

from logo_data import LOGO_PNG

IS_WIN = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"

# Version with the system's letter: 1.1.W (Windows), 1.1.M (Mac), 1.1.L (Linux).
APP_VERSION = "1.1." + ("W" if IS_WIN else "M" if IS_MAC else "L")
APP_NAME = "JJ's Musicians Setlist Organiser"  # shown in windows, help and printouts
APP_FILE_NAME = "JJs Setlist"           # for files and folders (no apostrophe)
APP_TITLE = f"{APP_NAME} - v{APP_VERSION}"
APP_CREDITS = ["Brought to you by JELLY JAZZ.", "Vibe Coding by Adrian Newington."]
APP_WEBSITE = "https://github.com/SunflowerGUY"
DONATE_URL = "https://paypal.me/jellyjazzsoftware"   # Help > Support, and in About

# Where settings and saved setlists live:
#  - Windows .exe: next to the .exe (PyInstaller unpacks the code itself into
#    a temporary folder that is deleted on exit), so the folder is portable.
#  - Mac and Linux (the .app, or jjs_setlist.py run directly): Documents/JJs Setlist.
#    An app bundle can't store files inside itself, and a script folder gets
#    replaced with each update - this way updates never touch saved setlists,
#    and the script and the .app share them.
#  - Running jjs_setlist.py directly on Windows: next to the script.
if not IS_WIN:
    APP_DIR = Path.home() / "Documents" / APP_FILE_NAME
    # The app used to be called "Musicians Setlist": move existing settings
    # and setlists across the first time the new name is used.
    _OLD_DIR = Path.home() / "Documents" / "Musicians Setlist"
    if _OLD_DIR.is_dir() and not APP_DIR.exists():
        try:
            _OLD_DIR.rename(APP_DIR)
        except OSError:
            pass
    APP_DIR.mkdir(parents=True, exist_ok=True)
elif getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent
SETLIST_DIR = APP_DIR / "setlists"
# The folder the app was delivered in (the unzipped folder on a Mac): the
# script's folder, or the folder containing "JJs Setlist.app".
if getattr(sys, "frozen", False) and IS_MAC:
    PACKAGE_DIR = Path(sys.executable).resolve().parents[3]
elif getattr(sys, "frozen", False):
    PACKAGE_DIR = Path(sys.executable).resolve().parent
else:
    PACKAGE_DIR = Path(__file__).resolve().parent
ICON_FILE = "setlist.ico"
CONFIG_FILE = APP_DIR / "config.json"
DEFAULT_DB = APP_DIR / "songs.csv"

DEFAULT_FONT_SIZE = 12
MIN_FONT_SIZE, MAX_FONT_SIZE = 9, 24

NUM_SETS = 4
MAX_SONGS = 16

# Spreadsheet header names recognised for each field (case-insensitive).
COLUMN_ALIASES = {
    "title": ["title", "song", "song title", "song name", "name", "track"],
    "artist": ["artist", "band", "performer", "original artist", "by"],
    "style": ["style", "genre", "type"],
    "vocalist": ["vocalist", "vocals", "vocal", "singer", "lead vocal"],
    # Optional column holding a songsheet path/URL (for CSV files, which
    # can't carry Excel's embedded hyperlinks).
    "link": ["link", "url", "pdf", "songsheet", "song sheet", "sheet", "chart"],
}

HYPERLINK_RE = re.compile(r'^=\s*HYPERLINK\(\s*"([^"]*)"', re.IGNORECASE)
SHEET_MARK = "📄 "

USED_FG = "#9e9e9e"            # grey: song already in the setlist

# Fonts that exist on each system.
UI_FONT = "Segoe UI" if IS_WIN else "Helvetica Neue" if IS_MAC else "DejaVu Sans"
MONO_FONT = "Consolas" if IS_WIN else "Menlo" if IS_MAC else "DejaVu Sans Mono"

# Highlighting of the current set.
ACTIVE_BORDER = "#0a64c8"      # blue: help headings, links
ACTIVE_FRAME = "#67c6c5"       # frame around the current set
MISSING_FG = "#d35400"         # orange: a set song no longer in the song database
MISSING_MARK = "⚠ "
# Same as the window background, so inactive sets show no border
# (the real colour is looked up from the theme in _build_ui).
INACTIVE_BORDER = "SystemButtonFace" if IS_WIN else "#ececec"
ACTIVE_LIST_BG = "#eef5ff"     # very light blue tint behind the songs
# Song library colours: navy background, white text, grey for songs
# already in the setlist, and a brighter blue for the selected song.
LIBRARY_BG = "#003d4b"
LIBRARY_FG = "#ffffff"
LIBRARY_SELECT_BG = "#3d8bfd"
LIBRARY_FRAME = "#afafaf"      # thin frame around the songlist
# Defaults for Font Size > Song Library Colours (choices saved in config.json).
LIBRARY_COLOUR_DEFAULTS = {"bg": LIBRARY_BG, "fg": LIBRARY_FG, "used": USED_FG,
                           "select": LIBRARY_SELECT_BG, "lib_frame": LIBRARY_FRAME,
                           "frame": ACTIVE_FRAME, "active_bg": ACTIVE_LIST_BG}
LIBRARY_COLOUR_NAMES = [("bg", "Songlist Background"), ("fg", "Song Title"),
                        ("used", "Songs assigned to a Setlist"),
                        ("select", "Selected Song"),
                        ("lib_frame", "Songlist Frame"),
                        ("frame", "Active Setlist Frame"),
                        ("active_bg", "Active Setlist Background")]
BORDER_WIDTH = 3

ORANGE = "#f57c00"             # "Create Template Spreadsheet" button
ORANGE_DARK = "#d65f00"        # ...when hovered / pressed


# ---------------------------------------------------------------- helpers

def cell_to_text(value):
    """Turn a CSV/Excel cell into a clean string."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def song_label(song):
    text = (SHEET_MARK if song.get("link") else "") + song["title"]
    if song.get("artist"):
        text += f"  —  {song['artist']}"
    extras = [song[k] for k in ("style", "vocalist") if song.get(k)]
    if extras:
        text += "   [" + " · ".join(extras) + "]"
    return text


def song_id(song):
    return (song["title"].casefold(), song.get("artist", "").casefold())


def rows_to_songs(rows, links=None):
    """rows: list of lists, first row is the header.
    links: optional matching list of lists with each cell's hyperlink."""
    if not rows:
        return []
    header = [cell_to_text(h).casefold() for h in rows[0]]
    cols = {}
    for field, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in header:
                cols[field] = header.index(alias)
                break
    cols.setdefault("title", 0)  # no recognised title column: use the first

    songs = []
    for r, row in enumerate(rows[1:], 1):
        song = {}
        for field, idx in cols.items():
            song[field] = cell_to_text(row[idx]) if idx < len(row) else ""
        if not song["title"]:
            continue
        if links and not song.get("link"):
            # Prefer the link on the song-name cell, else any link in the row.
            row_links = links[r]
            title_link = row_links[cols["title"]] if cols["title"] < len(row_links) else ""
            song["link"] = title_link or next((l for l in row_links if l), "")
        songs.append(song)
    return songs


def read_xlsx(path):
    """Return (rows, links) from the first worksheet of an Excel file."""
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError(
            "Reading Excel files needs the 'openpyxl' package.\n\n"
            "Install it with:  pip install openpyxl\n"
            "or save the spreadsheet as CSV instead (songsheet links "
            "are lost in CSV).")
    # Hyperlinks aren't available in read-only mode. The second copy (with
    # formulas) catches links made with =HYPERLINK("...", "...").
    values = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    formulas = openpyxl.load_workbook(path).worksheets[0]
    rows, links = [], []
    for row, frow in zip(values.iter_rows(), formulas.iter_rows()):
        rows.append([c.value for c in row])
        row_links = []
        for c, fc in zip(row, frow):
            hl = getattr(c, "hyperlink", None)
            target = (hl.target or "") if hl is not None else ""
            if not target and isinstance(fc.value, str):
                m = HYPERLINK_RE.match(fc.value)
                target = m.group(1) if m else ""
            row_links.append(target)
        links.append(row_links)
    return rows, links


def resolve_link(link, base_dir):
    """Make relative file links absolute (relative to the spreadsheet)."""
    link = link.strip()
    if not link or re.match(r"^[a-zA-Z][\w+.-]+:", link):  # URL (http:, file:, ...)
        return link
    p = Path(link)
    if not p.is_absolute():
        p = base_dir / p
    if not p.exists():
        unquoted = Path(unquote(str(p)))  # Excel sometimes stores "My%20Song.pdf"
        if unquoted.exists():
            p = unquoted
    return str(p.resolve())


# ---------------------------------------------------------------- Google Drive copy
DRIVE_CACHE = "drive_copy.xlsx"   # last download, kept for use offline
DRIVE_TIMEOUT = 10                # seconds


def drive_file_id(url):
    """The file id in a Google Drive / Google Docs share link, or None."""
    m = (re.search(r"/d/([\w-]{20,})", url or "")
         or re.search(r"[?&]id=([\w-]{20,})", url or ""))
    return m.group(1) if m else None


def download_drive_xlsx(url, dest):
    """Download a spreadsheet shared on Google Drive ("Anyone with the link")
    as .xlsx. Works for uploaded Excel files and for Google Sheets."""
    import ssl
    import urllib.request
    fid = drive_file_id(url)
    if not fid:
        raise OSError("That doesn't look like a Google Drive link.")
    try:
        import certifi                          # up-to-date certificates
        context = ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        context = ssl.create_default_context()
    last = "no response"
    for address in (f"https://drive.google.com/uc?export=download&id={fid}",
                    f"https://docs.google.com/spreadsheets/d/{fid}/export?format=xlsx"):
        try:
            request = urllib.request.Request(address, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(request, timeout=DRIVE_TIMEOUT,
                                        context=context) as response:
                data = response.read()
        except Exception as exc:                 # offline, timeout, HTTP error...
            last = getattr(exc, "reason", None) or exc
            continue
        if data[:2] == b"PK":                    # a real .xlsx (zip) file
            tmp = Path(str(dest) + ".part")
            tmp.write_bytes(data)
            os.replace(tmp, dest)
            return
        last = "Google sent a web page instead of the spreadsheet - check it's " \
               "shared as \"Anyone with the link\""
    raise OSError(f"Couldn't download it ({last}).")


def open_link(link, print_it=False, printer=None):
    """Open a web link or file with its usual program (or print a file)."""
    link = str(link)
    if IS_WIN:
        if print_it:
            print_text_file_windows(link, printer=printer)   # landscape via .NET
        else:
            os.startfile(link)
        return
    if not print_it and re.match(r"^https?://", link):
        webbrowser.open(link)
        return
    cmd = (["lp", "-o", "landscape", link] if print_it
           else ["open" if IS_MAC else "xdg-open", link])
    if print_it and printer:
        cmd[1:1] = ["-d", printer]
    try:
        subprocess.run(cmd, check=True, capture_output=True)
    except (subprocess.CalledProcessError, FileNotFoundError) as exc:
        raise OSError(f"Could not run {cmd[0]}: {exc}") from exc


# PowerShell + .NET printing: landscape, monospaced so the columns line up,
# a set is moved to a new page rather than split (when it fits on one page),
# a header on every page after the first (the first line of the text, marked
# "continued", so a loose page 2 still says which gig it belongs to), and a
# footer with the title and page numbers. Settings arrive in
# environment variables so no quoting of file names is needed.
PRINT_SCRIPT = r'''
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
$lines = [System.IO.File]::ReadAllLines($env:SETLIST_PRINT_FILE, [System.Text.Encoding]::UTF8)
$title = $env:SETLIST_PRINT_TITLE
$doc = New-Object System.Drawing.Printing.PrintDocument
$doc.DocumentName = $title
if ($env:SETLIST_PRINTER) { $doc.PrinterSettings.PrinterName = $env:SETLIST_PRINTER }
if ($env:SETLIST_PRINT_TO_FILE) {
    $doc.PrinterSettings.PrintToFile = $true
    $doc.PrinterSettings.PrintFileName = $env:SETLIST_PRINT_TO_FILE
}
if (-not $doc.PrinterSettings.IsValid) { throw "Printer not found: $($doc.PrinterSettings.PrinterName)" }
$doc.PrintController = New-Object System.Drawing.Printing.StandardPrintController
$doc.DefaultPageSettings.Landscape = $true
# Margins in 1/100 inch: 0.35" (9 mm) - inside what printers can reach.
$margin = 35
$doc.DefaultPageSettings.Margins = New-Object System.Drawing.Printing.Margins($margin, $margin, $margin, $margin)
$font = New-Object System.Drawing.Font("Consolas", 10)
$small = New-Object System.Drawing.Font("Consolas", 8)
$bold = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
# Pages after the first start with a header: heading line, rule, blank line.
$headerLines = 3
$heading = ""
if ($lines.Count -gt 0) { $heading = $lines[0].Trim() + "   (continued)" }

# Length of the block starting at each "SET n" line (up to the next one).
$block = @{}
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match '^SET \d') {
        $j = $i + 1
        while ($j -lt $lines.Count -and $lines[$j] -notmatch '^SET \d') { $j++ }
        $block[$i] = $j - $i
    }
}

function Layout($g) {
    # Returns the first line of every page (used for "Page n of N" and printing).
    $lh = $font.GetHeight(100)
    $m = $doc.DefaultPageSettings.Bounds
    $top = $margin; $bottom = $m.Height - $margin - 16   # 16 = room for the footer
    $bodyTop = $top + $headerLines * $lh                 # later pages: below the header
    $starts = @(0); $y = $top; $pageTop = $top
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($y -gt $pageTop) {
            # A set that fits on a page starts a new page rather than split.
            $setFits = $block.ContainsKey($i) -and ($block[$i] * $lh) -le ($bottom - $bodyTop)
            $setSpills = $setFits -and ($y + $block[$i] * $lh) -gt $bottom
            if ($setSpills -or ($y + $lh) -gt $bottom) { $starts += $i; $y = $bodyTop; $pageTop = $bodyTop }
        }
        $y += $lh
    }
    return ,$starts
}

$script:page = 0
$script:starts = $null
$doc.add_PrintPage({
    param($sender, $e)
    $g = $e.Graphics
    if ($null -eq $script:starts) { $script:starts = Layout $g }
    $lh = $font.GetHeight(100)
    $first = $script:starts[$script:page]
    $last = $lines.Count
    if ($script:page + 1 -lt $script:starts.Count) { $last = $script:starts[$script:page + 1] }
    $y = $e.MarginBounds.Top
    if ($script:page -gt 0) {
        # Header, so this page can't be mixed up with another gig's.
        $g.DrawString($heading, $bold, [System.Drawing.Brushes]::Black, $e.MarginBounds.Left, $y)
        $pageNo = "Page $($script:page + 1) of $($script:starts.Count)"
        $size = $g.MeasureString($pageNo, $bold)
        $g.DrawString($pageNo, $bold, [System.Drawing.Brushes]::Black,
                      $e.MarginBounds.Right - $size.Width, $y)
        $ruleY = $y + $lh * 1.3
        $g.DrawLine([System.Drawing.Pens]::Black, $e.MarginBounds.Left, $ruleY,
                    $e.MarginBounds.Right, $ruleY)
        $y += $headerLines * $lh
    }
    for ($i = $first; $i -lt $last; $i++) {
        $g.DrawString($lines[$i], $font, [System.Drawing.Brushes]::Black, $e.MarginBounds.Left, $y)
        $y += $lh
    }
    $script:page++
    $footer = "$title   -   Page $($script:page) of $($script:starts.Count)"
    $size = $g.MeasureString($footer, $small)
    $g.DrawString($footer, $small, [System.Drawing.Brushes]::Gray,
                  $e.MarginBounds.Right - $size.Width, $e.MarginBounds.Bottom + 4)
    $e.HasMorePages = $script:page -lt $script:starts.Count
})
if ($env:SETLIST_PRINT_DRYRUN) {
    # Test mode: report the page layout without printing anything.
    $bmp = New-Object System.Drawing.Bitmap(10, 10)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.PageUnit = [System.Drawing.GraphicsUnit]::Display
    $starts = Layout $g
    "paper=$($doc.DefaultPageSettings.PaperSize.PaperName) lineheight=$([math]::Round($font.GetHeight(100), 1)) pages=$($starts.Count) starts=$($starts -join ',')"
    return
}
$doc.Print()
'''


# Printers that save a file instead of printing: printing to them from the
# background would open a hidden "Save as" window, so the app asks for the
# file name itself.
FILE_PRINTERS = {"Microsoft Print to PDF": ".pdf",
                 "Microsoft XPS Document Writer": ".xps"}
PRINT_TIMEOUT = 90          # seconds before a silent printer is given up on


def default_printer():
    """Name of the default printer (Windows: instant, via the print system)."""
    if IS_WIN:
        try:
            import ctypes
            from ctypes import wintypes
            winspool = ctypes.WinDLL("winspool.drv")
            size = wintypes.DWORD(0)
            winspool.GetDefaultPrinterW(None, ctypes.byref(size))
            buf = ctypes.create_unicode_buffer(size.value or 1)
            if winspool.GetDefaultPrinterW(buf, ctypes.byref(size)):
                return buf.value
        except (OSError, AttributeError):
            pass
        return None
    try:
        out = subprocess.run([_find_tool("lpstat"), "-d"], capture_output=True, text=True,
                             timeout=10).stdout
        return out.split(":", 1)[1].strip() if ":" in out else None
    except (OSError, subprocess.SubprocessError):
        return None


def installed_printers():
    """Names of the installed printers."""
    if IS_WIN:
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                 r"Software\Microsoft\Windows NT\CurrentVersion\Devices")
            names, i = [], 0
            while True:
                try:
                    names.append(winreg.EnumValue(key, i)[0])
                    i += 1
                except OSError:
                    break
            return sorted(names, key=str.casefold)
        except OSError:
            return []
    try:
        out = subprocess.run([_find_tool("lpstat"), "-e"], capture_output=True, text=True,
                             timeout=10).stdout
        return sorted(out.split(), key=str.casefold)
    except (OSError, subprocess.SubprocessError):
        return []


# ---------------------------------------------------------------- PDF output
# A small built-in PDF maker, so printed pages look the same everywhere:
# A4 landscape, Courier (built into every PDF reader and printer, so nothing
# needs embedding), the same margins and page rules as Windows printing.
PDF_PAGE_W, PDF_PAGE_H = 842, 595       # A4 landscape, in points (1/72 inch)
PDF_MARGIN = 25                          # 0.35 inch
PDF_FONT_SIZE = 10
PDF_LEADING = 11.7                       # line spacing
PDF_FOOTER = 14                          # room for the footer line
HEADER_LINES = 3                         # later pages: heading, rule, blank line


def paginate(lines, per_page, later_per_page=None):
    """Split lines into pages; a "SET n" block that fits on one page is moved
    to a new page rather than split (same rule as Windows printing).
    later_per_page: room on pages after the first (they carry a header)."""
    later = later_per_page or per_page
    block = {}
    for i, line in enumerate(lines):
        if re.match(r"^SET \d", line):
            j = i + 1
            while j < len(lines) and not re.match(r"^SET \d", lines[j]):
                j += 1
            block[i] = j - i
    pages, current = [], []
    for i, line in enumerate(lines):
        if current:
            room = later if pages else per_page
            set_fits = i in block and block[i] <= later
            spills = set_fits and len(current) + block[i] > room
            if spills or len(current) + 1 > room:
                pages.append(current)
                current = []
        current.append(line)
    pages.append(current)
    return pages


def _pdf_string(text):
    raw = text.encode("cp1252", "replace")      # the PDF's WinAnsi encoding
    return raw.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)")


def make_pdf(text, title, path):
    """Write text as an A4-landscape PDF with a title/page-number footer, and
    a header on every page after the first (so a loose page 2 still says
    which gig it belongs to)."""
    lines = text.splitlines()
    per_page = int((PDF_PAGE_H - 2 * PDF_MARGIN - PDF_FOOTER) // PDF_LEADING)
    pages = paginate(lines, per_page, per_page - HEADER_LINES)
    heading = (lines[0].strip() + "   (continued)") if lines else ""
    objects = []

    def add(body):
        objects.append(body)
        return len(objects)

    font = add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Courier "
               b"/Encoding /WinAnsiEncoding >>")
    bold = add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Courier-Bold "
               b"/Encoding /WinAnsiEncoding >>")
    pages_id = add(b"")                          # filled in below
    kids = []
    for n, page in enumerate(pages, 1):
        ops = []
        y = PDF_PAGE_H - PDF_MARGIN - PDF_FONT_SIZE
        if n > 1:
            # Header, so this page can't be mixed up with another gig's.
            page_no = f"Page {n} of {len(pages)}"
            x = PDF_PAGE_W - PDF_MARGIN - len(page_no) * PDF_FONT_SIZE * 0.6
            rule = y - PDF_LEADING * 0.45
            ops += [b"BT /F2 %d Tf" % PDF_FONT_SIZE,
                    b"1 0 0 1 %.2f %.2f Tm (%s) Tj" % (PDF_MARGIN, y, _pdf_string(heading)),
                    b"1 0 0 1 %.2f %.2f Tm (%s) Tj" % (x, y, _pdf_string(page_no)),
                    b"ET",
                    b"0.8 w %.2f %.2f m %.2f %.2f l S"
                    % (PDF_MARGIN, rule, PDF_PAGE_W - PDF_MARGIN, rule)]
            y -= HEADER_LINES * PDF_LEADING
        ops += [b"BT", b"/F1 %d Tf" % PDF_FONT_SIZE]
        for line in page:
            ops.append(b"1 0 0 1 %.2f %.2f Tm (%s) Tj" % (PDF_MARGIN, y, _pdf_string(line)))
            y -= PDF_LEADING
        footer = f"{title}   -   Page {n} of {len(pages)}"
        x = PDF_PAGE_W - PDF_MARGIN - len(footer) * 7 * 0.6    # Courier: 0.6 em
        ops += [b"/F1 7 Tf 0.45 g",
                b"1 0 0 1 %.2f %.2f Tm (%s) Tj" % (x, PDF_MARGIN, _pdf_string(footer)),
                b"ET"]
        content = b"\n".join(ops)
        stream = add(b"<< /Length %d >>\nstream\n%s\nendstream" % (len(content), content))
        kids.append(add(b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %d %d] "
                        b"/Resources << /Font << /F1 %d 0 R /F2 %d 0 R >> >> /Contents %d 0 R >>"
                        % (pages_id, PDF_PAGE_W, PDF_PAGE_H, font, bold, stream)))
    objects[pages_id - 1] = (b"<< /Type /Pages /Kids [%s] /Count %d >>"
                             % (b" ".join(b"%d 0 R" % k for k in kids), len(kids)))
    catalog = add(b"<< /Type /Catalog /Pages %d 0 R >>" % pages_id)
    info = add(b"<< /Title (%s) /Producer (%s) >>" % (_pdf_string(title), _pdf_string(APP_NAME)))

    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = []
    for number, body in enumerate(objects, 1):
        offsets.append(len(out))
        out += b"%d 0 obj\n%s\nendobj\n" % (number, body)
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objects) + 1)
    out += b"".join(b"%010d 00000 n \n" % o for o in offsets)
    out += (b"trailer\n<< /Size %d /Root %d 0 R /Info %d 0 R >>\nstartxref\n%d\n%%%%EOF\n"
            % (len(objects) + 1, catalog, info, xref))
    Path(path).write_bytes(bytes(out))
    return len(pages)


def _find_tool(name):
    """Full path of a Mac/Linux command (apps started from the Dock get a
    minimal search path)."""
    for folder in ("/usr/bin", "/usr/local/bin", "/opt/homebrew/bin"):
        if Path(folder, name).exists():
            return str(Path(folder, name))
    return name


def print_pdf_unix(pdf_path, printer=None, title="Setlist"):
    """Print a PDF through the Mac (CUPS) print system."""
    cmd = [_find_tool("lp"), "-t", title, "-o", "fit-to-page"]
    if printer:
        cmd += ["-d", printer]
    cmd.append(str(pdf_path))
    try:
        result = subprocess.run(cmd, capture_output=True, text=True,
                                timeout=PRINT_TIMEOUT, stdin=subprocess.DEVNULL)
    except subprocess.TimeoutExpired:
        raise OSError(f"The printer didn't respond within {PRINT_TIMEOUT} seconds. "
                      "Check it's switched on and connected.") from None
    except OSError as exc:
        raise OSError(f"Could not run the print command: {exc}") from exc
    if result.returncode != 0:
        raise OSError(f"Printing failed: {(result.stderr or result.stdout).strip()[:300]}")


def print_text_file_windows(path, printer=None, to_file=None, dry_run=False):
    """Print a text file in landscape through .NET (via PowerShell).
    printer / to_file are for testing (e.g. "Microsoft Print to PDF")."""
    import base64
    env = dict(os.environ, SETLIST_PRINT_FILE=str(path),
               SETLIST_PRINT_TITLE=Path(path).stem)
    if printer:
        env["SETLIST_PRINTER"] = printer
    if to_file:
        env["SETLIST_PRINT_TO_FILE"] = str(to_file)
    if dry_run:
        env["SETLIST_PRINT_DRYRUN"] = "1"
    encoded = base64.b64encode(PRINT_SCRIPT.encode("utf-16-le")).decode()
    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded],
            env=env, capture_output=True, text=True, timeout=PRINT_TIMEOUT,
            stdin=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired:
        raise OSError(f"The printer didn't respond within {PRINT_TIMEOUT} seconds. "
                      "Check it's switched on and connected.") from None
    except (OSError, subprocess.SubprocessError) as exc:
        raise OSError(f"Printing failed: {exc}") from exc
    if result.returncode != 0:
        raise OSError(f"Printing failed: {(result.stderr or result.stdout).strip()[:300]}")
    return result.stdout.strip()


# Simple steps for building your own song spreadsheet from the template
# (shown in the app after creating it, and on the template's "How to use" tab).
TEMPLATE_STEPS = [
    "Open the template in Excel.",
    "Replace the two example rows with your own songs - one song per row. "
    "Only SONG NAME is required; Artist, Style and Vocalist are optional.",
    "Optional - songsheet links: click a song name, press Ctrl+K, paste the web "
    "address of its PDF (e.g. a Google Drive link) and click OK.",
    "Save with Ctrl+S. Keep it as an Excel Workbook (.xlsx) - saving as CSV "
    "loses the links.",
    f"Back in {APP_NAME}, press F5 (Reload Song Database) and your songs appear.",
]

def location_problem():
    """A warning if settings and setlists can't be kept where the app is
    running (e.g. started from inside a zip file, or a read-only folder)."""
    where = str(APP_DIR).casefold()
    if IS_WIN and "\\temp\\" in where and (".zip" in where or "\\temp1_" in where):
        return ("It looks like the app was started from inside a zip file. Windows "
                "runs it from a temporary folder, so your setlists and settings "
                "would be lost when you close it.\n\n"
                "Please close the app, right-click the zip file ▸ Extract All…, "
                "and run it from the extracted folder.")
    try:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        probe = APP_DIR / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
    except OSError:
        return (f"The app can't save in this folder:\n{APP_DIR}\n\n"
                "Your setlists and settings won't be kept. Please copy the app to "
                "a folder you can save in (e.g. Documents) and run it from there.")
    return None


def for_platform(text):
    """Show Mac key names on a Mac: Ctrl -> Cmd, Alt -> Option, etc."""
    if not IS_MAC:
        return text
    for old, new in (("Alt+F4", "Cmd+Q"), ("Ctrl+", "Cmd+"), ("Ctrl ", "Cmd "),
                     ("Alt+Up", "Option+Up"), ("Alt+Down", "Option+Down"),
                     ("Right-click", "Right-click (or Control-click)"),
                     ("Notepad", "TextEdit")):
        text = text.replace(old, new)
    return text


def acc(keys):
    """Menu accelerator text in each system's style (Tk on a Mac draws
    'Command-O' as the usual ⌘O symbol)."""
    if not IS_MAC:
        return keys
    keys = keys.replace("Alt+F4", "Command-Q").replace("Ctrl+", "Command-")
    keys = keys.replace("Shift+", "Shift-").replace("Alt+", "Option-")
    return keys


def load_song_database(path):
    path = Path(path)
    links = None
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        rows, links = read_xlsx(path)
    else:
        for enc in ("utf-8-sig", "cp1252"):
            try:
                with open(path, newline="", encoding=enc) as f:
                    rows = list(csv.reader(f))
                break
            except UnicodeDecodeError:
                continue
    songs = rows_to_songs(rows, links)
    for song in songs:
        if song.get("link"):
            song["link"] = resolve_link(song["link"], path.parent)
    songs.sort(key=lambda s: s["title"].casefold())
    return songs


def text_heading(title, subtitle):
    line = f"{title}  —  {subtitle}"
    return [line, "=" * len(line), dt.date.today().strftime("%A %d %B %Y"), ""]


def song_table(songs, numbered=False, size_to=None):
    """Plain-text columns (for printing): Song Name, Artist, Style, Vocalist.
    size_to: songs to size the columns by (so several tables line up)."""
    # Maximum column widths - sized for landscape pages.
    cols = [("title", "SONG NAME", 46), ("artist", "ARTIST", 26),
            ("style", "STYLE", 18), ("vocalist", "VOCALIST", 16)]
    widths = [min(cap, max([len(head)] + [len(s.get(key, "")) for s in size_to or songs]))
              for key, head, cap in cols]

    def fit(text, w):
        return (text if len(text) <= w else text[:w - 1] + "…").ljust(w)

    indent = "      " if numbered else "  "
    lines = [indent + "  ".join(fit(h, w) for (_, h, _), w in zip(cols, widths)),
             indent + "  ".join("-" * w for w in widths)]
    for i, s in enumerate(songs, 1):
        prefix = f"  {i:>2}. " if numbered else "  "
        lines.append((prefix + "  ".join(fit(s.get(k, ""), w)
                                         for (k, _, _), w in zip(cols, widths))).rstrip())
    return lines


def safe_filename(name):
    return re.sub(r'[<>:"/\\|?*]+', "_", name).strip(" .")


# ---------------------------------------------------------------- app

# ---------------------------------------------------------------- colour picker
# Adapted from colour_picker.py: a colour wheel, brightness bar, RGB / HSV /
# CMYK / hex fields and named colours. Each change is passed straight to a
# callback, so the Song Library changes colour live while you pick.
_NAMED_SRC = """
aliceblue f0f8ff antiquewhite faebd7 aqua 00ffff aquamarine 7fffd4 azure f0ffff
beige f5f5dc bisque ffe4c4 black 000000 blanchedalmond ffebcd blue 0000ff
blueviolet 8a2be2 brown a52a2a burlywood deb887 cadetblue 5f9ea0 chartreuse 7fff00
chocolate d2691e coral ff7f50 cornflowerblue 6495ed cornsilk fff8dc crimson dc143c
darkblue 00008b darkcyan 008b8b darkgoldenrod b8860b darkgray a9a9a9
darkgreen 006400 darkkhaki bdb76b darkmagenta 8b008b darkolivegreen 556b2f
darkorange ff8c00 darkorchid 9932cc darkred 8b0000 darksalmon e9967a
darkseagreen 8fbc8f darkslateblue 483d8b darkslategray 2f4f4f darkturquoise 00ced1
darkviolet 9400d3 deeppink ff1493 deepskyblue 00bfff dimgray 696969
dodgerblue 1e90ff firebrick b22222 floralwhite fffaf0 forestgreen 228b22
fuchsia ff00ff gainsboro dcdcdc ghostwhite f8f8ff gold ffd700 goldenrod daa520
gray 808080 green 008000 greenyellow adff2f honeydew f0fff0 hotpink ff69b4
indianred cd5c5c indigo 4b0082 ivory fffff0 khaki f0e68c lavender e6e6fa
lavenderblush fff0f5 lawngreen 7cfc00 lemonchiffon fffacd lightblue add8e6
lightcoral f08080 lightcyan e0ffff lightgoldenrodyellow fafad2 lightgray d3d3d3
lightgreen 90ee90 lightpink ffb6c1 lightsalmon ffa07a lightseagreen 20b2aa
lightskyblue 87cefa lightslategray 778899 lightsteelblue b0c4de lightyellow ffffe0
lime 00ff00 limegreen 32cd32 linen faf0e6 maroon 800000 mediumaquamarine 66cdaa
mediumblue 0000cd mediumorchid ba55d3 mediumpurple 9370db mediumseagreen 3cb371
mediumslateblue 7b68ee mediumspringgreen 00fa9a mediumturquoise 48d1cc
mediumvioletred c71585 midnightblue 191970 mintcream f5fffa mistyrose ffe4e1
moccasin ffe4b5 navajowhite ffdead navy 000080 oldlace fdf5e6 olive 808000
olivedrab 6b8e23 orange ffa500 orangered ff4500 orchid da70d6 palegoldenrod eee8aa
palegreen 98fb98 paleturquoise afeeee palevioletred db7093 papayawhip ffefd5
peachpuff ffdab9 peru cd853f pink ffc0cb plum dda0dd powderblue b0e0e6
purple 800080 rebeccapurple 663399 red ff0000 rosybrown bc8f8f royalblue 4169e1
saddlebrown 8b4513 salmon fa8072 sandybrown f4a460 seagreen 2e8b57 seashell fff5ee
sienna a0522d silver c0c0c0 skyblue 87ceeb slateblue 6a5acd slategray 708090
snow fffafa springgreen 00ff7f steelblue 4682b4 tan d2b48c teal 008080
thistle d8bfd8 tomato ff6347 turquoise 40e0d0 violet ee82ee wheat f5deb3
white ffffff whitesmoke f5f5f5 yellow ffff00 yellowgreen 9acd32
""".split()
NAMED_COLOURS = {_NAMED_SRC[i]: "#" + _NAMED_SRC[i + 1] for i in range(0, len(_NAMED_SRC), 2)}
_HEX_TO_NAME = {}
for _name, _hex in NAMED_COLOURS.items():
    _HEX_TO_NAME.setdefault(_hex, _name)

_WHEEL = 220          # colour wheel diameter (px)
_BAR_W, _BAR_H = 22, 220
_MARK_W = 10          # space left of the brightness bar for its marker


def _to_hex(r, g, b):
    return f"#{r:02x}{g:02x}{b:02x}"


def _hsv_to_rgb255(h, s, v):
    import colorsys
    return tuple(int(round(c * 255)) for c in colorsys.hsv_to_rgb(h, s, v))


def _hex_to_rgb(colour):
    colour = colour.lstrip("#")
    return tuple(int(colour[i:i + 2], 16) for i in (0, 2, 4))


def is_hex_colour(value):
    return isinstance(value, str) and re.fullmatch(r"#[0-9a-fA-F]{6}", value) is not None


def contrast_ratio(fg, bg):
    """WCAG contrast ratio between two #rrggbb colours."""
    def lum(colour):
        def lin(c):
            c /= 255
            return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
        r, g, b = _hex_to_rgb(colour)
        return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)
    a, b = sorted((lum(fg), lum(bg)), reverse=True)
    return (a + 0.05) / (b + 0.05)


class ColourPicker(ttk.Frame):
    """The colour wheel / brightness bar / fields from colour_picker.py.
    Calls on_change("#rrggbb") whenever the colour changes."""

    def __init__(self, parent, on_change):
        super().__init__(parent)
        self.on_change = on_change
        self.h = self.s = self.v = 0.0
        self._updating = False
        self._build()

    def _build(self):
        bg = _to_hex(*(c // 256 for c in self.winfo_rgb(self.winfo_toplevel().cget("bg"))))
        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="w")
        self.bar = tk.Canvas(top, width=_MARK_W + _BAR_W, height=_BAR_H, bg=bg,
                             highlightthickness=0, cursor="sb_v_double_arrow")
        self.bar.grid(row=0, column=0, padx=(0, 10))
        self.bar_img = tk.PhotoImage(width=_BAR_W, height=_BAR_H)
        self.bar.create_image(_MARK_W, 0, image=self.bar_img, anchor="nw")
        self.bar.create_polygon(0, 0, 0, 0, 0, 0, fill="black", tags="mark")
        for ev in ("<Button-1>", "<B1-Motion>"):
            self.bar.bind(ev, self._on_bar)
        self.bar.bind("<MouseWheel>", self._on_bar_wheel)

        self.wheel = tk.Canvas(top, width=_WHEEL, height=_WHEEL, bg=bg,
                               highlightthickness=0, cursor="crosshair")
        self.wheel.grid(row=0, column=1)
        self.wheel_img = self._make_wheel(bg)
        self.wheel.create_image(0, 0, image=self.wheel_img, anchor="nw")
        self.wheel.create_line(0, 0, 0, 0, width=2, tags="cross")
        self.wheel.create_line(0, 0, 0, 0, width=2, tags="cross")
        for ev in ("<Button-1>", "<B1-Motion>"):
            self.wheel.bind(ev, self._on_wheel)
        self.wheel.bind("<MouseWheel>", self._on_bar_wheel)

        mid = ttk.Frame(self, padding=(0, 12, 0, 0))
        mid.grid(row=1, column=0, sticky="w")
        self.swatch = tk.Frame(mid, width=64, height=112, relief="sunken", bd=1)
        self.swatch.grid(row=0, column=0, rowspan=4, padx=(0, 12))
        self.fields = {}
        layout = [("R", "H", "C"), ("G", "S", "M"), ("B", "V", "Y"), (None, None, "K")]
        for r, row in enumerate(layout):
            for c, name in enumerate(row):
                if name:
                    self._add_field(mid, name, r, 1 + c * 2)
        ttk.Label(mid, text="#").grid(row=3, column=1, sticky="e", padx=(0, 4))
        self.hex_var = tk.StringVar()
        hex_entry = ttk.Entry(mid, textvariable=self.hex_var, width=10)
        hex_entry.grid(row=3, column=2, columnspan=3, sticky="w", pady=2)
        hex_entry.bind("<KeyRelease>", self._on_hex)

        bottom = ttk.Frame(self, padding=(0, 10, 0, 0))
        bottom.grid(row=2, column=0, sticky="w")
        ttk.Label(bottom, text="Named:").grid(row=0, column=0, padx=(0, 4))
        self.name_var = tk.StringVar()
        names = ttk.Combobox(bottom, textvariable=self.name_var, width=22,
                             values=sorted(NAMED_COLOURS))
        names.grid(row=0, column=1)
        names.bind("<<ComboboxSelected>>", self._on_named)
        names.bind("<Return>", self._on_named)

    def _add_field(self, parent, name, row, col):
        maximum = 359 if name == "H" else 255
        var = tk.StringVar()
        ttk.Label(parent, text=f"{name}:").grid(row=row, column=col, sticky="e",
                                                padx=(8 if col > 1 else 0, 4))
        spin = tk.Spinbox(parent, from_=0, to=maximum, width=5, textvariable=var,
                          wrap=(name == "H"), command=lambda n=name: self._on_field(n))
        spin.grid(row=row, column=col + 1, sticky="w", pady=2)
        spin.bind("<KeyRelease>", lambda e, n=name: self._on_field(n))
        spin.bind("<MouseWheel>", lambda e, sp=spin: sp.invoke(
            "buttonup" if e.delta > 0 else "buttondown"))
        self.fields[name] = (var, maximum)

    def _make_wheel(self, bg):
        img = tk.PhotoImage(width=_WHEEL, height=_WHEEL)
        c = _WHEEL / 2
        radius = c - 1
        rows = []
        for y in range(_WHEEL):
            dy = c - y - 0.5
            row = []
            for x in range(_WHEEL):
                dx = x + 0.5 - c
                d = math.hypot(dx, dy)
                if d > radius:
                    row.append(bg)
                else:
                    h = (math.atan2(dy, dx) / (2 * math.pi)) % 1
                    row.append(_to_hex(*_hsv_to_rgb255(h, d / radius, 1)))
            rows.append("{" + " ".join(row) + "}")
        img.put(" ".join(rows))
        return img

    # ------------------------------------------------------------ colour
    def set_colour(self, colour, notify=False):
        """Show a colour (e.g. when switching which colour is being edited)."""
        r, g, b = _hex_to_rgb(colour)
        self.set_rgb(r, g, b, source=None, notify=notify)

    def set_rgb(self, r, g, b, source=None, notify=True):
        import colorsys
        h, s, v = colorsys.rgb_to_hsv(r / 255, g / 255, b / 255)
        if s == 0:          # grey: hue is undefined, keep the current one
            h = self.h
        if v == 0:          # black: keep saturation so the wheel marker stays put
            s = self.s
        self.set_hsv(h, s, v, source, notify)

    def set_hsv(self, h, s, v, source=None, notify=True):
        hs_changed = (h, s) != (self.h, self.s) or source is None
        self.h, self.s, self.v = h, s, v
        r, g, b = _hsv_to_rgb255(h, s, v)
        colour = _to_hex(r, g, b)
        self.swatch.config(bg=colour)

        c = _WHEEL / 2
        x = c + s * (c - 1) * math.cos(2 * math.pi * h)
        y = c - s * (c - 1) * math.sin(2 * math.pi * h)
        line_col = "white" if v < 0.5 or s > 0.6 else "black"
        h_line, v_line = self.wheel.find_withtag("cross")
        self.wheel.coords(h_line, x - 9, y, x + 10, y)
        self.wheel.coords(v_line, x, y - 9, x, y + 10)
        self.wheel.itemconfig("cross", fill="black" if s < 0.35 else line_col)
        if hs_changed:
            rows = []
            for yy in range(_BAR_H):
                col = _to_hex(*_hsv_to_rgb255(h, s, 1 - yy / (_BAR_H - 1)))
                rows.append("{" + " ".join([col] * _BAR_W) + "}")
            self.bar_img.put(" ".join(rows))
        my = (1 - v) * (_BAR_H - 1)
        self.bar.coords("mark", 0, my - 6, _MARK_W - 2, my, 0, my + 6)

        k = 1 - max(r, g, b) / 255
        if k >= 1:
            cm = mm = ym = 0
        else:
            cm = (1 - r / 255 - k) / (1 - k)
            mm = (1 - g / 255 - k) / (1 - k)
            ym = (1 - b / 255 - k) / (1 - k)
        values = {"R": r, "G": g, "B": b,
                  "H": int(round(h * 360)) % 360, "S": round(s * 255), "V": round(v * 255),
                  "C": round(cm * 255), "M": round(mm * 255), "Y": round(ym * 255),
                  "K": round(k * 255)}
        self._updating = True
        try:
            for name, val in values.items():
                if name != source:
                    self.fields[name][0].set(str(val))
            if source != "hex":
                self.hex_var.set(colour[1:])
            if source != "named":
                self.name_var.set(_HEX_TO_NAME.get(colour, ""))
        finally:
            self._updating = False
        if notify:
            self.on_change(colour)

    # ------------------------------------------------------------ events
    def _on_wheel(self, e):
        c = _WHEEL / 2
        dx, dy = e.x - c, c - e.y
        s = min(math.hypot(dx, dy) / (c - 1), 1.0)
        h = (math.atan2(dy, dx) / (2 * math.pi)) % 1
        self.set_hsv(h, s, self.v, source="wheel")

    def _on_bar(self, e):
        self.set_hsv(self.h, self.s, min(max(1 - e.y / (_BAR_H - 1), 0.0), 1.0), source="bar")

    def _on_bar_wheel(self, e):
        step = 3 / 255 * (1 if e.delta > 0 else -1)
        self.set_hsv(self.h, self.s, min(max(self.v + step, 0.0), 1.0), source="bar")

    def _on_field(self, name):
        if self._updating:
            return
        var, maximum = self.fields[name]
        try:
            val = int(var.get())
        except ValueError:
            return
        if not 0 <= val <= maximum:
            return
        try:
            vals = {n: int(v.get() or 0) for n, (v, _) in self.fields.items()}
        except ValueError:
            return
        vals[name] = val
        if name in "RGB":
            self.set_rgb(vals["R"], vals["G"], vals["B"], source=name)
        elif name in "HSV":
            self.set_hsv(vals["H"] / 360, vals["S"] / 255, vals["V"] / 255, source=name)
        else:
            k = vals["K"] / 255
            r, g, b = (round(255 * (1 - vals[n] / 255) * (1 - k)) for n in "CMY")
            self.set_rgb(r, g, b, source=name)

    def _on_hex(self, _e=None):
        text = self.hex_var.get().strip().lstrip("#")
        if len(text) == 3:
            text = "".join(ch * 2 for ch in text)
        if len(text) != 6:
            return
        try:
            r, g, b = (int(text[i:i + 2], 16) for i in (0, 2, 4))
        except ValueError:
            return
        self.set_rgb(r, g, b, source="hex")

    def _on_named(self, _e=None):
        hex_code = NAMED_COLOURS.get(self.name_var.get().strip().lower())
        if hex_code:
            self.set_rgb(*_hex_to_rgb(hex_code), source="named")


class SetlistApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self._set_window_icon()
        # Roomier default window (bigger text), but never larger than the screen.
        w = min(1500, self.winfo_screenwidth() - 60)
        h = min(920, self.winfo_screenheight() - 100)
        self.geometry(f"{w}x{h}")
        self.minsize(min(1000, w), min(620, h))
        self._restore_window()

        self.library = []          # all songs from the database
        self.filtered = []         # songs currently shown in the library box
        self.sets = [[] for _ in range(NUM_SETS)]
        self.active_set = 0
        self.dirty = False
        self.db_path = None
        self.db_name = None        # e.g. "JELLY JAZZ SETLIST - with links"
        self.db_desc = None        # where it came from, for Help > About
        self._drag = None

        SETLIST_DIR.mkdir(parents=True, exist_ok=True)
        imported = self._import_package_files()
        self._build_ui()
        self._build_menu()
        self._refresh_saved_list()
        self._refresh_all_sets()

        first_start = self._first_start_setup()
        self._load_database(startup=True)
        self._update_get_started()

        if imported:
            self._status(f"Added {len(imported)} setlist(s) from the update: "
                         + ", ".join(imported))
        if first_start:
            self._status(first_start)

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        # After the window is up: warn about an unsafe location, and welcome
        # new users who have no song database yet.
        self.after(300, self._startup_checks)

    # ---------- first start

    def _first_start_setup(self):
        """No song database chosen yet? Look next to the app for one:
          - database_link.txt  -> that Google Drive copy (used first)
          - a single spreadsheet (.xlsx) -> used as the main database
        Returns a status message if something was set up, else None."""
        cfg = self._read_config()
        if cfg.get("database") or cfg.get("backup_url"):
            return None
        updates, found = {}, []
        for folder in dict.fromkeys((PACKAGE_DIR, APP_DIR)):
            link_file = folder / "database_link.txt"
            if "backup_url" not in updates and link_file.exists():
                try:
                    link = link_file.read_text(encoding="utf-8").strip()
                except OSError:
                    link = ""
                if drive_file_id(link):
                    updates["backup_url"] = link
                    found.append("the Google Drive link")
            if "database" not in updates:
                sheets = [f for f in sorted(folder.glob("*.xls[xm]"))
                          if not f.name.startswith("~$") and f.name != DRIVE_CACHE]
                if len(sheets) == 1:
                    try:
                        if folder != APP_DIR:
                            shutil.copy2(sheets[0], APP_DIR / sheets[0].name)
                        updates["database"] = sheets[0].name
                        found.append(sheets[0].name)
                    except OSError:
                        pass
        if not updates:
            return None
        # With only a Drive link, it's the source; with both, the file is the
        # main one and Drive the backup.
        updates["drive_first"] = "database" not in updates
        self._write_config(**updates)
        return "First start: found " + " and ".join(found) + " - all set up."

    def _startup_checks(self):
        problem = location_problem()
        if problem:
            messagebox.showwarning(f"{APP_NAME} can't save here", problem)
        if not self.library:
            local, url, _ = self._db_settings()
            if not local and not url:
                self._show_welcome()

    def _show_welcome(self):
        """Help > Getting Started, and shown when there's no song database yet."""
        win = tk.Toplevel(self)
        win.title(f"Welcome to {APP_NAME}")
        win.transient(self)
        win.resizable(False, False)
        win.attributes("-alpha", 0.0)
        body = ttk.Frame(win, padding=20)
        body.pack(fill="both", expand=True)
        try:
            self._welcome_logo = tk.PhotoImage(data=LOGO_PNG)
            ttk.Label(body, image=self._welcome_logo).grid(row=0, column=0, rowspan=2,
                                                           sticky="n", padx=(0, 18))
        except tk.TclError:
            pass
        ttk.Label(body, text=f"Welcome to {APP_NAME}!", font=self.title_font).grid(
            row=0, column=1, sticky="w")
        ttk.Label(body, wraplength=420, justify="left", text=(
            "To get started, tell the app where your songs are. It's a one-time "
            "step - you can change it later in File ▸ Song Database Settings.")).grid(
            row=1, column=1, sticky="nw", pady=(6, 0))

        style = ttk.Style(self)
        style.configure("Welcome.TButton", font=self.heading_font, padding=(10, 8))
        choices = ttk.Frame(body)
        choices.grid(row=2, column=0, columnspan=2, sticky="we", pady=(18, 0))

        def then(action):
            def go():
                win.destroy()
                action()
            return go

        def use_template():
            self._create_template(parent=self, use_as_database=True)

        for text, note, action in (
                ("📂  Open my song spreadsheet…",
                 "An Excel file (.xlsx) or CSV with your songs.", self._choose_database),
                ("☁  Use a Google Drive link…",
                 "The share link of a spreadsheet on Google Drive.", self._database_settings),
                ("📝  Create a template spreadsheet…",
                 "Starting from scratch? Get a ready-to-fill Excel file, with simple "
                 "steps.", use_template),
                ("❓  Show me how",
                 "A step-by-step guide to setting up your song spreadsheet.",
                 self._show_spreadsheet_guide)):
            row = ttk.Frame(choices)
            row.pack(fill="x", pady=4)
            ttk.Button(row, text=text, style="Welcome.TButton", width=32,
                       command=then(action)).pack(side="left")
            ttk.Label(row, text=note, foreground="#666", wraplength=260,
                      justify="left").pack(side="left", padx=(12, 0))

        ttk.Label(body, foreground="#666", wraplength=560, justify="left",
                  text=f"Your setlists and settings are kept in:\n{APP_DIR}").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(18, 0))
        ttk.Button(body, text="Later", command=win.destroy).grid(
            row=4, column=0, columnspan=2, sticky="e", pady=(12, 0))
        win.bind("<Escape>", lambda e: win.destroy())
        self._center_over_main(win)
        win.attributes("-alpha", 1.0)
        win.focus_set()

    def _import_package_files(self):
        """Mac/Linux: the data folder (Documents/JJs Setlist) is separate from
        the unzipped folder, so bring in what came with the package:
          - setlists not seen before (never overwrites; one the user deleted
            isn't brought back, as imports are remembered in config.json)
          - the song spreadsheet, if no song database has been chosen yet.
        Returns the names of the setlists added."""
        if IS_WIN or PACKAGE_DIR == APP_DIR:
            return []
        cfg = self._read_config()
        seen = set(cfg.get("imported_setlists", []))
        added = []
        for src in sorted((PACKAGE_DIR / "setlists").glob("*.json")):
            if src.name in seen:
                continue
            seen.add(src.name)
            dest = SETLIST_DIR / src.name
            if not dest.exists():
                try:
                    shutil.copy2(src, dest)
                    added.append(src.stem)
                except OSError:
                    seen.discard(src.name)      # try again next time
        updates = {"imported_setlists": sorted(seen)}
        link_file = PACKAGE_DIR / "database_link.txt"
        if not cfg.get("backup_url") and link_file.exists():
            try:
                link = link_file.read_text(encoding="utf-8").strip()
                if drive_file_id(link):
                    updates["backup_url"] = link
            except OSError:
                pass
        if not cfg.get("database"):
            for src in sorted(PACKAGE_DIR.glob("*.xlsx")):
                if src.name.startswith("~$"):   # Excel's temporary lock files
                    continue
                dest = APP_DIR / src.name
                try:
                    if not dest.exists():
                        shutil.copy2(src, dest)
                    updates["database"] = src.name
                except OSError:
                    pass
                break
        self._write_config(**updates)
        return added

    # ---------- UI construction

    def _set_window_icon(self):
        """Windows: setlist.ico (bundled inside the .exe, or next to it).
        Mac/Linux: the baked-in logo (the .app's own icon covers the Dock)."""
        if not IS_WIN:
            try:
                self._icon_image = tk.PhotoImage(data=LOGO_PNG)
                self.iconphoto(True, self._icon_image)
            except tk.TclError:
                pass
            return
        for folder in (Path(getattr(sys, "_MEIPASS", APP_DIR)), APP_DIR):
            ico = folder / ICON_FILE
            if ico.exists():
                try:
                    self.iconbitmap(default=str(ico))
                except tk.TclError:
                    pass
                return

    def _build_ui(self):
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        # Fonts that follow Font Size > Larger / Smaller Text.
        self.list_font = tkfont.Font(family=UI_FONT, size=DEFAULT_FONT_SIZE)
        self.heading_font = tkfont.Font(family=UI_FONT, size=DEFAULT_FONT_SIZE,
                                        weight="bold")
        self.title_font = tkfont.Font(family=UI_FONT, size=DEFAULT_FONT_SIZE + 8,
                                      weight="bold")
        # Background of the current theme, for the invisible border of the
        # sets that aren't current.
        self.inactive_border = style.lookup("TFrame", "background") or INACTIVE_BORDER
        style.configure("Active.TLabelframe.Label", foreground="#0a64c8",
                        font=self.heading_font)
        style.configure("TLabelframe.Label", font=self.heading_font)
        size =self._read_config().get("font_size", DEFAULT_FONT_SIZE)
        self._apply_font_size(size if isinstance(size, int) else DEFAULT_FONT_SIZE,
                              save=False)

        # --- top bar: database + saved setlists
        top = ttk.Frame(self, padding=(8, 8, 8, 4))
        top.pack(fill="x")

        ttk.Button(top, text="Open Song Database…",
                   command=self._choose_database).pack(side="left")
        self.db_label = ttk.Label(top, text="No database loaded", foreground="#666")
        self.db_label.pack(side="left", padx=8)
        # While there's no song database, the label is a "get started" link.
        self.db_label.bind("<Button-1>",
                           lambda e: None if self.library else self._show_welcome())

        ttk.Button(top, text="Export…", command=self._export).pack(side="right")
        ttk.Button(top, text="New", command=self._new_setlist).pack(side="right", padx=4)
        ttk.Button(top, text="Delete", command=self._delete_saved).pack(side="right")
        ttk.Button(top, text="Load", command=self._load_selected).pack(side="right", padx=4)
        self.saved_var = tk.StringVar()
        self.saved_combo = ttk.Combobox(top, textvariable=self.saved_var,
                                        state="readonly", width=32)
        self.saved_combo.pack(side="right")
        self.saved_combo.bind("<<ComboboxSelected>>", lambda e: self._load_selected())
        ttk.Label(top, text="Saved setlists:").pack(side="right", padx=(0, 4))

        # --- name + save bar
        namebar = ttk.Frame(self, padding=(8, 0, 8, 4))
        namebar.pack(fill="x")
        ttk.Label(namebar, text="Setlist name:").pack(side="left")
        self.name_var = tk.StringVar()
        name_entry = ttk.Entry(namebar, textvariable=self.name_var, width=40)
        name_entry.pack(side="left", padx=4)
        name_entry.bind("<Return>", lambda e: self._save())
        ttk.Button(namebar, text="Save Setlist", command=self._save).pack(side="left")
        ttk.Label(namebar, text="Include Venue & Date in Setlist Name.",
                  foreground="#666").pack(side="left", padx=8)
        self.total_label = ttk.Label(namebar, text="", foreground="#444")
        self.total_label.pack(side="right")

        # status bar
        self.status = ttk.Label(self, text="Tip: drag songs from the library into a set, "
                                           "and drag within / between sets to reorder.",
                                relief="sunken", anchor="w", padding=(6, 2))
        self.status.pack(fill="x", side="bottom")

        # --- main area
        main = ttk.PanedWindow(self, orient="horizontal")
        main.pack(fill="both", expand=True, padx=8, pady=4)

        # library: built like a set (below), so the two borders line up. The
        # border is coloured in _colour_library_frame.
        lib_pane = ttk.Frame(main)
        main.add(lib_pane, weight=2)
        self.lib_border = tk.Frame(lib_pane, bg=self.inactive_border, padx=BORDER_WIDTH,
                                   pady=BORDER_WIDTH)
        self.lib_border.pack(fill="both", expand=True, padx=2, pady=2)
        lib = ttk.LabelFrame(self.lib_border, text="Song Library", padding=6,
                             style="Active.TLabelframe")    # blue title, like the current set
        lib.pack(fill="both", expand=True)
        search_row = ttk.Frame(lib)
        search_row.pack(fill="x")
        ttk.Label(search_row, text="Search:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self._refresh_library())
        self.search_entry = ttk.Entry(search_row, textvariable=self.search_var)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=4)
        self.search_entry.bind("<Escape>", lambda e: self.search_var.set(""))
        self.search_entry.bind("<Down>", lambda e: self._focus_box(self.lib_box))
        ttk.Button(search_row, text="✕", width=3,
                   command=lambda: self.search_var.set("")).pack(side="left")

        self.lib_box = self._make_listbox(lib)
        saved = self._read_config().get("library_colours") or {}
        self.lib_colours = {k: (saved[k] if is_hex_colour(saved.get(k)) else v)
                            for k, v in LIBRARY_COLOUR_DEFAULTS.items()}
        self._colour_lib_box(self.lib_box, self.lib_colours)
        self._colour_library_frame(self.lib_colours)
        self.lib_box.bind("<Double-Button-1>", lambda e: self._add_selected())
        self.lib_box.bind("<Return>", lambda e: self._add_selected())

        lib_btns = ttk.Frame(lib)
        lib_btns.pack(fill="x", pady=(4, 0))
        self.add_btn = ttk.Button(lib_btns, text="Add to Set 1  ▶",
                                  command=self._add_selected)
        self.add_btn.pack(side="left")
        ttk.Button(lib_btns, text=SHEET_MARK + "Open Songsheet",
                   command=lambda: self._open_sheet(self.lib_box)).pack(side="left", padx=4)
        self._bind_right_click(self.lib_box, self._library_menu)
        self.lib_count = ttk.Label(lib_btns, text="", foreground="#666")
        self.lib_count.pack(side="right")

        # sets, 2 x 2 grid
        sets_frame = ttk.Frame(main)
        main.add(sets_frame, weight=5)
        self.set_frames, self.set_boxes, self.set_borders = [], [], []
        for i in range(NUM_SETS):
            # A plain coloured frame around each set acts as its border; the
            # current set's border is coloured in _set_active().
            border = tk.Frame(sets_frame, bg=self.inactive_border, padx=BORDER_WIDTH,
                              pady=BORDER_WIDTH)
            border.grid(row=i // 2, column=i % 2, sticky="nsew", padx=2, pady=2)
            frame = ttk.LabelFrame(border, padding=6)
            frame.pack(fill="both", expand=True)
            for w in (border, frame):
                w.bind("<Button-1>", lambda e, n=i: self._focus_set(n))
            box = self._make_listbox(frame)
            box.bind("<FocusIn>", lambda e, n=i: self._set_active(n))
            box.bind("<Button-1>", lambda e, n=i: self._set_active(n), add="+")
            box.bind("<Delete>", lambda e, n=i: self._remove(n))
            if IS_MAC:  # the Mac "delete" key is BackSpace
                box.bind("<BackSpace>", lambda e, n=i: self._remove(n))
            alt = "Option" if IS_MAC else "Alt"
            box.bind(f"<{alt}-Up>", lambda e, n=i: self._move(n, -1))
            box.bind(f"<{alt}-Down>", lambda e, n=i: self._move(n, +1))
            box.bind("<Double-Button-1>", lambda e, b=box: self._open_sheet(b))
            self._bind_right_click(box, lambda e, n=i: self._set_menu(e, n))

            btns = ttk.Frame(frame)
            btns.pack(fill="x", pady=(4, 0))
            ttk.Button(btns, text="▲ Up", width=7,
                       command=lambda n=i: self._move(n, -1)).pack(side="left")
            ttk.Button(btns, text="▼ Down", width=8,
                       command=lambda n=i: self._move(n, +1)).pack(side="left", padx=4)
            ttk.Button(btns, text="Remove",
                       command=lambda n=i: self._remove(n)).pack(side="left")
            ttk.Button(btns, text="Clear Set",
                       command=lambda n=i: self._clear_set(n)).pack(side="right")

            self.set_frames.append(frame)
            self.set_boxes.append(box)
            self.set_borders.append(border)
        for c in range(2):
            sets_frame.columnconfigure(c, weight=1)
        for r in range(NUM_SETS // 2):
            sets_frame.rowconfigure(r, weight=1)

        # drag and drop on every listbox
        for box in [self.lib_box] + self.set_boxes:
            box.bind("<ButtonPress-1>", self._drag_start, add="+")
            box.bind("<B1-Motion>", self._drag_motion, add="+")
            box.bind("<ButtonRelease-1>", self._drag_release, add="+")
            # Stop Tk auto-scrolling the list sideways when the pointer
            # leaves it mid-drag, and keep lists pinned to the left edge.
            box.bind("<B1-Leave>", lambda e: "break")
            box.bind("<Shift-MouseWheel>", lambda e: "break")
            box.configure(xscrollcommand=lambda *a, b=box: b.xview_moveto(0)
                          if float(a[0]) > 0 else None)

        self._set_active(0)

    @staticmethod
    def _bind_right_click(widget, handler):
        # Right button is Button-3 on Windows; on a Mac it can be Button-2
        # (depending on the Tk version) or Control-click.
        widget.bind("<Button-3>", handler)
        if IS_MAC:
            widget.bind("<Button-2>", handler)
            widget.bind("<Control-Button-1>", handler)

    def _make_listbox(self, parent, height=4):
        # A small requested height: the list stretches to fill the space, and
        # the buttons packed below it are never squeezed out on a short window
        # or with large text (the list scrolls instead).
        wrap = ttk.Frame(parent)
        wrap.pack(fill="both", expand=True, pady=(4, 0))
        box = tk.Listbox(wrap, height=height, activestyle="none",
                         exportselection=False, font=self.list_font,
                         selectbackground="#0a64c8", selectforeground="white")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=box.yview)
        box.configure(yscrollcommand=sb.set)
        box.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        return box

    # ---------- menu bar

    def _build_menu(self):
        bar = tk.Menu(self)
        self.config(menu=bar)

        # The song list (library or a set) the user last worked in.
        self._last_box = self.lib_box
        for box in [self.lib_box] + self.set_boxes:
            box.bind("<FocusIn>", lambda e, b=box: setattr(self, "_last_box", b), add="+")

        m = tk.Menu(bar, tearoff=0)
        bar.add_cascade(label="File", underline=0, menu=m)
        m.add_command(label="Open Song Database…", underline=0, accelerator=acc("Ctrl+O"),
                      command=self._choose_database)
        m.add_command(label="Reload Song Database", underline=0, accelerator=acc("F5"),
                      command=self._reload_database)
        m.add_command(label="Song Database Settings…", underline=10,
                      command=self._database_settings)
        m.add_separator()
        m.add_command(label="New Setlist", underline=0, accelerator=acc("Ctrl+N"),
                      command=self._new_setlist)
        self.load_menu = tk.Menu(m, tearoff=0, postcommand=self._fill_load_menu)
        m.add_cascade(label="Load Setlist", underline=0, menu=self.load_menu)
        m.add_command(label="Save Setlist", underline=0, accelerator=acc("Ctrl+S"),
                      command=self._save)
        m.add_command(label="Save Setlist As…", underline=13, accelerator=acc("Ctrl+Shift+S"),
                      command=self._save_as)
        self.delete_menu = tk.Menu(m, tearoff=0, postcommand=self._fill_delete_menu)
        m.add_cascade(label="Delete Saved Setlist", underline=0, menu=self.delete_menu)
        m.add_separator()
        m.add_command(label="Export Setlist…", underline=1, accelerator=acc("Ctrl+E"),
                      command=self._export)
        m.add_command(label="Export Song Database as CSV (with web links)…",
                      underline=12, command=self._export_database_csv)
        m.add_command(label="Print Song List…", underline=6,
                      command=self._print_songlist)
        m.add_command(label="Print This Setlist…", underline=0, accelerator=acc("Ctrl+Shift+P"),
                      command=self._print_setlist)
        m.add_command(label="Save Setlist as PDF…", underline=14,
                      command=self._save_setlist_pdf)
        m.add_command(label="Save Song List as PDF…", underline=17,
                      command=self._save_songlist_pdf)
        self.printer_var = tk.StringVar()
        self.printer_menu = tk.Menu(m, tearoff=0, postcommand=self._fill_printer_menu)
        m.add_cascade(label="Printer", underline=2, menu=self.printer_menu)
        m.add_command(label="Open Setlists Folder", underline=13,
                      command=lambda: open_link(str(SETLIST_DIR)))
        m.add_separator()
        m.add_command(label="Exit", underline=1, accelerator=acc("Alt+F4"),
                      command=self._on_close)

        m = tk.Menu(bar, tearoff=0)
        bar.add_cascade(label="Edit", underline=0, menu=m)
        m.add_command(label="Add Song to Current Set", underline=0, accelerator=acc("Enter"),
                      command=self._add_selected)
        m.add_command(label="Remove Song from Set", underline=0, accelerator=acc("Delete"),
                      command=lambda: self._remove(self.active_set))
        m.add_command(label="Move Song Up", underline=10, accelerator=acc("Alt+Up"),
                      command=lambda: self._move(self.active_set, -1))
        m.add_command(label="Move Song Down", underline=10, accelerator=acc("Alt+Down"),
                      command=lambda: self._move(self.active_set, +1))
        m.add_separator()
        m.add_command(label="Clear Current Set…", underline=0,
                      command=lambda: self._clear_set(self.active_set))
        m.add_command(label="Clear All Sets…", underline=6, command=self._clear_all_sets)
        m.add_separator()
        m.add_command(label="Find Song", underline=0, accelerator=acc("Ctrl+F"),
                      command=self._focus_search)
        m.add_command(label="Clear Search", underline=1, accelerator=acc("Esc"),
                      command=lambda: self.search_var.set(""))

        m = tk.Menu(bar, tearoff=0)
        bar.add_cascade(label="Sets", underline=0, menu=m)
        self.active_var = tk.IntVar(value=self.active_set)
        for n in range(NUM_SETS):
            m.add_radiobutton(label=f"Set {n + 1}", underline=4, variable=self.active_var,
                              value=n, accelerator=acc(f"Ctrl+{n + 1}"),
                              command=lambda n=n: self._focus_set(n))
        m.add_separator()
        m.add_command(label="Song Library", underline=5, accelerator=acc("Ctrl+L"),
                      command=lambda: self._focus_box(self.lib_box))

        m = tk.Menu(bar, tearoff=0)
        bar.add_cascade(label="Songsheet", underline=2, menu=m)  # Alt+N (S is Sets)
        m.add_command(label="Open Songsheet for Selected Song", underline=0,
                      accelerator=acc("Ctrl+P"),
                      command=lambda: self._open_sheet(self._last_box))

        m = tk.Menu(bar, tearoff=0)
        bar.add_cascade(label="Font Size", underline=1, menu=m)  # Alt+O (F is File)
        m.add_command(label="Larger Text", underline=0, accelerator=acc("Ctrl++"),
                      command=lambda: self._apply_font_size(self.font_size + 1))
        m.add_command(label="Smaller Text", underline=0, accelerator=acc("Ctrl+-"),
                      command=lambda: self._apply_font_size(self.font_size - 1))
        m.add_command(label="Normal Text Size", underline=0, accelerator=acc("Ctrl+0"),
                      command=lambda: self._apply_font_size(DEFAULT_FONT_SIZE))
        m.add_separator()
        m.add_command(label="Songlist & Setlist Colours…", underline=0,
                      command=self._library_colours_dialog)

        m = tk.Menu(bar, tearoff=0)
        bar.add_cascade(label="Help", underline=0, menu=m)
        m.add_command(label="How to use - the basics", underline=0, accelerator=acc("F1"),
                      command=self._show_help)
        m.add_command(label="Getting Started…", underline=0,
                      command=self._show_welcome)
        m.add_command(label="Keyboard Shortcuts", underline=0,
                      command=self._show_shortcuts)
        m.add_separator()
        m.add_command(label="Setting Up Your Song Spreadsheet", underline=0,
                      command=self._show_spreadsheet_guide)
        m.add_command(label="Create Template Spreadsheet…", underline=7,
                      command=self._create_template)
        m.add_separator()
        m.add_command(label="Using a CSV File Instead", underline=6,
                      command=self._show_csv_guide)
        m.add_command(label="Create Template CSV…", underline=16,
                      command=self._create_csv_template)
        m.add_separator()
        m.add_command(label=f"Support {APP_NAME} ☕…", underline=0,
                      command=self._show_support)
        m.add_command(label="About", underline=0, command=self._show_about)

        keys = {
            "<Control-o>": self._choose_database,
            "<F5>": self._reload_database,
            "<Control-n>": self._new_setlist,
            "<Control-s>": self._save,
            "<Control-Shift-S>": self._save_as,
            "<Control-e>": self._export,
            "<Control-f>": self._focus_search,
            "<Control-l>": lambda: self._focus_box(self.lib_box),
            "<Control-p>": lambda: self._open_sheet(self._last_box),
            "<Control-Shift-P>": self._print_setlist,
            "<F1>": self._show_help,
            "<Control-plus>": lambda: self._apply_font_size(self.font_size + 1),
            "<Control-equal>": lambda: self._apply_font_size(self.font_size + 1),
            "<Control-KP_Add>": lambda: self._apply_font_size(self.font_size + 1),
            "<Control-minus>": lambda: self._apply_font_size(self.font_size - 1),
            "<Control-KP_Subtract>": lambda: self._apply_font_size(self.font_size - 1),
            "<Control-Key-0>": lambda: self._apply_font_size(DEFAULT_FONT_SIZE),
        }
        for n in range(NUM_SETS):
            keys[f"<Control-Key-{n + 1}>"] = lambda n=n: self._focus_set(n)
        for seq, fn in keys.items():
            if IS_MAC:  # Mac shortcuts use the Command (⌘) key
                seq = seq.replace("Control-", "Command-")
            self.bind_all(seq, lambda e, fn=fn: (fn(), "break")[1])
        if IS_MAC:
            # The Mac's own application menu: About, Help and Quit (Cmd+Q)
            # go through the app's usual handlers (e.g. "save first?").
            self.createcommand("tkAboutDialog", self._show_about)
            self.createcommand("tk::mac::ShowHelp", self._show_help)
            self.createcommand("tk::mac::Quit", self._on_close)

    def _apply_font_size(self, size, save=True):
        size = max(MIN_FONT_SIZE, min(MAX_FONT_SIZE, size))
        self.font_size = size
        # Tk's built-in named fonts are used by all the ttk labels, buttons,
        # entries and dropdowns.
        for name in ("TkDefaultFont", "TkTextFont", "TkHeadingFont", "TkMenuFont",
                     "TkCaptionFont", "TkSmallCaptionFont", "TkIconFont",
                     "TkTooltipFont"):
            try:
                tkfont.nametofont(name).configure(size=size)
            except tk.TclError:
                pass
        self.list_font.configure(size=size)
        self.heading_font.configure(size=size)
        self.title_font.configure(size=size + 8)
        if save:
            self._write_config(font_size=size)
            self._status(for_platform(
                f"Text size {size}  (Font Size menu, or Ctrl+ / Ctrl- / Ctrl+0)"))

    def _fill_load_menu(self):
        self.load_menu.delete(0, "end")
        names = [p.stem for p in self._saved_files()]
        for name in names:
            self.load_menu.add_command(
                label=name, command=lambda n=name: (self.saved_var.set(n),
                                                    self._load_selected()))
        if not names:
            self.load_menu.add_command(label="(no saved setlists)", state="disabled")

    def _fill_delete_menu(self):
        self.delete_menu.delete(0, "end")
        names = [p.stem for p in self._saved_files()]
        for name in names:
            self.delete_menu.add_command(label=name,
                                         command=lambda n=name: self._delete_saved(n))
        if not names:
            self.delete_menu.add_command(label="(no saved setlists)", state="disabled")

    def _save_as(self):
        name = simpledialog.askstring("Save Setlist As", "Setlist name:",
                                      initialvalue=self.name_var.get(), parent=self)
        if name and name.strip():
            self.name_var.set(name.strip())
            self.saved_var.set("")      # so an existing name asks before replacing
            self._save()

    def _reload_database(self):
        local, url, _ = self._db_settings()
        if local or url:
            self._load_database()
        else:
            self._choose_database()

    def _database_settings(self):
        """File > Song Database Settings: main file, Google Drive link, order."""
        local, url, drive_first = self._db_settings()
        win = tk.Toplevel(self)
        win.title("Song Database Settings")
        win.transient(self)
        win.resizable(False, False)
        body = ttk.Frame(win, padding=16)
        body.pack(fill="both", expand=True)

        path_var = tk.StringVar(value=local or "")
        url_var = tk.StringVar(value=url)
        first_var = tk.BooleanVar(value=drive_first)
        result = tk.StringVar()

        ttk.Label(body, text="Main song database - a spreadsheet on this computer:",
                  font=self.heading_font).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Entry(body, textvariable=path_var, width=70).grid(row=1, column=0, sticky="we",
                                                               pady=(4, 0))

        def browse():
            chosen = filedialog.askopenfilename(
                parent=win, title="Main song database", initialdir=APP_DIR,
                filetypes=[("Spreadsheets", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")])
            if chosen:
                path_var.set(chosen)
        ttk.Button(body, text="Browse…", command=browse).grid(row=1, column=1, padx=(6, 0),
                                                              pady=(4, 0))

        ttk.Label(body, text="Backup - Google Drive link to the same spreadsheet (optional):",
                  font=self.heading_font).grid(row=2, column=0, columnspan=2, sticky="w",
                                               pady=(16, 0))
        ttk.Entry(body, textvariable=url_var, width=70).grid(row=3, column=0, sticky="we",
                                                              pady=(4, 0))

        def test():
            link = url_var.get().strip()
            if not drive_file_id(link):
                result.set("✗  That doesn't look like a Google Drive link.")
                return
            win.config(cursor="watch")
            win.update_idletasks()
            tmp = Path(tempfile.gettempdir()) / "jjs_drive_test.xlsx"
            try:
                download_drive_xlsx(link, tmp)
                songs = load_song_database(tmp)
                links = sum(1 for x in songs if x.get("link"))
                result.set(f"✓  Downloaded OK: {len(songs)} songs, {links} songsheets.")
            except Exception as exc:
                result.set(f"✗  {exc}")
            finally:
                win.config(cursor="")
        ttk.Button(body, text="Test link", command=test).grid(row=3, column=1, padx=(6, 0),
                                                              pady=(4, 0))
        ttk.Label(body, textvariable=result, wraplength=560).grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))

        ttk.Label(body, text="Which to use first:", font=self.heading_font).grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(14, 0))
        ttk.Radiobutton(body, variable=first_var, value=False,
                        text="This computer's spreadsheet first  (Google Drive copy "
                             "as the backup)").grid(row=6, column=0, columnspan=2,
                                                    sticky="w", pady=(4, 0))
        ttk.Radiobutton(body, variable=first_var, value=True,
                        text="The Google Drive copy first - always the latest  "
                             "(this computer's spreadsheet as the backup)").grid(
            row=7, column=0, columnspan=2, sticky="w")

        ttk.Label(body, foreground="#666", wraplength=600, justify="left", text=(
            "Tips: share the Drive file as “Anyone with the link”. To update it, use "
            "Manage versions ▸ Upload new version in Google Drive, so the link stays "
            "the same. The last download is kept, so the Drive copy also works "
            "offline.")).grid(row=8, column=0, columnspan=2, sticky="w", pady=(14, 0))

        def save():
            link = url_var.get().strip()
            if link and not drive_file_id(link):
                messagebox.showerror("Song Database Settings",
                                     "The Google Drive link doesn't look right.", parent=win)
                return
            main = path_var.get().strip()
            if not main and not link:
                messagebox.showerror("Song Database Settings",
                                     "Please choose a spreadsheet or a Google Drive link.",
                                     parent=win)
                return
            self._write_config(database=self._stored_path(main) if main else None,
                               backup_url=link or None, drive_first=first_var.get())
            win.destroy()
            self._load_database()

        btns = ttk.Frame(body)
        btns.grid(row=9, column=0, columnspan=2, sticky="we", pady=(16, 0))
        ttk.Button(btns, text="Save", command=save).pack(side="left")
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="right")
        win.bind("<Escape>", lambda e: win.destroy())
        self._center_over_main(win)

    def _focus_search(self):
        self.search_entry.focus_set()
        self.search_entry.select_range(0, "end")

    def _focus_box(self, box):
        box.focus_set()
        if box.size() and not box.curselection():
            box.selection_set(0)
            box.activate(0)

    def _focus_set(self, n):
        self._set_active(n)
        self._focus_box(self.set_boxes[n])

    def _show_help(self):
        messagebox.showinfo("How to use - the basics", for_platform(
            "BUILDING SETS\n"
            "• Double-click a library song (or press Enter) to add it to the "
            "current set (highlighted in blue).\n"
            "• Or drag songs from the library into any set.\n"
            "• Drag within or between sets to reorder; drag back to the library "
            "to remove.\n"
            "• Ctrl+1 … Ctrl+4 jump to a set, Ctrl+L to the library, "
            "Ctrl+F to search.\n"
            "• In a set: Alt+Up / Alt+Down move a song, Delete removes it.\n"
            "• A song marked ⚠ in orange is no longer in the song database - "
            "right-click it to Remove it or Replace with… another song.\n\n"
            "SONGSHEETS\n"
            "• Songs marked 📄 have a songsheet. Double-click a song in a set, "
            "or press Ctrl+P, to open it.\n\n"
            "SAVING\n"
            "• Type a name and press Ctrl+S. Load saved setlists from the "
            "dropdown or File ▸ Load Setlist.\n\n"
            "SONG DATABASE\n"
            "• Use an Excel .xlsx file (CSV loses the songsheet links) with the "
            "columns SONG NAME, Artist, Style, Vocalist. Press F5 to reload it "
            "after editing.\n"
            "• See Help ▸ Setting Up Your Song Spreadsheet for full details."))

    # ---------- spreadsheet guide & template

    SPREADSHEET_GUIDE = [
        ("h1", "Setting Up Your Song Spreadsheet"),
        ("p", "The song library comes from an Excel spreadsheet (.xlsx). Each row is "
              "one song, and each song name can carry a link to its songsheet (PDF)."),
        ("tip", "Quickest start: click “Create Template Spreadsheet…” below, "
                "fill it in, then open it with File ▸ Open Song Database."),

        ("h2", "1.  The column headings (row 1)"),
        ("p", "Put these headings in the first row of the first worksheet. "
              "Capitals don't matter and the columns can be in any order."),
        ("code", "  SONG NAME      Artist      Style      Vocalist"),
        ("b", "SONG NAME — required. The title as you want it shown."),
        ("b", "Artist — optional. Original artist or band."),
        ("b", "Style — optional. e.g. Ballad, 60's  ·  Rock  ·  Latin"),
        ("b", "Vocalist — optional. Who sings it."),
        ("p", "Other accepted heading names:"),
        ("b", "SONG NAME:  Title, Song, Song Title, Name, Track"),
        ("b", "Artist:  Band, Performer, Original Artist"),
        ("b", "Style:  Genre, Type"),
        ("b", "Vocalist:  Singer, Vocals, Vocal, Lead Vocal"),
        ("p", "Any other columns (notes, keys, etc.) are simply ignored, so you can "
              "keep extra information in the sheet."),

        ("h2", "2.  One song per row"),
        ("b", "Start on row 2, one song per row. Empty rows are skipped."),
        ("b", "If you have two versions of a song, make the names different, e.g. "
              "“Scarborough Fair [v1]” and “Scarborough Fair [v2]”."),
        ("b", "Style and Vocalist are searchable in the app — type “Ballad” or a "
              "singer's name in the Search box."),

        ("h2", "3.  Adding songsheet links (embedded web links)"),
        ("p", "Attach the link to the SONG NAME cell. Songs with a link show 📄 in "
              "the app and open with a double-click (in a set) or Ctrl+P."),
        ("h3", "Method A — Insert Link (recommended)"),
        ("n", "1.  Click the song-name cell."),
        ("n", "2.  Press Ctrl+K  (or Insert ▸ Link)."),
        ("n", "3.  In “Address”, paste the web address of the songsheet — or choose "
              "“Existing File” and browse to a PDF on your computer."),
        ("n", "4.  Click OK. The name turns blue and underlined."),
        ("h3", "Method B — HYPERLINK formula"),
        ("code", '  =HYPERLINK("https://drive.google.com/file/d/…/view", "El Paso")'),
        ("p", "The first part is the link, the second is the song name shown."),
        ("h3", "Method C — a separate link column"),
        ("p", "Add a column headed Songsheet (or Link, URL, PDF) and type the web "
              "address or file path in it. This also works in CSV files."),

        ("h2", "4.  Getting a Google Drive link for a PDF"),
        ("n", "1.  In Google Drive, right-click the PDF ▸ Share ▸ Copy link."),
        ("n", "2.  If bandmates will use the app too, set General access to "
              "“Anyone with the link”."),
        ("n", "3.  Paste the link in Excel with Ctrl+K as in Method A."),
        ("p", "Songsheets on your computer work too. Tip: keep the PDFs in a folder "
              "beside the spreadsheet (e.g. a Songsheets folder next to it) so the "
              "links keep working if you copy the whole folder to another PC."),

        ("h2", "5.  Editing a linked cell"),
        ("b", "Clicking a linked cell opens the link. To edit the name instead, "
              "select the cell with the arrow keys (or click and hold) and press F2."),
        ("b", "Editing the text keeps the link. To change the link, press Ctrl+K."),

        ("h2", "6.  Saving — important!"),
        ("warn", "Always save as an Excel Workbook, because saving as CSV "
                 "throws away every songsheet link."),
        ("b", "Use Ctrl+S. If Excel asks about the format, choose Excel Workbook."),
        ("b", "Only the first worksheet is read. Formatting (fonts, colours, "
              "column widths) doesn't matter to the app."),

        ("h2", "7.  Loading it into the app"),
        ("b", "File ▸ Open Song Database… (Ctrl+O) and choose the .xlsx. "
              "The app remembers it."),
        ("b", "After editing the spreadsheet, press F5 (File ▸ Reload Song Database)."),
        ("b", "The top bar shows how many songs and songsheets were found — if the "
              "songsheet count is 0, the file was probably saved as CSV."),

        ("h2", "8.  A backup copy on Google Drive"),
        ("p", "Keep a copy of the spreadsheet on Google Drive and the app can use it "
              "as a backup — or as the main source, so every computer (and bandmate) "
              "always has the latest version."),
        ("n", "1.  Upload the .xlsx to Google Drive and share it as “Anyone with the "
              "link”. Leave it as an Excel file (don't “Save as Google Sheets”)."),
        ("n", "2.  Right-click it in Drive ▸ Share ▸ Copy link."),
        ("n", "3.  In the app: File ▸ Song Database Settings, paste the link and click "
              "Test link. Choose which to use first, then Save."),
        ("b", "If the first choice can't be loaded (file missing, no internet), the "
              "other one is used and the top bar shows ⚠ BACKUP in orange."),
        ("b", "The last download is kept on the computer, so the Drive copy also "
              "works offline (e.g. at a gig)."),
        ("warn", "To update the Drive copy, use Manage versions ▸ Upload new version in "
                 "Google Drive. Deleting it and uploading a new file gives it a new link."),
    ]

    CSV_GUIDE = [
        ("h1", "Using a CSV File Instead"),
        ("p", "The song library doesn't have to be an Excel workbook. A CSV file "
              "(comma-separated values) works too — handy if you keep your list in "
              "Google Sheets, Numbers, LibreOffice or a plain text editor."),
        ("warn", "A CSV can't hold embedded links. Excel drops every link hidden "
                 "behind a song name when it saves as CSV. Put the links in their "
                 "own column instead (see step 3)."),
        ("tip", "Quickest start: click “Create Template CSV…” below — it already "
                "has the right headings and an example link column."),
        ("tip", "Already have an Excel sheet with embedded links? Load it, then use "
                "File ▸ Export Song Database as CSV (with web links). Every link is "
                "written into a Songsheet column, so none are lost."),

        ("h2", "1.  The column headings (first line)"),
        ("p", "The same headings as the Excel version — any order, capitals don't "
              "matter, and extra columns are ignored:"),
        ("code", "  SONG NAME,Artist,Style,Vocalist,Songsheet"),
        ("b", "SONG NAME — required. Also accepted: Title, Song, Name, Track"),
        ("b", "Artist, Style, Vocalist — optional (same alternatives as Excel)."),
        ("b", "Songsheet — optional link column. Also accepted: Link, URL, PDF, "
              "Sheet, Chart"),

        ("h2", "2.  One song per line"),
        ("code", "  El Paso,Marty Robbins,Country,Adrian,https://drive.google.com/…\n"
                 "  Blue Moon,Nat King Cole,Ballad,Gary,Songsheets\\Blue Moon.pdf\n"
                 "  Moon River,Henry Mancini,Ballad,,"),
        ("b", "Leave a value empty by putting nothing between the commas "
              "(Moon River has no vocalist or songsheet above)."),
        ("b", "If a value contains a comma, wrap it in double quotes, e.g. "
              "\"Ballad, 60's\". Excel and Google Sheets do this for you."),
        ("b", "Extra commas at the end of lines are fine — they're ignored."),

        ("h2", "3.  Songsheet links in a CSV"),
        ("p", "Type (or paste) the full link into the Songsheet column:"),
        ("b", "A web address, e.g. a Google Drive link "
              "(Drive: right-click the PDF ▸ Share ▸ Copy link)."),
        ("b", "A file on your computer, e.g.  C:\\Music\\Songsheets\\El Paso.pdf"),
        ("b", "Or a path relative to the CSV's own folder, e.g.  "
              "Songsheets\\El Paso.pdf  — this keeps working if you copy the whole "
              "folder to a USB stick or another PC."),
        ("p", "Songs with a link show 📄 in the app, exactly as with Excel."),

        ("h2", "4.  Saving a CSV"),
        ("b", "Excel: File ▸ Save As ▸ “CSV UTF-8 (Comma delimited) (*.csv)”. "
              "UTF-8 keeps accented names (Hasta Mañana) correct."),
        ("b", "Google Sheets: File ▸ Download ▸ Comma-separated values (.csv)."),
        ("b", "Text editor: save with a .csv extension (UTF-8 if offered)."),

        ("h2", "5.  Loading it into the app"),
        ("b", "File ▸ Open Song Database… (Ctrl+O), choose the .csv. The app "
              "remembers it."),
        ("b", "After editing, press F5 to reload."),
        ("b", "Check the top bar: it shows how many songs and songsheets were found."),

        ("h2", "Excel or CSV?"),
        ("b", "Excel (.xlsx): links can hide behind the song name, and your "
              "formatting is kept."),
        ("b", "CSV: simple plain text that any program can edit, but links must go "
              "in their own column and there's no formatting."),
        ("p", "Both work fully in the app — choose whichever suits the way you keep "
              "your song list."),
    ]

    def _create_csv_template(self, parent=None):
        """Write a starter CSV with the right headings and example rows."""
        path = filedialog.asksaveasfilename(
            parent=parent or self, title="Create template CSV",
            initialdir=APP_DIR, initialfile="My Songs.csv",
            defaultextension=".csv", filetypes=[("CSV file", "*.csv")])
        if not path:
            return
        rows = [
            ["SONG NAME", "Artist", "Style", "Vocalist", "Songsheet"],
            ["Example Song (web link)", "Example Artist", "Ballad, 60's", "Adrian",
             "https://example.com/songsheets/example-song.pdf"],
            ["Another Song (file link)", "Another Artist", "Rock", "",
             "Songsheets\\Another Song.pdf"],
            ["A Third Song (no link yet)", "Third Artist", "", "", ""],
        ]
        try:
            # UTF-8 with BOM so Excel opens accented names correctly.
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                csv.writer(f).writerows(rows)
        except OSError as exc:
            messagebox.showerror("Create template", f"Could not save:\n{exc}",
                                 parent=parent)
            return
        self._status(f"Template created: {path}")
        if messagebox.askyesno(
                "Template created",
                f"Created:\n{path}\n\nIt has the right column headings, including a "
                "Songsheet column, and three example rows.\nReplace them with your "
                "songs, save, then use File ▸ Open Song Database.\n\n"
                "Open it now?", parent=parent):
            try:
                open_link(path)
            except OSError as exc:
                messagebox.showerror("Open template", str(exc), parent=parent)

    SHORTCUTS_GUIDE = [
        ("h1", "Keyboard Shortcuts"),

        ("h2", "Menus  (hold Alt to see the underlined letters)"),
        ("k", "Alt+F\tFile"),
        ("k", "Alt+E\tEdit"),
        ("k", "Alt+S\tSets"),
        ("k", "Alt+N\tSongsheet"),
        ("k", "Alt+O\tFont Size"),
        ("k", "Alt+H\tHelp"),

        ("h2", "File"),
        ("k", "Ctrl+O\tOpen Song Database"),
        ("k", "F5\tReload Song Database (after editing it)"),
        ("k", "Ctrl+N\tNew Setlist"),
        ("k", "Ctrl+S\tSave Setlist  (or Enter in the Setlist name box)"),
        ("k", "Ctrl+Shift+S\tSave Setlist As"),
        ("k", "Ctrl+E\tExport Setlist"),
        ("k", "Ctrl+Shift+P\tPrint This Setlist"),
        ("k", "Alt+F4\tExit"),

        ("h2", "Finding and adding songs"),
        ("k", "Ctrl+F\tFind Song (jump to the Search box)"),
        ("k", "Esc\tClear the search (in the Search box)"),
        ("k", "Down arrow\tFrom the Search box into the song library"),
        ("k", "Enter\tAdd the selected library song to the current set"),
        ("k", "Double-click\tAdd a library song to the current set"),

        ("h2", "Sets"),
        ("k", "Ctrl+1 … Ctrl+4\tGo to Set 1 … Set 4"),
        ("k", "Ctrl+L\tGo to the Song Library"),
        ("k", "Alt+Up\tMove the selected song up"),
        ("k", "Alt+Down\tMove the selected song down"),
        ("k", "Delete\tRemove the selected song from the set"),

        ("h2", "Songsheets"),
        ("k", "Ctrl+P\tOpen the selected song's songsheet"),
        ("k", "Double-click\tOpen the songsheet (on a song in a set)"),
        ("k", "Right-click\tSong menu: Open Songsheet, Move, Remove, Add to Set"),

        ("h2", "Font Size"),
        ("k", "Ctrl +\tLarger text  (Ctrl = also works)"),
        ("k", "Ctrl −\tSmaller text"),
        ("k", "Ctrl+0\tNormal text size"),

        ("h2", "Help"),
        ("k", "F1\tHow to use - the basics"),
        ("k", "Esc\tClose a help window"),

        ("h2", "Mouse"),
        ("b", "Drag a song from the library into any set, at any position."),
        ("b", "Drag within a set to reorder, or to another set to move it."),
        ("b", "Drag a song from a set back to the library to remove it."),
    ]

    def _show_shortcuts(self):
        content = self.SHORTCUTS_GUIDE
        if IS_MAC:
            # Macs have no Alt+letter menu keys: drop the "Menus" section.
            content, skip = [], False
            for tag, line in self.SHORTCUTS_GUIDE:
                if tag == "h2":
                    skip = line.startswith("Menus")
                if not skip:
                    content.append((tag, line))
        self._show_guide("Keyboard Shortcuts", content)

    def _show_spreadsheet_guide(self):
        self._show_guide("Setting Up Your Song Spreadsheet", self.SPREADSHEET_GUIDE,
                         "Create Template Spreadsheet…", self._create_template)

    def _show_csv_guide(self):
        self._show_guide("Using a CSV File Instead", self.CSV_GUIDE,
                         "Create Template CSV…", self._create_csv_template)

    def _show_guide(self, title, content, action_text=None, action=None):
        """Scrollable help window with an orange action button."""
        win = tk.Toplevel(self)
        win.title(title)
        win.transient(self)
        w = min(820, self.winfo_screenwidth() - 80)
        h = min(760, self.winfo_screenheight() - 120)
        win.geometry(f"{w}x{h}")
        win.attributes("-alpha", 0.0)       # hidden until centred (no jump)

        size = self.font_size
        frame = ttk.Frame(win, padding=(12, 12, 12, 0))
        frame.pack(fill="both", expand=True)
        text = tk.Text(frame, wrap="word", relief="flat", padx=14, pady=10,
                       font=(UI_FONT, size), cursor="arrow", spacing1=2,
                       spacing3=2, bg="white")
        sb = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=sb.set)
        text.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        blue = ACTIVE_BORDER
        text.tag_configure("h1", font=(UI_FONT, size + 6, "bold"), foreground=blue,
                           spacing3=8)
        text.tag_configure("h2", font=(UI_FONT, size + 2, "bold"), foreground=blue,
                           spacing1=14, spacing3=4)
        text.tag_configure("h3", font=(UI_FONT, size, "bold"), spacing1=8)
        text.tag_configure("p", spacing3=4)
        text.tag_configure("b", lmargin1=18, lmargin2=34)
        key_col = self.list_font.measure("Ctrl+Shift+S") + 60
        text.tag_configure("k", lmargin1=18, lmargin2=key_col + 18,
                           tabs=(key_col + 18,))
        text.tag_configure("n", lmargin1=18, lmargin2=40)
        text.tag_configure("code", font=(MONO_FONT, size), background="#f2f2f2",
                           lmargin1=18, spacing1=4, spacing3=6)
        text.tag_configure("tip", background=ACTIVE_LIST_BG, lmargin1=8, lmargin2=8,
                           spacing1=6, spacing3=6)
        text.tag_configure("warn", background="#fff1e6", foreground="#b33c00",
                           font=(UI_FONT, size, "bold"), lmargin1=8, lmargin2=8,
                           spacing1=6, spacing3=6)
        for tag, line in content:
            line = for_platform(line)
            prefix = "•  " if tag == "b" else "💡  " if tag == "tip" else \
                     "⚠  " if tag == "warn" else ""
            text.insert("end", prefix + line + "\n", tag)
        text.configure(state="disabled")

        btns = ttk.Frame(win, padding=12)
        btns.pack(fill="x")
        if action:
            # A label styled as a button: macOS won't colour a real tk.Button.
            make = tk.Label(btns, text=f"  {action_text}  ", bg=ORANGE, fg="white",
                            font=self.heading_font, cursor="hand2", padx=6, pady=4)
            make.pack(side="left")
            make.bind("<Enter>", lambda e: make.configure(bg=ORANGE_DARK))
            make.bind("<Leave>", lambda e: make.configure(bg=ORANGE))
            make.bind("<ButtonRelease-1>", lambda e: action(parent=win))
        ttk.Button(btns, text="Close", command=win.destroy).pack(side="right")
        win.bind("<Escape>", lambda e: win.destroy())
        self._center_over_main(win)
        win.attributes("-alpha", 1.0)
        win.focus_set()

    def _center_over_main(self, win):
        """Centre a window over the main window, measuring the real outer
        edges (title bars, borders and the main window's menu bar)."""
        win.update()
        border = win.winfo_rootx() - win.winfo_x()
        outer_w = win.winfo_width() + 2 * border
        outer_h = win.winfo_height() + (win.winfo_rooty() - win.winfo_y()) + border
        main_border = self.winfo_rootx() - self.winfo_x()
        main_w = self.winfo_width() + 2 * main_border
        main_h = self.winfo_rooty() + self.winfo_height() + main_border - self.winfo_y()
        x = self.winfo_x() + (main_w - outer_w) // 2
        y = self.winfo_y() + (main_h - outer_h) // 2
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        win.update_idletasks()

    def _create_template(self, parent=None, use_as_database=False):
        """Write a starter .xlsx with the right headings, example rows and a
        "How to use" tab, then show simple next steps. With use_as_database
        (from the Welcome window) it also becomes the song database."""
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill
        except ImportError:
            messagebox.showerror("Create template",
                                 "Creating a spreadsheet needs the 'openpyxl' "
                                 "package:  pip install openpyxl", parent=parent)
            return
        path = filedialog.asksaveasfilename(
            parent=parent or self, title="Create template spreadsheet",
            initialdir=APP_DIR, initialfile="My Songs.xlsx",
            defaultextension=".xlsx", filetypes=[("Excel Workbook", "*.xlsx")])
        if not path:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Songs"
        ws.append(["SONG NAME", "Artist", "Style", "Vocalist"])
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="0A64C8")
        examples = [
            ("Example Song (with a web link)", "Example Artist", "Ballad, 60's",
             "Lead singer", "https://example.com/songsheets/example-song.pdf"),
            ("Another Song (no link yet)", "Another Artist", "Rock", "", ""),
        ]
        for title, artist, style, vocal, link in examples:
            ws.append([title, artist, style, vocal])
            if link:
                cell = ws.cell(row=ws.max_row, column=1)
                cell.hyperlink = link
                cell.style = "Hyperlink"
        for col, width in zip("ABCD", (42, 28, 20, 14)):
            ws.column_dimensions[col].width = width
        ws.freeze_panes = "A2"
        # A second tab with the steps - the app only reads the first tab.
        how = wb.create_sheet("How to use")
        how.column_dimensions["A"].width = 110
        how.append([f"How to build your song list for {APP_NAME}"])
        how["A1"].font = Font(bold=True, size=14, color="0A64C8")
        how.append([""])
        for number, step in enumerate(TEMPLATE_STEPS, 1):
            how.append([f"{number}.  {for_platform(step)}"])
        how.append([""])
        how.append(["The app only reads the first tab (Songs). Extra columns are "
                    "ignored, so you can keep notes in the sheet."])
        how.append(["More help in the app: Help ▸ Setting Up Your Song Spreadsheet."])
        try:
            wb.save(path)
        except OSError as exc:
            messagebox.showerror("Create template", f"Could not save:\n{exc}",
                                 parent=parent)
            return None
        self._status(f"Template created: {path}")
        if use_as_database:
            self._open_database(path)
        self._show_template_steps(path, use_as_database)
        return path

    def _show_template_steps(self, path, is_database):
        """Simple next steps after creating the template."""
        win = tk.Toplevel(self)
        win.title("Your song spreadsheet - next steps")
        win.transient(self)
        win.resizable(False, False)
        win.attributes("-alpha", 0.0)
        body = ttk.Frame(win, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="✓  Template created", font=self.title_font,
                  foreground="#2e7d32").pack(anchor="w")
        ttk.Label(body, text=str(path), foreground="#666", wraplength=560).pack(
            anchor="w", pady=(2, 12))
        ttk.Label(body, text="Now build your own song list:",
                  font=self.heading_font).pack(anchor="w")
        steps = list(TEMPLATE_STEPS)
        if not is_database:
            steps[-1] = (f"Back in {APP_NAME}, choose File ▸ Open Song Database and "
                         "pick this spreadsheet.")
        for number, step in enumerate(steps, 1):
            row = ttk.Frame(body)
            row.pack(fill="x", pady=(6, 0))
            ttk.Label(row, text=f"{number}.", font=self.heading_font, width=3,
                      foreground=ACTIVE_BORDER).pack(side="left", anchor="n")
            ttk.Label(row, text=for_platform(step), wraplength=520,
                      justify="left").pack(side="left", anchor="w")
        ttk.Label(body, foreground="#666", wraplength=560, justify="left", text=(
            "These steps are also on the “How to use” tab inside the template.")).pack(
            anchor="w", pady=(12, 0))

        btns = ttk.Frame(body)
        btns.pack(fill="x", pady=(16, 0))

        def open_it():
            try:
                open_link(path)
            except OSError as exc:
                messagebox.showerror("Open template", str(exc), parent=win)
        make = tk.Label(btns, text="  Open it in Excel  ", bg=ORANGE, fg="white",
                        font=self.heading_font, cursor="hand2", padx=6, pady=4)
        make.pack(side="left")
        make.bind("<Enter>", lambda e: make.configure(bg=ORANGE_DARK))
        make.bind("<Leave>", lambda e: make.configure(bg=ORANGE))
        make.bind("<ButtonRelease-1>", lambda e: open_it())
        ttk.Button(btns, text="Full guide",
                   command=self._show_spreadsheet_guide).pack(side="left", padx=(10, 0))
        ttk.Button(btns, text="Done", command=win.destroy).pack(side="right")
        win.bind("<Escape>", lambda e: win.destroy())
        self._center_over_main(win)
        win.attributes("-alpha", 1.0)
        win.focus_set()

    def _show_support(self):
        """Help > Support: a friendly, no-pressure invitation to donate, with
        a clickable link and a single Close button."""
        win = tk.Toplevel(self)
        win.title(f"Support {APP_NAME}")
        win.transient(self)
        win.resizable(False, False)
        win.attributes("-alpha", 0.0)
        body = ttk.Frame(win, padding=22)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text="☕  Buy us a coffee", font=self.title_font).pack(anchor="w")
        ttk.Label(body, wraplength=440, justify="left", text=(
            f"{APP_NAME} is free and open source, brought to you by JELLY JAZZ.\n\n"
            "If it's useful to you, you can buy us a coffee - any amount, any time, "
            "entirely optional. It helps keep the music and the app going.")).pack(
            anchor="w", pady=(10, 14))

        ttk.Label(body, text="Click the link to open our PayPal page:",
                  foreground="#666").pack(anchor="w")
        link_font = tkfont.Font(font=self.title_font)
        link_font.configure(underline=True, weight="normal")
        link = tk.Label(body, text=DONATE_URL.replace("https://", ""), fg=ACTIVE_BORDER,
                        cursor="hand2", font=link_font, bg=win.cget("bg"))
        link.pack(anchor="w", pady=(2, 0))
        link.bind("<Button-1>", lambda e: webbrowser.open(DONATE_URL))
        link.bind("<Enter>", lambda e: link.configure(fg=ORANGE))
        link.bind("<Leave>", lambda e: link.configure(fg=ACTIVE_BORDER))
        self._support_link_font = link_font                 # keep a reference
        ttk.Label(body, text="(opens in your web browser)", foreground="#666").pack(
            anchor="w", pady=(2, 0))

        close = ttk.Button(body, text="Close", command=win.destroy, default="active")
        close.pack(anchor="e", pady=(18, 0))
        win.bind("<Escape>", lambda e: win.destroy())
        win.bind("<Return>", lambda e: win.destroy())
        self._center_over_main(win)
        win.attributes("-alpha", 1.0)
        close.focus_set()

    def _show_about(self):
        win = tk.Toplevel(self)
        win.title(f"About {APP_NAME}")
        win.resizable(False, False)
        win.transient(self)
        body = ttk.Frame(win, padding=20)
        body.pack(fill="both", expand=True)

        try:
            self._about_logo = tk.PhotoImage(data=LOGO_PNG)  # keep a reference
            ttk.Label(body, image=self._about_logo).grid(row=0, column=0, rowspan=3,
                                                         sticky="n", padx=(0, 20))
        except tk.TclError:
            pass

        ttk.Label(body, text=APP_NAME,
                  font=self.title_font).grid(row=0, column=1, sticky="w")
        ttk.Label(body, text=f"Version {APP_VERSION}",
                  foreground="#666").grid(row=1, column=1, sticky="nw")
        ttk.Label(body, justify="left", wraplength=380, text=(
            f"Build up to {NUM_SETS} sets of {MAX_SONGS} songs from your song "
            "database, reorder them by drag and drop, open songsheets, then save "
            "and print your setlists.")).grid(row=2, column=1, sticky="nw", pady=(8, 0))

        # Credits and a clickable link to the website.
        credits = ttk.Frame(body)
        credits.grid(row=3, column=1, sticky="nw", pady=(10, 0))
        for line in APP_CREDITS:
            ttk.Label(credits, text=line, font=self.heading_font).pack(anchor="w")
        link_font = tkfont.Font(font=self.list_font)
        link_font.configure(underline=True)
        link = tk.Label(credits, text=APP_WEBSITE, fg=ACTIVE_BORDER, cursor="hand2",
                        font=link_font, bg=win.cget("bg"))
        link.pack(anchor="w", pady=(4, 0))
        link.bind("<Button-1>", lambda e: webbrowser.open(APP_WEBSITE))
        self._about_link_font = link_font                   # keep a reference
        donate = tk.Label(credits, text="☕  Enjoying it? Buy us a coffee", fg=ORANGE,
                          cursor="hand2", font=link_font, bg=win.cget("bg"))
        donate.pack(anchor="w", pady=(6, 0))
        donate.bind("<Button-1>", lambda e: webbrowser.open(DONATE_URL))

        songs = len(self.library)
        info = ttk.Frame(body)
        info.grid(row=4, column=0, columnspan=2, sticky="we", pady=(16, 0))
        for r, (label, value) in enumerate([
                ("Song database:", f"{self.db_desc or self.db_path or '(none loaded)'}"
                                   + (f"  ({songs} songs)" if songs else "")),
                ("Saved setlists:", str(SETLIST_DIR))]):
            ttk.Label(info, text=label, foreground="#666").grid(row=r, column=0, sticky="nw")
            ttk.Label(info, text=value, wraplength=420, justify="left").grid(
                row=r, column=1, sticky="w", padx=(6, 0))

        ok = ttk.Button(body, text="OK", command=win.destroy, default="active")
        ok.grid(row=5, column=0, columnspan=2, pady=(18, 0))
        win.bind("<Return>", lambda e: win.destroy())
        win.bind("<Escape>", lambda e: win.destroy())

        # Centre over the main window and make it modal.
        win.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - win.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - win.winfo_height()) // 3
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        ok.focus_set()
        win.grab_set()
        self.wait_window(win)

    # ---------- config

    def _read_config(self):
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {}

    def _write_config(self, **updates):
        cfg = self._read_config()
        cfg.update(updates)
        try:
            CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
        except OSError:
            pass

    # ---------- database

    def _choose_database(self):
        path = filedialog.askopenfilename(
            title="Open song database",
            initialdir=APP_DIR,
            filetypes=[("Spreadsheets", "*.csv *.xlsx *.xlsm"),
                       ("CSV", "*.csv"), ("Excel", "*.xlsx *.xlsm"),
                       ("All files", "*.*")])
        if path:
            self._open_database(path)

    @staticmethod
    def _stored_path(path):
        # Relative to the app folder when inside it, so the whole folder can
        # be moved (USB stick, another PC) and still work.
        try:
            return str(Path(path).resolve().relative_to(APP_DIR))
        except ValueError:
            return str(Path(path).resolve())

    def _open_database(self, path):
        """A spreadsheet chosen with File > Open Song Database becomes the
        main song database."""
        try:
            songs = load_song_database(path)
        except Exception as exc:
            messagebox.showerror("Could not load database", str(exc))
            return
        if not songs:
            messagebox.showwarning("Empty database",
                                   "No songs were found in that file.\n\n"
                                   "The first row should be a header, e.g. "
                                   "SONG NAME, Artist, Style, Vocalist.")
            return
        self._write_config(database=self._stored_path(path))
        self._use_songs(songs, path, Path(path).name, Path(path).stem)

    def _use_songs(self, songs, path, label, name, backup=False):
        """Show a freshly loaded song database."""
        self.library = songs
        self.db_path = path
        self.db_name = name
        self.db_desc = label if label == Path(path).name else f"{label}\n{path}"
        sheets = sum(1 for s in songs if s.get("link"))
        self.db_label.config(
            text=(f"{MISSING_MARK}BACKUP: " if backup else "")
                 + f"{label}  ({len(songs)} songs, {sheets} songsheets)",
            foreground=MISSING_FG if backup else "#666")
        self._refresh_library()
        self._refresh_all_sets()
        self._status(f"Loaded {len(songs)} songs from {label}")
        if any(self.sets):
            self._warn_missing()

    # ---------- song database: main file + Google Drive backup

    def _db_settings(self):
        """(main file path or None, Google Drive link, Drive-first?)"""
        cfg = self._read_config()
        local = cfg.get("database")
        if local:
            local = str(APP_DIR / local)
        elif DEFAULT_DB.exists():
            local = str(DEFAULT_DB)
        return local, cfg.get("backup_url") or "", bool(cfg.get("drive_first"))

    def _fetch_drive_copy(self, url):
        """Download the Google Drive copy; if that fails (e.g. offline), use
        the last download. Returns (path, label)."""
        cache = APP_DIR / DRIVE_CACHE
        try:
            download_drive_xlsx(url, cache)
            when = dt.datetime.now()
            return str(cache), f"Google Drive copy (downloaded {when:%d %b %Y %H:%M})"
        except OSError as exc:
            if not cache.exists():
                raise OSError(f"Google Drive copy: {exc}") from None
            when = dt.datetime.fromtimestamp(cache.stat().st_mtime)
            when = f"{when:%d %b %Y %H:%M}"
            self._drive_problem = str(exc)
            return str(cache), f"Google Drive copy (offline - saved {when})"

    def _refresh_drive_copy_quietly(self, url):
        """Keep the saved Google Drive copy up to date in the background, so a
        backup is ready even offline. No messages - it's only a spare copy."""
        def work():
            try:
                download_drive_xlsx(url, APP_DIR / DRIVE_CACHE)
            except OSError:
                pass
        threading.Thread(target=work, daemon=True).start()

    def _load_database(self, startup=False):
        """Load the song database from the main file and/or the Google Drive
        copy, in the order set in File > Song Database Settings. If the first
        choice fails, the other is used - and the app says so."""
        local, url, drive_first = self._db_settings()
        sources = (["file"] if local else []) + (["drive"] if url else [])
        if drive_first:
            sources.reverse()
        problems = []
        self._drive_problem = None
        for i, kind in enumerate(sources):
            if kind == "file":
                if not Path(local).exists():
                    problems.append(f"Main file not found:\n   {local}")
                    continue
                path, label = local, Path(local).name
            else:
                self.config(cursor="watch")
                self.update_idletasks()
                try:
                    path, label = self._fetch_drive_copy(url)
                except OSError as exc:
                    problems.append(str(exc))
                    continue
                finally:
                    self.config(cursor="")
            try:
                songs = load_song_database(path)
                if not songs:
                    raise ValueError("no songs found in it")
            except Exception as exc:
                problems.append(f"{label}: {exc}")
                continue
            name = Path(local).stem if local else "Song List"
            backup = i > 0
            self._use_songs(songs, path, label, name, backup=backup)
            if backup:
                self._status(f"{MISSING_MARK}Using the BACKUP song database: {label}")
                messagebox.showwarning(
                    "Using the backup song database",
                    "The main song database couldn't be loaded:\n\n"
                    + "\n".join(problems)
                    + f"\n\nUsing the backup instead:\n   {label}\n\n"
                    "Press F5 to try the main one again.")
            elif self._drive_problem:
                self._status(f"Google Drive couldn't be reached - using the copy "
                             f"saved earlier. ({self._drive_problem})")
            if kind == "file" and url:
                self._refresh_drive_copy_quietly(url)
            return True
        if problems:
            messagebox.showerror("Could not load the song database",
                                 "\n\n".join(problems)
                                 + "\n\nCheck File ▸ Song Database Settings.")
        return False

    # ---------- refresh

    def _used_ids(self):
        return {song_id(s) for st in self.sets for s in st}

    def _refresh_library(self):
        q = self.search_var.get().strip().casefold()
        self.filtered = [s for s in self.library
                         if not q or q in " ".join(v for k, v in s.items()
                                                   if k != "link").casefold()]
        used = self._used_ids()
        self.lib_box.delete(0, "end")
        for i, song in enumerate(self.filtered):
            self.lib_box.insert("end", song_label(song))
            if song_id(song) in used:
                self.lib_box.itemconfig(i, foreground=self.lib_colours["used"])
        self.lib_count.config(text=f"{len(self.filtered)} of {len(self.library)} songs")
        self._update_get_started()

    @staticmethod
    def _colour_lib_box(box, colours):
        box.configure(bg=colours["bg"], fg=colours["fg"],
                      selectbackground=colours["select"])

    def _colour_library_frame(self, colours):
        """Colour the border around the "Song Library" frame."""
        self.lib_border.configure(bg=colours["lib_frame"])

    def _apply_library_colours(self, colours):
        """Recolour the Song Library (live, while picking)."""
        self.lib_colours = dict(colours)
        self._colour_lib_box(self.lib_box, colours)
        self._colour_library_frame(colours)
        if hasattr(self, "_get_started"):        # rebuild the panel in the new colours
            self._get_started.destroy()
            del self._get_started
        self._refresh_library()
        self._set_active(self.active_set)        # frame + background of the current set

    def _library_colours_dialog(self):
        """Font Size > Song Library Colours: pick the four library colours with
        a live preview on the real library. OK saves, Cancel puts them back."""
        if getattr(self, "_colour_win", None) and self._colour_win.winfo_exists():
            self._colour_win.lift()
            return
        original = dict(self.lib_colours)
        working = dict(self.lib_colours)
        win = self._colour_win = tk.Toplevel(self)
        win.title("Songlist & Setlist Colours")
        win.transient(self)
        win.resizable(False, False)
        body = ttk.Frame(win, padding=14)
        body.pack(fill="both", expand=True)

        # Left: which colour to change, each with a small swatch.
        left = ttk.Frame(body)
        left.grid(row=0, column=0, sticky="n", padx=(0, 16))
        ttk.Label(left, text="Colour to change:", font=self.heading_font).pack(anchor="w")
        target = tk.StringVar(value="bg")
        swatches = {}
        for key, label in LIBRARY_COLOUR_NAMES:
            row = ttk.Frame(left)
            row.pack(fill="x", pady=3)
            sw = tk.Frame(row, width=22, height=22, bg=working[key], relief="solid", bd=1)
            sw.pack(side="left", padx=(0, 6))
            swatches[key] = sw
            ttk.Radiobutton(row, text=label, variable=target, value=key,
                            command=lambda: picker.set_colour(working[target.get()])
                            ).pack(side="left")
        contrast = ttk.Label(left, font=self.heading_font, justify="left")
        contrast.pack(anchor="w", pady=(16, 0))
        ttk.Label(left, foreground="#666", wraplength=230, justify="left", text=(
            "The Song Library and the active setlist change as you pick. Aim for a "
            "contrast of 4.5 or more between text and background, so songs are "
            "easy to read.")).pack(
            anchor="w", pady=(8, 0))

        def show_contrast():
            lines = []
            for name, fg, bg in (("Song Title", working["fg"], working["bg"]),
                                 ("Assigned songs", working["used"], working["bg"]),
                                 ("Active setlist text", "#000000", working["active_bg"])):
                ratio = contrast_ratio(fg, bg)
                lines.append(f"{name}: {ratio:.1f} : 1  {'✓' if ratio >= 4.5 else '✗ hard to read'}")
            contrast.config(text="\n".join(lines),
                            foreground="#1a7f37" if contrast_ratio(
                                working["fg"], working["bg"]) >= 4.5 else "#c62828")

        def changed(colour):
            working[target.get()] = colour
            swatches[target.get()].config(bg=colour)
            self._apply_library_colours(working)
            show_contrast()

        # Right: the picker.
        picker = ColourPicker(body, changed)
        picker.grid(row=0, column=1, sticky="n")
        picker.set_colour(working["bg"])
        show_contrast()

        def finish(save):
            if save:
                self._write_config(library_colours=dict(working))
                self._status("Songlist & Setlist colours saved.")
            else:
                self._apply_library_colours(original)
            win.destroy()

        def reset():
            working.update(LIBRARY_COLOUR_DEFAULTS)
            for key, sw in swatches.items():
                sw.config(bg=working[key])
            self._apply_library_colours(working)
            picker.set_colour(working[target.get()])
            show_contrast()

        btns = ttk.Frame(body)
        btns.grid(row=1, column=0, columnspan=2, sticky="we", pady=(14, 0))
        ttk.Button(btns, text="Reset to defaults", command=reset).pack(side="left")
        ttk.Button(btns, text="OK", command=lambda: finish(True)).pack(side="right")
        ttk.Button(btns, text="Cancel", command=lambda: finish(False)).pack(
            side="right", padx=(0, 6))
        win.protocol("WM_DELETE_WINDOW", lambda: finish(False))
        win.bind("<Escape>", lambda e: finish(False))

        # Beside the main window (right-hand side), so the library stays visible.
        win.update_idletasks()
        x = self.winfo_rootx() + self.winfo_width() - win.winfo_width() - 20
        y = self.winfo_rooty() + 60
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _update_get_started(self):
        """No song database yet: a "Get started" panel in the empty library and
        a clickable top bar, so the Welcome window is always easy to find."""
        if not hasattr(self, "_get_started"):
            c = self.lib_colours
            panel = tk.Frame(self.lib_box, bg=c["bg"])
            tk.Label(panel, text="No songs yet", bg=c["bg"], fg=c["fg"],
                     font=self.title_font).pack()
            tk.Label(panel, text="Tell the app where your songs are.", bg=c["bg"],
                     fg=c["used"], font=self.list_font).pack(pady=(4, 14))
            go = tk.Label(panel, text="  Get started…  ", bg=ORANGE, fg="white",
                          font=self.heading_font, cursor="hand2", padx=10, pady=6)
            go.pack()
            go.bind("<Enter>", lambda e: go.configure(bg=ORANGE_DARK))
            go.bind("<Leave>", lambda e: go.configure(bg=ORANGE))
            go.bind("<ButtonRelease-1>", lambda e: self._show_welcome())
            self._get_started = panel
        if self.library:
            self._get_started.place_forget()
            self.db_label.configure(cursor="")
        else:
            self._get_started.place(relx=0.5, rely=0.4, anchor="center")
            self.db_label.configure(
                text=f"{MISSING_MARK}No song database yet - click here to get started",
                foreground=MISSING_FG, cursor="hand2")

    def _is_missing(self, song):
        """True if a set song is no longer in the song database (only once a
        database is loaded - without one, nothing is marked)."""
        if not self.library:
            return False
        if not hasattr(self, "_library_ids") or self._library_ids[0] is not self.library:
            self._library_ids = (self.library, {song_id(s) for s in self.library})
        return song_id(song) not in self._library_ids[1]

    def _missing_songs(self):
        return [song["title"] for st in self.sets for song in st if self._is_missing(song)]

    def _refresh_set(self, n, select=None):
        box, songs = self.set_boxes[n], self.sets[n]
        box.delete(0, "end")
        for i, song in enumerate(songs, 1):
            missing = self._is_missing(song)
            mark = MISSING_MARK if missing else ""
            box.insert("end", f"{i:>2}.  {mark}{song_label(song)}")
            if missing:
                box.itemconfig("end", foreground=MISSING_FG)
        if select is not None and songs:
            select = max(0, min(select, len(songs) - 1))
            box.selection_set(select)
            box.activate(select)
            box.see(select)
        self.set_frames[n].config(
            text=f"Set {n + 1}   —   {len(songs)}/{MAX_SONGS} songs")

    def _refresh_all_sets(self):
        for n in range(NUM_SETS):
            self._refresh_set(n)
        self._refresh_totals()

    def _refresh_totals(self):
        count = sum(len(s) for s in self.sets)
        text = f"Total: {count} songs"
        if self.dirty:
            text += "   (unsaved changes)"
        self.total_label.config(text=text)

    def _changed(self, *set_numbers, select=None):
        self.dirty = True
        for n in set_numbers:
            self._refresh_set(n, select if n == set_numbers[-1] else None)
        self._refresh_library()
        self._refresh_totals()

    def _status(self, text):
        self.status.config(text=text)

    def _set_active(self, n):
        self.active_set = n
        for i, frame in enumerate(self.set_frames):
            active = i == n
            frame.configure(style="Active.TLabelframe" if active else "TLabelframe")
            colours = getattr(self, "lib_colours", LIBRARY_COLOUR_DEFAULTS)
            self.set_borders[i].configure(bg=colours["frame"] if active
                                          else self.inactive_border)
            self.set_boxes[i].configure(bg=colours["active_bg"] if active else "white")
        self.add_btn.config(text=f"Add to Set {n + 1}  ▶")
        if hasattr(self, "active_var"):
            self.active_var.set(n)

    # ---------- editing

    def _insert_song(self, n, index, song):
        if len(self.sets[n]) >= MAX_SONGS:
            self._status(f"Set {n + 1} is full ({MAX_SONGS} songs).")
            self.bell()
            return False
        where = [f"Set {i + 1}" for i, st in enumerate(self.sets)
                 if song_id(song) in {song_id(s) for s in st}]
        self.sets[n].insert(index, dict(song))
        self._changed(n, select=index)
        msg = f"Added “{song['title']}” to Set {n + 1}."
        if where:
            msg += f"  (Note: also in {', '.join(where)})"
        self._status(msg)
        return True

    def _add_selected(self):
        sel = self.lib_box.curselection()
        if not sel:
            self._status("Select a song in the library first.")
            return
        n = self.active_set
        self._insert_song(n, len(self.sets[n]), self.filtered[sel[0]])

    def _selected_index(self, n):
        sel = self.set_boxes[n].curselection()
        return sel[0] if sel else None

    def _move(self, n, delta):
        i = self._selected_index(n)
        if i is None:
            return "break"
        j = i + delta
        if 0 <= j < len(self.sets[n]):
            songs = self.sets[n]
            songs[i], songs[j] = songs[j], songs[i]
            self._changed(n, select=j)
        return "break"

    def _remove(self, n):
        i = self._selected_index(n)
        if i is None:
            return
        song = self.sets[n].pop(i)
        self._changed(n, select=i)
        self._status(f"Removed “{song['title']}” from Set {n + 1}.")

    def _clear_all_sets(self):
        if any(self.sets) and messagebox.askyesno(
                "Clear all sets", "Remove all songs from all 4 sets?"):
            for st in self.sets:
                st.clear()
            self._changed(*range(NUM_SETS))

    def _clear_set(self, n):
        if self.sets[n] and messagebox.askyesno(
                "Clear set", f"Remove all songs from Set {n + 1}?"):
            self.sets[n].clear()
            self._changed(n)

    # ---------- songsheets

    def _song_at(self, box, index):
        if box is self.lib_box:
            return self.filtered[index]
        return self.sets[self.set_boxes.index(box)][index]

    def _open_sheet(self, box):
        sel = box.curselection()
        if not sel:
            self._status("Select a song first.")
            return
        song = self._song_at(box, sel[0])
        link = song.get("link")
        if not link:
            self._status(f"“{song['title']}” has no songsheet link.")
            self.bell()
            return
        try:
            open_link(link)
            self._status(f"Opening songsheet for “{song['title']}”…")
        except OSError as exc:
            messagebox.showerror("Could not open songsheet",
                                 f"{song['title']}\n\n{link}\n\n{exc}")

    def _select_under_pointer(self, event):
        box = event.widget
        if box.size() == 0:
            return None
        idx = box.nearest(event.y)
        box.selection_clear(0, "end")
        box.selection_set(idx)
        box.activate(idx)
        return idx

    def _library_menu(self, event):
        if self._select_under_pointer(event) is None:
            return
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Open Songsheet",
                         command=lambda: self._open_sheet(self.lib_box))
        menu.add_separator()
        for n in range(NUM_SETS):
            menu.add_command(label=f"Add to Set {n + 1}",
                             command=lambda n=n: (self._set_active(n), self._add_selected()))
        menu.tk_popup(event.x_root, event.y_root)

    def _set_menu(self, event, n):
        if self._select_under_pointer(event) is None:
            return
        self._set_active(n)
        box = self.set_boxes[n]
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Open Songsheet", command=lambda: self._open_sheet(box))
        menu.add_separator()
        menu.add_command(label="Move Up", command=lambda: self._move(n, -1))
        menu.add_command(label="Move Down", command=lambda: self._move(n, +1))
        menu.add_separator()
        i = box.curselection()[0]
        missing = self._is_missing(self.sets[n][i])
        if missing:
            menu.add_command(label=f"{MISSING_MARK}Not in the song database",
                             state="disabled")
        menu.add_command(label="Replace with…", command=lambda: self._replace_song(n, i))
        menu.add_command(label="Remove", command=lambda: self._remove(n))
        menu.tk_popup(event.x_root, event.y_root)

    def _replace_song(self, n, i):
        """Swap a set song for a library song, keeping its place in the set.
        The closest matching titles are listed first."""
        import difflib
        old = self.sets[n][i]
        if not self.library:
            messagebox.showinfo("Replace song", "Open a song database first.")
            return
        win = tk.Toplevel(self)
        win.title("Replace song")
        win.transient(self)
        win.geometry("620x520")
        body = ttk.Frame(win, padding=14)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=f"Replace  “{old['title']}”  in Set {n + 1}, position "
                             f"{i + 1}, with:", font=self.heading_font,
                  wraplength=580).pack(anchor="w")
        row = ttk.Frame(body)
        row.pack(fill="x", pady=(8, 4))
        ttk.Label(row, text="Search:").pack(side="left")
        query = tk.StringVar()
        entry = ttk.Entry(row, textvariable=query)
        entry.pack(side="left", fill="x", expand=True, padx=4)
        box = self._make_listbox(body)
        self._colour_lib_box(box, self.lib_colours)

        # Most likely replacements first: same artist, then shared words in
        # the title (e.g. a renamed song), then similar spelling.
        old_title = old["title"].casefold()
        old_words = set(re.findall(r"\w+", old_title)) - {"the", "a", "of", "in", "my"}
        old_artist = old.get("artist", "").casefold()

        def closeness(song):
            title = song["title"].casefold()
            same_artist = bool(old_artist) and song.get("artist", "").casefold() == old_artist
            shared = len(old_words & set(re.findall(r"\w+", title)))
            spelling = difflib.SequenceMatcher(None, old_title, title).ratio()
            return (same_artist, shared, spelling)
        ranked = sorted(self.library, key=closeness, reverse=True)
        shown = []

        def refresh(*_):
            q = query.get().strip().casefold()
            shown[:] = [x for x in ranked
                        if not q or q in " ".join(v for k, v in x.items()
                                                  if k != "link").casefold()]
            box.delete(0, "end")
            for song in shown:
                box.insert("end", song_label(song))
            if shown:
                box.selection_set(0)
        query.trace_add("write", refresh)
        refresh()

        def choose(*_):
            sel = box.curselection()
            if not sel:
                return
            new = dict(shown[sel[0]])
            win.destroy()
            self.sets[n][i] = new
            self._changed(n, select=i)
            self._status(f"Replaced “{old['title']}” with “{new['title']}” "
                         f"in Set {n + 1}.")

        btns = ttk.Frame(body)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Replace", command=choose).pack(side="left")
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="right")
        box.bind("<Double-Button-1>", choose)
        box.bind("<Return>", choose)
        entry.bind("<Return>", choose)
        entry.bind("<Down>", lambda e: (box.focus_set(), "break")[1])
        win.bind("<Escape>", lambda e: win.destroy())
        self._center_over_main(win)
        entry.focus_set()

    # ---------- drag and drop

    def _drag_start(self, event):
        box = event.widget
        self._drag = None
        if box.size() == 0:
            return
        idx = box.nearest(event.y)
        bbox = box.bbox(idx)
        if not bbox or event.y > bbox[1] + bbox[3]:
            return  # clicked below the last item
        self._drag = {"box": box, "index": idx, "x": event.x_root,
                      "y": event.y_root, "moving": False}

    def _drag_motion(self, event):
        d = self._drag
        if not d:
            return
        if not d["moving"]:
            if abs(event.x_root - d["x"]) + abs(event.y_root - d["y"]) < 6:
                return
            d["moving"] = True
            self.config(cursor="hand2")
        # Keep the source item highlighted while dragging.
        d["box"].selection_clear(0, "end")
        d["box"].selection_set(d["index"])
        return "break"

    def _drop_index(self, box, y_root):
        if box.size() == 0:
            return 0
        y = y_root - box.winfo_rooty()
        idx = box.nearest(y)
        bbox = box.bbox(idx)
        if bbox and y > bbox[1] + bbox[3] / 2:
            idx += 1
        return idx

    def _drag_release(self, event):
        d, self._drag = self._drag, None
        self.config(cursor="")
        if not d or not d["moving"]:
            return
        target = self.winfo_containing(event.x_root, event.y_root)
        src_box = d["box"]
        src = None if src_box is self.lib_box else self.set_boxes.index(src_box)

        if target is self.lib_box and src is not None:
            # Dragged from a set back to the library = remove.
            song = self.sets[src].pop(d["index"])
            self._changed(src)
            self._status(f"Removed “{song['title']}” from Set {src + 1}.")
            return
        if target not in self.set_boxes:
            return

        dst = self.set_boxes.index(target)
        index = self._drop_index(target, event.y_root)
        self._set_active(dst)

        if src is None:                         # library -> set
            self._insert_song(dst, index, self.filtered[d["index"]])
        elif src == dst:                        # reorder within a set
            songs = self.sets[src]
            song = songs.pop(d["index"])
            if index > d["index"]:
                index -= 1
            songs.insert(index, song)
            self._changed(src, select=index)
        else:                                   # set -> another set
            if len(self.sets[dst]) >= MAX_SONGS:
                self._status(f"Set {dst + 1} is full ({MAX_SONGS} songs).")
                self.bell()
                return
            song = self.sets[src].pop(d["index"])
            self.sets[dst].insert(index, song)
            self._changed(src, dst, select=index)
            self._status(f"Moved “{song['title']}” to Set {dst + 1}.")

    # ---------- save / load

    def _saved_files(self):
        return sorted(SETLIST_DIR.glob("*.json"), key=lambda p: p.stem.casefold())

    def _refresh_saved_list(self):
        self.saved_combo["values"] = [p.stem for p in self._saved_files()]

    def _save(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showinfo("Save setlist", "Please type a name for the setlist first.")
            return
        fname = safe_filename(name)
        if not fname:
            messagebox.showerror("Save setlist", "That name can't be used as a file name.")
            return
        path = SETLIST_DIR / f"{fname}.json"
        if path.exists() and self.saved_var.get() != fname:
            if not messagebox.askyesno("Overwrite?",
                                       f"A setlist called “{fname}” already exists.\n"
                                       "Replace it?"):
                return
        data = {
            "name": name,
            "saved": dt.datetime.now().isoformat(timespec="seconds"),
            "database": str(self.db_path or ""),
            "sets": [{"name": f"Set {i + 1}", "songs": s}
                     for i, s in enumerate(self.sets)],
        }
        try:
            path.write_text(json.dumps(data, indent=2, ensure_ascii=False),
                            encoding="utf-8")
        except OSError as exc:
            messagebox.showerror("Could not save setlist", str(exc))
            return
        self.dirty = False
        self._refresh_saved_list()
        self.saved_var.set(fname)
        self._refresh_totals()
        self.title(f"{APP_TITLE}  —  {name}")
        self._status(f"Saved setlist “{name}”  →  {path}")
        messagebox.showinfo("Setlist saved",
                            f"“{name}” has been saved.\n\n"
                            "You can reload it any time from the Saved setlists list.")

    def _confirm_discard(self):
        if not self.dirty:
            return True
        ans = messagebox.askyesnocancel("Unsaved changes",
                                        "Save the current setlist first?")
        if ans is None:
            return False
        if ans:
            self._save()
            return not self.dirty
        return True

    def _load_selected(self):
        fname = self.saved_var.get()
        if not fname:
            return
        if not self._confirm_discard():
            return
        path = SETLIST_DIR / f"{fname}.json"
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            messagebox.showerror("Could not load setlist", str(exc))
            return
        sets = [st.get("songs", []) for st in data.get("sets", [])]
        sets = (sets + [[] for _ in range(NUM_SETS)])[:NUM_SETS]
        self.sets = [[s for s in st if s.get("title")][:MAX_SONGS] for st in sets]
        # Refresh saved songs with the latest details (e.g. new songsheet
        # links) from the current database.
        # Match on title + artist; failing that, on the title alone when only
        # one song has it (so a corrected artist in the spreadsheet carries
        # through to setlists saved before the fix).
        current = {song_id(s): s for s in self.library}
        titles = {}
        for s in self.library:
            titles.setdefault(s["title"].casefold(), []).append(s)

        def latest(song):
            if song_id(song) in current:
                return dict(current[song_id(song)])
            same = titles.get(song["title"].casefold(), [])
            return dict(same[0]) if len(same) == 1 else song

        self.sets = [[latest(s) for s in st] for st in self.sets]
        self.name_var.set(data.get("name", fname))
        self.dirty = False
        self._refresh_all_sets()
        self._refresh_library()
        self.title(f"{APP_TITLE}  —  {self.name_var.get()}")
        self._status(f"Loaded setlist “{self.name_var.get()}”.")
        self._warn_missing()

    def _warn_missing(self):
        """Say so if any set songs are no longer in the song database."""
        missing = self._missing_songs()
        if not missing:
            return
        n = len(missing)
        listed = "\n".join(f"   •  {t}" for t in missing[:12])
        if n > 12:
            listed += f"\n   …and {n - 12} more"
        self._status(f"{MISSING_MARK}{n} song{'s' if n != 1 else ''} in this setlist "
                     "no longer in the song database (marked in orange).")
        messagebox.showwarning(
            "Songs not in the song database",
            f"{n} song{'s are' if n != 1 else ' is'} in this setlist but no longer "
            f"in the song database:\n\n{listed}\n\n"
            "They're kept in the setlist (marked ⚠ in orange) with the details "
            "saved in the setlist, which may be out of date.\n\n"
            "Right-click one to Remove it, or Replace with… another song "
            "(e.g. if it was renamed).")

    def _delete_saved(self, fname=None):
        fname = fname or self.saved_var.get()
        if not fname:
            messagebox.showinfo("Delete setlist",
                                "Choose a setlist in the Saved setlists list first.")
            return
        if messagebox.askyesno("Delete setlist",
                               f"Delete the saved setlist “{fname}”?\n"
                               "The songs currently on screen are not affected."):
            (SETLIST_DIR / f"{fname}.json").unlink(missing_ok=True)
            if self.saved_var.get() == fname:
                self.saved_var.set("")
            self._refresh_saved_list()
            self._status(f"Deleted “{fname}”.")

    def _new_setlist(self):
        if not self._confirm_discard():
            return
        self.sets = [[] for _ in range(NUM_SETS)]
        self.name_var.set("")
        self.saved_var.set("")
        self.dirty = False
        self._refresh_all_sets()
        self._refresh_library()
        self.title(APP_TITLE)
        self._status("Started a new setlist.")

    def _export(self):
        """Write a printable plain-text copy of the setlist."""
        name = self.name_var.get().strip() or "Setlist"
        path = filedialog.asksaveasfilename(
            title="Export setlist", initialdir=APP_DIR,
            initialfile=f"{safe_filename(name)}.txt", defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("CSV", "*.csv")])
        if not path:
            return
        if path.lower().endswith(".csv"):
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["Set", "#", "Song Name", "Artist", "Style", "Vocalist",
                            "Songsheet"])
                for n, songs in enumerate(self.sets, 1):
                    for i, s in enumerate(songs, 1):
                        w.writerow([n, i] + [s.get(k, "") for k in
                                             ("title", "artist", "style", "vocalist", "link")])
        else:
            Path(path).write_text(self._setlist_text(), encoding="utf-8")
        self._status(f"Exported to {path}")

    # ---------- printing (plain text, via the default printer)

    def _setlist_text(self):
        name = self.name_var.get().strip() or "Setlist"
        count = sum(len(s) for s in self.sets)
        lines = text_heading(name, f"{count} songs")
        for n, songs in enumerate(self.sets, 1):
            if not songs:
                continue
            lines.append(f"SET {n}   ({len(songs)} songs)")
            lines += song_table(songs, numbered=True,
                                size_to=[x for st in self.sets for x in st])
            lines.append("")
        return "\n".join(lines)

    def _songlist_text(self):
        songs = self.filtered
        title = self.db_name or "Song List"
        sub = f"{len(songs)} songs"
        if self.search_var.get().strip():
            sub += f" matching “{self.search_var.get().strip()}”"
        return "\n".join(text_heading(title, sub) + song_table(songs))

    def _chosen_printer(self):
        """The printer picked in File > Printer, else the system default."""
        chosen = self._read_config().get("printer")
        if chosen and chosen in installed_printers():
            return chosen
        return default_printer()

    def _print_text(self, text, filename, what):
        if getattr(self, "_printing", False):
            messagebox.showinfo("Print", "Still printing the last job - "
                                         "please wait a moment.")
            return
        printer = self._chosen_printer()
        where = f"“{printer}”" if printer else "the default printer"
        if not messagebox.askyesno(
                "Print", f"Print {what}\n\non {where}?\n\n"
                         "(To use another printer: File ▸ Printer)"):
            return
        path = Path(tempfile.gettempdir()) / f"{safe_filename(filename)}.txt"
        if printer == "Microsoft Print to PDF":
            # Write the PDF ourselves: no hidden "Save as" window, and Windows
            # doesn't switch the default printer to the PDF printer.
            self._save_pdf(text, filename, what)
            return
        to_file = None
        if IS_WIN and printer in FILE_PRINTERS:
            ext = FILE_PRINTERS[printer]
            to_file = filedialog.asksaveasfilename(
                title=f"Save as {ext[1:].upper()}", initialdir=APP_DIR,
                initialfile=f"{safe_filename(filename)}{ext}", defaultextension=ext,
                filetypes=[(f"{ext[1:].upper()} file", f"*{ext}")])
            if not to_file:
                return
        try:
            # UTF-8 with BOM so Notepad shows accented names (Mañana) correctly.
            path.write_text(text, encoding="utf-8-sig" if IS_WIN else "utf-8")
        except OSError as exc:
            messagebox.showerror("Could not print", str(exc))
            return

        pdf = None
        if not IS_WIN:
            # Mac/Linux: print a PDF made here, so the layout matches Windows
            # (plain text would be laid out by the print system instead).
            pdf = path.with_suffix(".pdf")
            try:
                make_pdf(text, Path(filename).name, pdf)
            except OSError as exc:
                messagebox.showerror("Could not print", str(exc))
                return

        # Print in the background so the app never freezes on a slow printer.
        result = {}

        def work():
            try:
                if IS_WIN:
                    print_text_file_windows(path, printer=printer, to_file=to_file)
                else:
                    print_pdf_unix(pdf, printer=printer, title=Path(filename).name)
            except OSError as exc:
                result["error"] = exc

        self._printing = True
        self.config(cursor="watch")
        self._status(f"Printing {what} on {where}…")
        worker = threading.Thread(target=work, daemon=True)
        worker.start()

        def check():
            if worker.is_alive():
                self.after(300, check)
                return
            self._printing = False
            self.config(cursor="")
            if "error" not in result:
                self._status(f"Saved {what} to {to_file}" if to_file
                             else f"Sent {what} to {where}.")
                return
            self._status("Printing failed.")
            if messagebox.askyesno("Could not print", for_platform(
                    f"{result['error']}\n\nOpen it in Notepad instead, so you can "
                    "print it from there with Ctrl+P?"), icon="warning"):
                try:
                    if IS_WIN:
                        subprocess.Popen(["notepad.exe", str(path)])
                    else:
                        open_link(path)
                except OSError as exc2:
                    messagebox.showerror("Could not open", str(exc2))

        self.after(300, check)

    def _fill_printer_menu(self):
        m = self.printer_menu
        m.delete(0, "end")
        chosen = self._read_config().get("printer") or ""
        self.printer_var.set(chosen)
        default = default_printer()
        m.add_radiobutton(label=f"System default  ({default or 'none'})",
                          variable=self.printer_var, value="",
                          command=lambda: self._set_printer(""))
        m.add_separator()
        for name in installed_printers():
            m.add_radiobutton(label=name, variable=self.printer_var, value=name,
                              command=lambda n=name: self._set_printer(n))

    def _set_printer(self, name):
        self._write_config(printer=name or None)
        self._status(f"Printing to: {name or 'the system default printer'}")

    def _print_setlist(self):
        count = sum(len(s) for s in self.sets)
        if not count:
            messagebox.showinfo("Print Setlist", "This setlist has no songs yet.")
            return
        name = self.name_var.get().strip() or "Setlist"
        self._print_text(self._setlist_text(), name,
                         f"the setlist “{name}” ({count} songs)")

    def _export_database_csv(self):
        """Save the whole song library as CSV, with each song's link written
        out in a Songsheet column (a CSV can't hold Excel's embedded links)."""
        if not self.library:
            messagebox.showinfo("Export Song Database",
                                "Open a song database first (File ▸ Open Song Database).")
            return
        stem = self.db_name or "Songs"
        # Don't double up if the name already says "with links".
        stem = re.sub(r"[\s\-(]*with links\)?$", "", stem, flags=re.IGNORECASE) or stem
        path = filedialog.asksaveasfilename(
            title="Export Song Database as CSV (with web links)",
            initialdir=Path(self.db_path).parent if self.db_path else APP_DIR,
            initialfile=f"{safe_filename(stem)} (with links).csv",
            defaultextension=".csv", filetypes=[("CSV file", "*.csv")])
        if not path:
            return
        if self.db_path and Path(path).resolve() == Path(self.db_path).resolve():
            messagebox.showerror("Export Song Database",
                                 "Please choose a different file name - that is the "
                                 "song database currently in use.")
            return
        out_dir = Path(path).resolve().parent

        def link_text(link):
            # Files inside the export folder are written as relative paths, so
            # the CSV and its songsheets can be moved together.
            if not link or re.match(r"^[a-zA-Z][\w+.-]+:", link):
                return link
            try:
                return str(Path(link).resolve().relative_to(out_dir))
            except ValueError:
                return link

        try:
            # UTF-8 with BOM so Excel shows accented names (Mañana) correctly.
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["SONG NAME", "Artist", "Style", "Vocalist", "Songsheet"])
                for s in self.library:
                    w.writerow([s["title"], s.get("artist", ""), s.get("style", ""),
                                s.get("vocalist", ""), link_text(s.get("link", ""))])
        except OSError as exc:
            messagebox.showerror("Export Song Database", f"Could not save:\n{exc}")
            return
        links = sum(1 for s in self.library if s.get("link"))
        self._status(f"Exported {len(self.library)} songs ({links} links) to {path}")
        if messagebox.askyesno(
                "Export Song Database",
                f"Exported {len(self.library)} songs, {links} with songsheet links, to:\n"
                f"{path}\n\nThe links are in the Songsheet column, so this CSV can "
                "be used as a song database on its own.\n\nOpen it now?"):
            try:
                open_link(path)
            except OSError as exc:
                messagebox.showerror("Open CSV", str(exc))

    def _save_pdf(self, text, filename, what):
        path = filedialog.asksaveasfilename(
            title="Save as PDF", initialdir=APP_DIR,
            initialfile=f"{safe_filename(filename)}.pdf", defaultextension=".pdf",
            filetypes=[("PDF file", "*.pdf")])
        if not path:
            return
        try:
            pages = make_pdf(text, Path(filename).name, path)
        except OSError as exc:
            messagebox.showerror("Save as PDF", f"Could not save:\n{exc}")
            return
        self._status(f"Saved {what} as a {pages}-page PDF: {path}")
        if messagebox.askyesno("Saved as PDF",
                               f"Saved {what} ({pages} page{'s' if pages != 1 else ''}) to:"
                               f"\n{path}\n\nOpen it now?"):
            try:
                open_link(path)
            except OSError as exc:
                messagebox.showerror("Open PDF", str(exc))

    def _save_setlist_pdf(self):
        count = sum(len(s) for s in self.sets)
        if not count:
            messagebox.showinfo("Save Setlist as PDF", "This setlist has no songs yet.")
            return
        name = self.name_var.get().strip() or "Setlist"
        self._save_pdf(self._setlist_text(), name, f"the setlist “{name}”")

    def _save_songlist_pdf(self):
        if not self.filtered:
            messagebox.showinfo("Save Song List as PDF", "There are no songs to save.")
            return
        self._save_pdf(self._songlist_text(), "Song List",
                       f"the song list ({len(self.filtered)} songs)")

    def _print_songlist(self):
        if not self.filtered:
            messagebox.showinfo("Print Song List", "There are no songs to print.")
            return
        what = f"the song list ({len(self.filtered)} songs)"
        if self.search_var.get().strip():
            what += f" — only songs matching “{self.search_var.get().strip()}”"
        self._print_text(self._songlist_text(), "Song List", what)

    # ---------- window size & position (remembered between runs)

    def _restore_window(self):
        cfg = self._read_config()
        m = re.fullmatch(r"(\d+)x(\d+)([+-]-?\d+)([+-]-?\d+)", str(cfg.get("window", "")))
        if m:
            w, h, x, y = (int(v.replace("+", "")) for v in m.groups())
            if self._fits_on_screen(w, h, x, y):
                self.geometry(f"{w}x{h}+{x}+{y}")
        self._normal_geometry = self.geometry()
        self.bind("<Configure>", self._track_geometry, add="+")
        if cfg.get("maximized"):
            self.after(0, self._maximize)

    def _maximize(self):
        try:
            self.state("zoomed")
        except tk.TclError:         # not supported on every system
            pass

    def _fits_on_screen(self, w, h, x, y):
        """True if the title bar would be visible on the current monitor layout
        (a window saved on a now-unplugged monitor falls back to the default)."""
        try:
            import ctypes
            gsm = ctypes.windll.user32.GetSystemMetrics
            left, top, width, height = gsm(76), gsm(77), gsm(78), gsm(79)
        except (AttributeError, OSError):
            left, top = 0, 0
            width, height = self.winfo_screenwidth(), self.winfo_screenheight()
        return (w >= 400 and h >= 300
                and left - w + 150 <= x <= left + width - 150
                and top <= y <= top + height - 60)

    def _track_geometry(self, event):
        # Remember the normal (not maximised/minimised) size and position.
        if event.widget is self and self.state() == "normal":
            self._normal_geometry = self.geometry()

    def _save_window(self):
        self._write_config(window=self._normal_geometry,
                           maximized=self.state() == "zoomed")

    def _on_close(self):
        if self._confirm_discard():
            self._save_window()
            self.destroy()


if __name__ == "__main__":
    SetlistApp().mainloop()
