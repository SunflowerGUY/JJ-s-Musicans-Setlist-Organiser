<p align="center">
  <img src="assets/jjs-setlist.png" alt="JJ's Musicians Setlist Organiser logo" width="120">
</p>

<h1 align="center">JJ's Musicians Setlist Organiser — Linux</h1>

<p align="center">
  Build, reorder, save and print gig setlists from your song spreadsheet —<br>
  with one-click access to each song's PDF songsheet.<br>
  Python 3 · Tkinter · free and open source
</p>

---

![JJ's Musicians Setlist Organiser main window](https://raw.githubusercontent.com/SunflowerGUY/JJ-s-Musicans-Setlist-Organiser/main/docs/screenshot.jpg)

## Features

- **4 sets of up to 16 songs**, side by side, with drag and drop between the song library and the sets.
- **Song library from Excel (`.xlsx`) or CSV**, with instant search by song, artist, style and vocalist.
- **Songsheet links** — double-click a song in a set to open its PDF songsheet.
- **Save and load named setlists**, e.g. *"Venue Name - March 2027"*.
- **Landscape printing** (two full sets per A4 page) and **Save as PDF**.
- **Backup song database on Google Drive**, adjustable text size and colours, built-in help.

## Quick start

**1. Install Python's window toolkit** (one time only). Most Linux systems already have Python 3; the Tk toolkit is often a separate package:

| Distribution | Command |
|---|---|
| Ubuntu, Debian, Mint, Pop!_OS | `sudo apt install python3-tk python3-venv` |
| Fedora | `sudo dnf install python3-tkinter` |
| Arch, Manjaro | `sudo pacman -S tk` |
| openSUSE | `sudo zypper install python3-tk` |

**2. Get the app** — unzip `JJs Setlist - Linux.zip`, or clone this repository, and open a terminal in the folder.

**3. Run it:**

```bash
./run.sh
```

The first run installs two small add-ons (`openpyxl` for Excel files and `certifi` for secure downloads) into a private `.venv` folder, so it needs an internet connection once. Nothing is installed system-wide.

> If `./run.sh` says *Permission denied*, run `chmod +x *.sh` once (unzipping sometimes loses the "executable" setting), or use `bash run.sh`.

**4. Add it to your applications menu** (optional):

```bash
./install.sh
```

Then start it from the menu like any other app. To remove the menu entry: `./install.sh --uninstall`.

## First run

The app creates **`~/Documents/JJs Setlist`** for your settings and saved setlists, and copies in:

- the song spreadsheet that comes with it (the `.xlsx` file), which opens straight away, plus its Google Drive link as a backup if one is included, and
- any setlists in the `setlists` folder — pick one from **Saved setlists** at the top.

To use your own spreadsheet: **File ▸ Open Song Database…**. Starting from scratch? **Help ▸ Create Template Spreadsheet…** makes a ready-to-fill one.

## Printing

Choose the printer in **File ▸ Printer**. Printing uses the standard Linux print system (CUPS, the `lp` command), which desktop Linux sets up with your printers. Pages print in landscape, two sets per page. **File ▸ Save Setlist as PDF…** makes a PDF to email or print later.

## Keyboard shortcuts

The same as on Windows: **Ctrl+S** save, **Ctrl+O** open song database, **F5** reload it, **Ctrl+P** open a songsheet, **Ctrl+Shift+P** print, **Alt+↑ / Alt+↓** move a song, **Delete** remove it. The full list is in **Help ▸ Keyboard Shortcuts**.

## Updating

Replace this folder with the new version (or `git pull`), then run `./run.sh` as before. If you used `./install.sh`, run it again only if the folder moved.

Your setlists and settings in `~/Documents/JJs Setlist` are never touched by an update. New setlists that come with an update are added automatically the next time the app starts; ones you already have are never overwritten, and ones you've deleted don't come back.

## Troubleshooting

| Problem | Fix |
|---|---|
| *Python's window toolkit (Tk) is missing* | Install the Tk package for your distribution (table above). |
| *Could not set up the add-ons* | Ubuntu/Debian: `sudo apt install python3-venv`, or install the add-ons directly: `sudo apt install python3-openpyxl python3-certifi`. |
| Started from the menu and nothing appears | Run `./run.sh` in a terminal to see the message, or look in `~/.cache/jjs-setlist.log`. |
| Nothing prints | Check the printer works from another app, then choose it in **File ▸ Printer**. `lpstat -p` lists the printers Linux knows about. |
| Double-clicking `run.sh` opens it in a text editor | That's the file manager's default — use `./install.sh` and start the app from the menu instead. |

## Standalone program (optional)

To make a single program file that runs without Python (for example, to copy to other computers running the same Linux):

```bash
./build.sh
```

This produces `JJs Setlist` in this folder; run it with `./"JJs Setlist"`. It must be built on Linux, and runs on the same distribution or newer ones.

## Files

| File | Purpose |
|---|---|
| `jjs_setlist.py` | The application (the same script as the Windows and Mac versions) |
| `logo_data.py` | The logo, built into the app |
| `run.sh` | Starts the app, setting up what it needs the first time |
| `install.sh` | Adds the app to the applications menu (`--uninstall` removes it) |
| `build.sh` | Optional: builds a standalone program |
| `requirements.txt` | The Python add-ons the app uses |
| `setlists/` | Setlists that come with the package (if any), imported on first run |
| The `.xlsx` file | The starting song database |
| `database_link.txt` | Its Google Drive link, if included |
| `assets/` | The app icon |

## Support

JJ's Musicians Setlist Organiser is free and open source. If it's useful to you, you can buy us a coffee ☕ — any amount, entirely optional:
**[paypal.me/jellyjazzsoftware](https://paypal.me/jellyjazzsoftware)** (also in the app: *Help ▸ Support*).

## Licence

Released under the [MIT Licence](LICENSE): free to use, copy, change and share, as long as the copyright notice goes with it.

## Credits

Brought to you by **JELLY JAZZ**.
Vibe coding by **Adrian Newington** — [github.com/SunflowerGUY](https://github.com/SunflowerGUY) — built together with [Claude Code](https://claude.com/claude-code).
