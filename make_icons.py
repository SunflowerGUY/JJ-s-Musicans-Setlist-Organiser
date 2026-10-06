"""
Rebuild the app's icon and About-window logo from the artwork.

    ARTWORK/JJ SETLIST Icon Source.png  ->  setlist.ico   (window / .exe icon)
    ARTWORK/JJ SETLIST Logo Source.png  ->  logo_data.py  (Help > About logo,
                                                          baked into the program)
    ARTWORK/TIP LIght Bulb #1.png       ->  logo_data.py  (help-tip bulb icon)

Run:  python make_icons.py      (Build EXE.bat runs it automatically)
"""

import base64
import io
import sys
import textwrap
from pathlib import Path

from PIL import Image

HERE = Path(__file__).resolve().parent
ICON_SOURCE = HERE / "ARTWORK" / "JJ SETLIST Icon Source.png"
LOGO_SOURCE = HERE / "ARTWORK" / "JJ SETLIST Logo Source.png"
ICON_SIZES = [16, 20, 24, 32, 40, 48, 64, 128, 256]
TIP_SOURCE = HERE / "ARTWORK" / "TIP LIght Bulb #1.png"
ABOUT_SIZE = 160
TIP_HEIGHT = 48


def _png_b64(img):
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return "\n".join(textwrap.wrap(base64.b64encode(buf.getvalue()).decode(), 76))


def main():
    done = []
    if ICON_SOURCE.exists():
        # Windows icon with every common size, so it's sharp from 16 px up.
        icon = Image.open(ICON_SOURCE).convert("RGBA")
        icon.resize((256, 256), Image.LANCZOS).save(
            HERE / "setlist.ico", format="ICO", sizes=[(s, s) for s in ICON_SIZES])
        done.append(f"icon from {ICON_SOURCE.name}")
    else:
        print(f"No icon artwork at {ICON_SOURCE} - keeping the current icon.")
    if not LOGO_SOURCE.exists():
        print(f"No logo artwork at {LOGO_SOURCE} - keeping the current logo.")
        if done:
            print("Rebuilt " + ", ".join(done))
        return 0
    src = Image.open(LOGO_SOURCE).convert("RGBA")

    # About-window logo and help-tip bulb as base64 PNGs inside a Python module.
    b64 = _png_b64(src.resize((ABOUT_SIZE, ABOUT_SIZE), Image.LANCZOS))
    tip_b64 = ""
    if TIP_SOURCE.exists():
        tip = Image.open(TIP_SOURCE).convert("RGBA")
        tip_w = round(tip.width * TIP_HEIGHT / tip.height)
        tip_b64 = _png_b64(tip.resize((tip_w, TIP_HEIGHT), Image.LANCZOS))
        done.append(f"tip icon from {TIP_SOURCE.name}")
    else:
        print(f"No tip artwork at {TIP_SOURCE} - the tip icon will be left out.")
    (HERE / "logo_data.py").write_text(
        f'"""JJ\'s Setlist logo ({ABOUT_SIZE} x {ABOUT_SIZE} PNG) and help-tip bulb, '
        'baked in as\nbase64 so the program needs no image files.  Generated from '
        'ARTWORK/ by\nmake_icons.py - edit the artwork, not this '
        f'file."""\n\nLOGO_PNG = """\n{b64}\n"""\n\nTIP_PNG = """\n{tip_b64}\n"""\n',
        encoding="utf-8", newline="\n")

    done.append(f"About logo from {LOGO_SOURCE.name}")
    print("Rebuilt " + ", ".join(done))
    return 0


if __name__ == "__main__":
    sys.exit(main())
