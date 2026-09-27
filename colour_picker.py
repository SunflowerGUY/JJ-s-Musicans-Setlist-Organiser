"""Colour Tester - white text on a live background colour.

A list box of random white text beside a built-in colour picker. The list box
background updates in real time as you drag around the colour wheel, drag the
brightness bar, or type values, so you can explore dark colour combinations.
"""
import colorsys
import math
import random
import tkinter as tk
from tkinter import ttk

WORDS = (
    "alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima "
    "mike november oscar papa quebec romeo sierra tango uniform victor whiskey "
    "xray yankee zulu lorem ipsum dolor sit amet consectetur adipiscing elit "
    "mango river copper silent orbit velvet harbour ember"
).split()

# CSS named colours: name -> hex
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
NAMED = {_NAMED_SRC[i]: "#" + _NAMED_SRC[i + 1] for i in range(0, len(_NAMED_SRC), 2)}
HEX_TO_NAME = {}
for _name, _hex in NAMED.items():
    HEX_TO_NAME.setdefault(_hex, _name)

WHEEL = 240          # colour wheel diameter (px)
BAR_W, BAR_H = 22, 240
MARK_W = 10          # space left of the brightness bar for its marker


def to_hex(r, g, b):
    return f"#{r:02x}{g:02x}{b:02x}"


def hsv_to_rgb255(h, s, v):
    return tuple(int(round(c * 255)) for c in colorsys.hsv_to_rgb(h, s, v))


def contrast_with_white(r, g, b):
    """WCAG contrast ratio of white text on this background."""
    def lin(c):
        c /= 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    lum = 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)
    return 1.05 / (lum + 0.05)


class ColourTester(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Colour Tester - White Text on Dark Backgrounds")
        self.minsize(760, 500)
        self.h, self.s, self.v = 0.0, 0.0, 0.0
        self._updating = False
        self._build_ui()
        self.fill_list()
        self.set_rgb(0x1E, 0x1E, 0x2E)
        self.bind("<Escape>", lambda e: self.destroy())

    # ---------------------------------------------------------------- UI
    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # List box with random white text
        list_frame = ttk.Frame(self, padding=(10, 10, 5, 5))
        list_frame.grid(row=0, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.listbox = tk.Listbox(
            list_frame, fg="white", font=("Segoe UI", 11), bd=0,
            highlightthickness=0, activestyle="none", selectforeground="white",
        )
        self.listbox.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.listbox.config(yscrollcommand=sb.set)

        # Colour picker panel
        picker = ttk.Frame(self, padding=(5, 10, 10, 5))
        picker.grid(row=0, column=1, sticky="n")
        bg = to_hex(*(c // 256 for c in self.winfo_rgb(self.cget("bg"))))

        top = ttk.Frame(picker)
        top.grid(row=0, column=0, columnspan=2)
        self.bar = tk.Canvas(top, width=MARK_W + BAR_W, height=BAR_H, bg=bg,
                             highlightthickness=0, cursor="sb_v_double_arrow")
        self.bar.grid(row=0, column=0, padx=(0, 10))
        self.bar_img = tk.PhotoImage(width=BAR_W, height=BAR_H)
        self.bar.create_image(MARK_W, 0, image=self.bar_img, anchor="nw")
        self.bar.create_polygon(0, 0, 0, 0, 0, 0, fill="black", tags="mark")
        for ev in ("<Button-1>", "<B1-Motion>"):
            self.bar.bind(ev, self._on_bar)
        self.bar.bind("<MouseWheel>", self._on_bar_wheel)

        self.wheel = tk.Canvas(top, width=WHEEL, height=WHEEL, bg=bg,
                               highlightthickness=0, cursor="crosshair")
        self.wheel.grid(row=0, column=1)
        self.wheel_img = self._make_wheel(bg)
        self.wheel.create_image(0, 0, image=self.wheel_img, anchor="nw")
        self.wheel.create_line(0, 0, 0, 0, width=2, tags="cross")
        self.wheel.create_line(0, 0, 0, 0, width=2, tags="cross")
        for ev in ("<Button-1>", "<B1-Motion>"):
            self.wheel.bind(ev, self._on_wheel)
        self.wheel.bind("<MouseWheel>", self._on_bar_wheel)

        # Swatch + numeric fields
        mid = ttk.Frame(picker, padding=(0, 12, 0, 0))
        mid.grid(row=1, column=0, columnspan=2, sticky="w")
        self.swatch = tk.Frame(mid, width=70, height=118, relief="sunken", bd=1)
        self.swatch.grid(row=0, column=0, rowspan=4, padx=(0, 12))

        self.fields = {}
        layout = [("R", "H", "C"), ("G", "S", "M"), ("B", "V", "Y"), (None, None, "K")]
        for r, row in enumerate(layout):
            for c, name in enumerate(row):
                if name is None:
                    continue
                self._add_field(mid, name, r, 1 + c * 2)
        ttk.Label(mid, text="#").grid(row=3, column=1, sticky="e", padx=(0, 4))
        self.hex_var = tk.StringVar()
        hex_entry = ttk.Entry(mid, textvariable=self.hex_var, width=12)
        hex_entry.grid(row=3, column=2, columnspan=3, sticky="w", pady=2)
        hex_entry.bind("<KeyRelease>", self._on_hex)

        # Named colours + random
        bottom = ttk.Frame(picker, padding=(0, 10, 0, 0))
        bottom.grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Label(bottom, text="Named:").grid(row=0, column=0, padx=(0, 4))
        self.name_var = tk.StringVar()
        names = ttk.Combobox(bottom, textvariable=self.name_var, width=22,
                             values=sorted(NAMED))
        names.grid(row=0, column=1)
        names.bind("<<ComboboxSelected>>", self._on_named)
        names.bind("<Return>", self._on_named)
        ttk.Button(bottom, text="Rand", width=6, command=self.random_colour).grid(
            row=0, column=2, padx=(8, 0))
        self.dark_only = tk.BooleanVar(value=True)
        ttk.Checkbutton(bottom, text="Dark colours only (Rand)",
                        variable=self.dark_only).grid(row=1, column=1, columnspan=2,
                                                      sticky="w", pady=(6, 0))

        self.contrast = ttk.Label(picker, font=("Segoe UI", 10, "bold"))
        self.contrast.grid(row=3, column=0, columnspan=2, sticky="w", pady=(12, 0))

        # Buttons
        buttons = ttk.Frame(self, padding=(10, 5, 10, 10))
        buttons.grid(row=1, column=0, columnspan=2, sticky="ew")
        ttk.Button(buttons, text="New Text", command=self.fill_list).pack(side="left")
        ttk.Button(buttons, text="Close", command=self.destroy).pack(side="right")

    def _add_field(self, parent, name, row, col):
        maximum = 359 if name == "H" else 255
        var = tk.StringVar()
        ttk.Label(parent, text=f"{name}:").grid(row=row, column=col, sticky="e",
                                                padx=(8 if col > 1 else 0, 4))
        spin = tk.Spinbox(parent, from_=0, to=maximum, width=5, textvariable=var,
                          wrap=(name == "H"),
                          command=lambda n=name: self._on_field(n))
        spin.grid(row=row, column=col + 1, sticky="w", pady=2)
        spin.bind("<KeyRelease>", lambda e, n=name: self._on_field(n))
        spin.bind("<MouseWheel>", lambda e, s=spin: s.invoke(
            "buttonup" if e.delta > 0 else "buttondown"))
        self.fields[name] = (var, maximum)

    def _make_wheel(self, bg):
        img = tk.PhotoImage(width=WHEEL, height=WHEEL)
        c = WHEEL / 2
        radius = c - 1
        rows = []
        for y in range(WHEEL):
            dy = c - y - 0.5
            row = []
            for x in range(WHEEL):
                dx = x + 0.5 - c
                d = math.hypot(dx, dy)
                if d > radius:
                    row.append(bg)
                else:
                    h = (math.atan2(dy, dx) / (2 * math.pi)) % 1
                    row.append(to_hex(*hsv_to_rgb255(h, d / radius, 1)))
            rows.append("{" + " ".join(row) + "}")
        img.put(" ".join(rows))
        return img

    # ------------------------------------------------------------ colour
    @property
    def rgb(self):
        return hsv_to_rgb255(self.h, self.s, self.v)

    def set_rgb(self, r, g, b, source=None):
        h, s, v = colorsys.rgb_to_hsv(r / 255, g / 255, b / 255)
        if s == 0:          # grey: hue is undefined, keep the current one
            h = self.h
        if v == 0:          # black: keep saturation so the wheel marker stays put
            s = self.s
        self.set_hsv(h, s, v, source)

    def set_hsv(self, h, s, v, source=None):
        hs_changed = (h, s) != (self.h, self.s) or source is None
        self.h, self.s, self.v = h, s, v
        r, g, b = self.rgb
        colour = to_hex(r, g, b)

        # The whole point: live background update
        self.listbox.config(bg=colour,
                            selectbackground=to_hex(*hsv_to_rgb255(h, s, min(1, v + 0.25))))
        self.swatch.config(bg=colour)

        # Wheel crosshair
        c = WHEEL / 2
        x = c + s * (c - 1) * math.cos(2 * math.pi * h)
        y = c - s * (c - 1) * math.sin(2 * math.pi * h)
        line_col = "white" if v < 0.5 or s > 0.6 else "black"
        h_line, v_line = self.wheel.find_withtag("cross")
        self.wheel.coords(h_line, x - 9, y, x + 10, y)
        self.wheel.coords(v_line, x, y - 9, x, y + 10)
        self.wheel.itemconfig("cross", fill="black" if s < 0.35 else line_col)

        # Brightness bar gradient (only when hue/saturation change) + marker
        if hs_changed:
            rows = []
            for yy in range(BAR_H):
                col = to_hex(*hsv_to_rgb255(h, s, 1 - yy / (BAR_H - 1)))
                rows.append("{" + " ".join([col] * BAR_W) + "}")
            self.bar_img.put(" ".join(rows))
        my = (1 - v) * (BAR_H - 1)
        self.bar.coords("mark", 0, my - 6, MARK_W - 2, my, 0, my + 6)

        # Fields (skip the one being typed in so the cursor isn't disturbed)
        k = 1 - max(r, g, b) / 255
        if k >= 1:
            cm, mm, ym = 0, 0, 0
        else:
            cm = (1 - r / 255 - k) / (1 - k)
            mm = (1 - g / 255 - k) / (1 - k)
            ym = (1 - b / 255 - k) / (1 - k)
        values = {
            "R": r, "G": g, "B": b,
            "H": int(round(h * 360)) % 360, "S": round(s * 255), "V": round(v * 255),
            "C": round(cm * 255), "M": round(mm * 255), "Y": round(ym * 255),
            "K": round(k * 255),
        }
        self._updating = True
        try:
            for name, val in values.items():
                if name != source:
                    self.fields[name][0].set(str(val))
            if source != "hex":
                self.hex_var.set(colour[1:])
            if source != "named":
                self.name_var.set(HEX_TO_NAME.get(colour, ""))
        finally:
            self._updating = False

        ratio = contrast_with_white(r, g, b)
        aa = "pass" if ratio >= 4.5 else "fail"
        aaa = "pass" if ratio >= 7 else "fail"
        self.contrast.config(
            text=f"White text contrast: {ratio:.2f} : 1\nWCAG AA {aa}   |   AAA {aaa}",
            foreground="#1a7f37" if ratio >= 4.5 else "#c62828",
        )

    # ------------------------------------------------------------ events
    def _on_wheel(self, e):
        c = WHEEL / 2
        dx, dy = e.x - c, c - e.y
        s = min(math.hypot(dx, dy) / (c - 1), 1.0)
        h = (math.atan2(dy, dx) / (2 * math.pi)) % 1
        self.set_hsv(h, s, self.v, source="wheel")

    def _on_bar(self, e):
        v = min(max(1 - e.y / (BAR_H - 1), 0.0), 1.0)
        self.set_hsv(self.h, self.s, v, source="bar")

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
        vals = {n: int(v.get() or 0) for n, (v, _) in self.fields.items()}
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
        hex_code = NAMED.get(self.name_var.get().strip().lower())
        if hex_code:
            self.set_rgb(*(int(hex_code[i:i + 2], 16) for i in (1, 3, 5)), source="named")

    def random_colour(self):
        v = random.uniform(0.05, 0.45) if self.dark_only.get() else random.random()
        self.set_hsv(random.random(), random.random(), v)

    def fill_list(self):
        self.listbox.delete(0, "end")
        for _ in range(40):
            line = " ".join(random.choice(WORDS) for _ in range(random.randint(3, 8)))
            self.listbox.insert("end", line.capitalize())


if __name__ == "__main__":
    ColourTester().mainloop()
