#!/usr/bin/env python3
"""Build a single contact sheet of all QuickTodo icons -> icons-preview.png.

Reads the actual PNGs from Images/ (whatever produced them), so it always
reflects the current state. Run after any icon change:
    python3 Scripts/preview_icons.py
"""
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime
import os

# dark-mode palette (matches the user's design system)
BG     = (15, 15, 15)
TILE   = (18, 18, 18)
BORDER = (48, 48, 48)
TITLE  = (255, 255, 255)
LABEL  = (208, 208, 208)
MUTED  = (128, 128, 128)
ACCENT = (91, 159, 239)

ICON   = 112
TILE_W = 156
TILE_H = 156
LBL_H  = 36
GAP    = 16
MARGIN = 30
HEADER = 78
GROUP_H = 38

GROUPS = [
    ("Priority set  ·  td", [
        ("td.png", "td / default"), ("td-high.png", "high"),
        ("td-medium.png", "medium"), ("td-low.png", "low"), ("td-done.png", "done"),
    ]),
    ("Outlook set  ·  tdo", [
        ("tdo.png", "tdo"), ("tdo-large.png", "large"), ("tdo-small.png", "small"),
        ("tdo-checked.png", "checked"), ("tdo-unchecked.png", "unchecked"),
    ]),
]

COLS = max(len(g[1]) for g in GROUPS)
SHEET_W = MARGIN*2 + COLS*TILE_W + (COLS-1)*GAP
SHEET_H = HEADER + len(GROUPS)*(GROUP_H + TILE_H + LBL_H + GAP) + MARGIN


def load_font(size, bold=False):
    names = (["segoeuib.ttf", "seguisb.ttf", "arialbd.ttf"] if bold
             else ["segoeui.ttf", "arial.ttf"])
    paths = [f"/mnt/c/Windows/Fonts/{n}" for n in names] + [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
        else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]
    for p in paths:
        if os.path.exists(p):
            try: return ImageFont.truetype(p, size)
            except Exception: pass
    return ImageFont.load_default()


def main():
    f_title = load_font(34, bold=True)
    f_group = load_font(19, bold=True)
    f_label = load_font(18)
    f_size  = load_font(14)

    sheet = Image.new("RGBA", (SHEET_W, SHEET_H), BG + (255,))
    d = ImageDraw.Draw(sheet)

    d.text((MARGIN, 26), "QuickTodo Icons", font=f_title, fill=TITLE, anchor="lm")
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    d.text((SHEET_W - MARGIN, 26), stamp, font=f_size, fill=MUTED, anchor="rm")

    y = HEADER
    for gtitle, items in GROUPS:
        d.text((MARGIN, y + GROUP_H//2), gtitle, font=f_group, fill=ACCENT, anchor="lm")
        y += GROUP_H
        x = MARGIN
        for fname, label in items:
            path = f"Images/{fname}"
            d.rounded_rectangle([x, y, x+TILE_W, y+TILE_H], radius=12,
                                fill=TILE, outline=BORDER, width=1)
            if os.path.exists(path):
                ic = Image.open(path).convert("RGBA")
                native = ic.size[0]
                ic = ic.resize((ICON, ICON), Image.LANCZOS)
                ox = x + (TILE_W - ICON)//2
                oy = y + (TILE_H - ICON)//2
                sheet.paste(ic, (ox, oy), ic)
                d.text((x+TILE_W//2, y+TILE_H+8), label, font=f_label,
                       fill=LABEL, anchor="ma")
                d.text((x+TILE_W//2, y+TILE_H+8+20), f"{native}px", font=f_size,
                       fill=MUTED, anchor="ma")
            else:
                d.text((x+TILE_W//2, y+TILE_H//2), "missing", font=f_label,
                       fill=(200, 80, 80), anchor="mm")
            x += TILE_W + GAP
        y += TILE_H + LBL_H + GAP

    sheet.convert("RGB").save("icons-preview.png")
    print(f"wrote icons-preview.png  ({SHEET_W}x{SHEET_H})  @ {stamp}")


if __name__ == "__main__":
    main()
