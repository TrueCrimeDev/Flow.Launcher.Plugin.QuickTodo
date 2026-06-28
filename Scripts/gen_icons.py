#!/usr/bin/env python3
"""Generate the QuickTodo plugin icons (priority set + Outlook/tdo set).

Pure PIL geometry — no SVG, no image model. All shapes are drawn at 4x
supersample then downscaled LANCZOS for crisp anti-aliasing. Every visual
parameter lives at the top so the icons are reproducible and tweakable.

Run from the repo root:  python3 Scripts/gen_icons.py
"""
from PIL import Image, ImageDraw

SS   = 4            # supersample factor
BASE = 360         # design space
R    = BASE * SS   # render resolution (1440)

# ---- body (rounded-rect card) ----------------------------------------------
BODY      = (48, 50, 312, 310)   # outer rect in design space
OUTER_R   = 34                   # outer corner radius
BORDER_W  = 16                   # border thickness
INNER     = (BODY[0]+BORDER_W, BODY[1]+BORDER_W, BODY[2]-BORDER_W, BODY[3]-BORDER_W)
INNER_R   = OUTER_R - BORDER_W   # 18

# ---- border treatment ------------------------------------------------------
LIGHTEN  = 0.5     # blend the base border this far toward the fill (lighter)
TOP_AMP  = 12      # gradient: +lighten at top edge   (was 22; toned down)
BOT_AMP  = 22      # gradient: -darken at bottom edge

# ---- interior shapes (3 rows of check + pill) ------------------------------
ROWS       = [124, 186, 248]
CHK_SCALE  = 0.87      # checkmark size vs original
CHK_STROKE = 11.0
PILL_CX    = 219.0
PILL_LEN   = 102.0
PILL_H     = 14.0      # skinnier lines
PAD        = 0.94      # scale whole content block inward (interior padding)
CCX = CCY  = 180.0     # padding anchor = body center
WHITE      = (255, 255, 255, 255)

# ---- circle icons (tdo checkbox states) ------------------------------------
CIR_C        = 180.0   # center (x = y)
CIR_ROUT     = 164.0   # outer radius
CIR_BW       = 20.0    # ring thickness
CIR_RIN      = CIR_ROUT - CIR_BW            # inner radius (144)
DOT_R        = 38.0    # unchecked: center dot radius
CHK_C_STROKE = 24.0    # checked: stroke width
CHK_C_PTS    = [(112, 184), (156, 232), (256, 120)]   # checkmark vertices

# files -> color key, output size
COLORS = {
    "blue":    ((15, 108, 189), (18, 72, 118)),   # td  — normal todo
    "outlook": ((0, 120, 212),  (15, 52, 122)),   # tdo — Outlook blue (#0078D4)
    "high":    ((197, 42, 52),  (123, 24, 30)),
    "medium":  ((224, 124, 20), (140, 74, 8)),
    "low":     ((30, 140, 62),  (16, 84, 36)),
    "done":    ((90, 98, 110),  (52, 58, 66)),
}
FILES = [
    ("td.png",        "blue",    360),
    ("td-high.png",   "high",    360),
    ("td-medium.png", "medium",  360),
    ("td-low.png",    "low",     360),
    ("td-done.png",   "done",    360),
    ("tdo.png",       "outlook", 360),
    ("tdo-large.png", "outlook", 512),
    ("tdo-small.png", "outlook", 128),
]
CIRCLES = [
    ("tdo-checked.png",   "check"),
    ("tdo-unchecked.png", "dot"),
]


def lerp(a, b, t):
    return tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def clamp(v):
    return 0 if v < 0 else 255 if v > 255 else int(round(v))


def border_color(y360, base, top=BODY[1], bot=BODY[3]):
    t = (y360 - top) / float(bot - top)
    t = 0.0 if t < 0 else 1.0 if t > 1 else t
    amp = TOP_AMP - t * (TOP_AMP + BOT_AMP)      # +TOP at top -> -BOT at bottom
    return tuple(clamp(base[i] + amp) for i in range(3))


def tx(x): return CCX + PAD * (x - CCX)
def ty(y): return CCY + PAD * (y - CCY)


def check_pts(cy):
    pts = [(86, cy - 4), (102, cy + 14), (136, cy - 24)]
    gx = sum(p[0] for p in pts) / 3.0
    gy = sum(p[1] for p in pts) / 3.0
    return [(gx + CHK_SCALE * (x - gx), gy + CHK_SCALE * (y - gy)) for (x, y) in pts]


def render(fill, border):
    base = lerp(border, fill, LIGHTEN)

    # vertical gradient across the whole canvas (only the ring will show it)
    grad = Image.new("RGBA", (R, R), (0, 0, 0, 0))
    gd = ImageDraw.Draw(grad)
    for yy in range(R):
        gd.line([(0, yy), (R, yy)], fill=border_color(yy / SS, base) + (255,))

    mask = Image.new("L", (R, R), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [BODY[0]*SS, BODY[1]*SS, BODY[2]*SS, BODY[3]*SS],
        radius=OUTER_R*SS, fill=255)

    img = Image.new("RGBA", (R, R), (0, 0, 0, 0))
    img.paste(grad, (0, 0), mask)                       # border ring (gradient)

    d = ImageDraw.Draw(img)
    d.rounded_rectangle([INNER[0]*SS, INNER[1]*SS, INNER[2]*SS, INNER[3]*SS],
                        radius=INNER_R*SS, fill=fill + (255,))  # interior fill

    stroke = CHK_STROKE * PAD
    ph     = (PILL_H / 2.0) * PAD
    plen   = PILL_LEN * PAD
    cr     = stroke * SS / 2.0
    for cy in ROWS:
        pts = [(tx(x)*SS, ty(y)*SS) for (x, y) in check_pts(cy)]
        d.line(pts, fill=WHITE, width=int(round(stroke*SS)), joint="curve")
        for (x, y) in pts:
            d.ellipse([x-cr, y-cr, x+cr, y+cr], fill=WHITE)   # round caps
        x0 = tx(PILL_CX - plen/2.0) * SS
        x1 = tx(PILL_CX + plen/2.0) * SS
        yc = ty(cy) * SS
        rr = ph * SS
        d.rounded_rectangle([x0, yc-rr, x1, yc+rr], radius=rr, fill=WHITE)
    return img


def render_circle(fill, border, mark):
    """Circle checkbox icon: ring (lightened gradient) + interior disk + mark.

    Drawn cleanly from geometry (replaces the old in-place JPEG recolor that
    accumulated lossy compression tears on the top edge).
    """
    base = lerp(border, fill, LIGHTEN)
    top, bot = CIR_C - CIR_ROUT, CIR_C + CIR_ROUT
    c, o, i = CIR_C*SS, CIR_ROUT*SS, CIR_RIN*SS

    grad = Image.new("RGBA", (R, R), (0, 0, 0, 0))
    gd = ImageDraw.Draw(grad)
    for yy in range(R):
        gd.line([(0, yy), (R, yy)], fill=border_color(yy/SS, base, top, bot) + (255,))

    mask = Image.new("L", (R, R), 0)               # ring = outer disk - inner disk
    md = ImageDraw.Draw(mask)
    md.ellipse([c-o, c-o, c+o, c+o], fill=255)
    md.ellipse([c-i, c-i, c+i, c+i], fill=0)

    img = Image.new("RGBA", (R, R), (0, 0, 0, 0))
    img.paste(grad, (0, 0), mask)                  # ring (gradient)
    d = ImageDraw.Draw(img)
    d.ellipse([c-i, c-i, c+i, c+i], fill=fill + (255,))   # interior disk

    if mark == "check":
        pts = [(x*SS, y*SS) for (x, y) in CHK_C_PTS]
        d.line(pts, fill=WHITE, width=int(round(CHK_C_STROKE*SS)), joint="curve")
        cr = CHK_C_STROKE * SS / 2.0
        for (x, y) in pts:
            d.ellipse([x-cr, y-cr, x+cr, y+cr], fill=WHITE)   # round caps
    elif mark == "dot":
        dr = DOT_R * SS
        d.ellipse([c-dr, c-dr, c+dr, c+dr], fill=WHITE)
    return img


def main():
    cache = {}
    for fname, key, n in FILES:
        if key not in cache:
            cache[key] = render(*COLORS[key])
        out = cache[key].resize((n, n), Image.LANCZOS)
        out.save(f"Images/{fname}")
        print(f"  {fname:18} {n}x{n}  ({key})")

    for fname, mark in CIRCLES:                    # tdo checkbox states (Outlook blue)
        out = render_circle(*COLORS["outlook"], mark).resize((360, 360), Image.LANCZOS)
        out.save(f"Images/{fname}")
        print(f"  {fname:18} 360x360  (outlook/{mark})")


if __name__ == "__main__":
    main()
