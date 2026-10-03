"""Generate the toolbar icons (16 px and 26 px PNG, transparent background).

Run:  python tools/make_icons.py        (needs Pillow; output is committed, so only
re-run when changing the artwork). Colours are mid-tones that stay legible on both light
and dark LibreOffice themes.
"""

import os

from PIL import Image, ImageDraw

BLUE = (61, 133, 198, 255)
DARK = (31, 78, 121, 255)
RED = (214, 69, 65, 255)
GRAY = (150, 150, 150, 255)
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src", "icons")
S = 10  # supersampling factor; artwork is drawn on a 26x26 grid


def arrow(d, direction):
    # block arrow on the 26 grid
    pts = [(4, 13), (12, 5), (12, 10), (22, 10), (22, 16), (12, 16), (12, 21)]
    if direction == "right":
        pts = [(26 - x, y) for x, y in pts]
    d.polygon([(x * S, y * S) for x, y in pts], fill=BLUE, outline=DARK)


def pencil_last_edit(d):
    # pencil body, tip bottom-left
    body = [(17, 3), (22, 8), (10, 20), (5, 21), (6, 16)]
    d.polygon([(x * S, y * S) for x, y in body], fill=BLUE, outline=DARK)
    d.polygon([(x * S, y * S) for x, y in [(15, 5), (20, 10), (22, 8), (17, 3)]], fill=RED)
    d.line([(4 * S, 24 * S), (22 * S, 24 * S)], fill=DARK, width=2 * S)


def select_visible(d):
    # three rows: visible, hidden (dashed), visible
    for y in (3, 17):
        d.rectangle([3 * S, y * S, 23 * S, (y + 6) * S], fill=BLUE, outline=DARK, width=S // 2)
    for x in range(3, 23, 4):
        d.line([(x * S, 13 * S), ((x + 2) * S, 13 * S)], fill=GRAY, width=S // 2)


def dedupe(d):
    d.rectangle([3 * S, 3 * S, 16 * S, 16 * S], fill=BLUE, outline=DARK, width=S // 2)
    d.rectangle([10 * S, 10 * S, 23 * S, 23 * S], fill=(255, 255, 255, 235), outline=DARK, width=S // 2)
    d.line([(13 * S, 13 * S), (20 * S, 20 * S)], fill=RED, width=2 * S)
    d.line([(20 * S, 13 * S), (13 * S, 20 * S)], fill=RED, width=2 * S)


YELLOW = (250, 204, 21, 255)
PURPLE = (120, 80, 170, 255)


def setops(d):
    # two overlapping circles (Venn); the lens is filled purple
    import math
    ca, cb, r = (10.0, 13.0), (16.0, 13.0), 7.5
    d.ellipse([(ca[0] - r) * S, (ca[1] - r) * S, (ca[0] + r) * S, (ca[1] + r) * S], fill=BLUE, outline=DARK)
    d.ellipse([(cb[0] - r) * S, (cb[1] - r) * S, (cb[0] + r) * S, (cb[1] + r) * S], fill=RED, outline=DARK)
    # lens polygon: points on circle A that are inside B, then points of B inside A
    pts = []
    for i in range(0, 360, 4):
        x, y = ca[0] + r * math.cos(math.radians(i)), ca[1] + r * math.sin(math.radians(i))
        if math.hypot(x - cb[0], y - cb[1]) <= r:
            pts.append((x, y))
    for i in range(180, 540, 4):
        x, y = cb[0] + r * math.cos(math.radians(i)), cb[1] + r * math.sin(math.radians(i))
        if math.hypot(x - ca[0], y - ca[1]) <= r:
            pts.append((x, y))
    if len(pts) > 2:
        cx, cy = sum(p[0] for p in pts) / len(pts), sum(p[1] for p in pts) / len(pts)
        pts.sort(key=lambda p: math.atan2(p[1] - cy, p[0] - cx))
        d.polygon([(x * S, y * S) for x, y in pts], fill=PURPLE)


def fill_blanks(d):
    # filled cell on top, arrow down, two empty (dashed) cells below
    d.rectangle([5 * S, 2 * S, 21 * S, 8 * S], fill=BLUE, outline=DARK, width=S // 2)
    d.polygon([(10 * S, 10 * S), (16 * S, 10 * S), (13 * S, 14 * S)], fill=DARK)
    for y in (15, 21):
        d.rectangle([5 * S, y * S, 21 * S, (y + 4) * S], outline=BLUE, width=S // 2)


def remove_blanks(d):
    # rows: filled, blank (red minus), filled
    for y in (2, 18):
        d.rectangle([3 * S, y * S, 23 * S, (y + 6) * S], fill=BLUE, outline=DARK, width=S // 2)
    d.rectangle([3 * S, 10 * S, 23 * S, 16 * S], outline=GRAY, width=S // 2)
    d.line([(8 * S, 13 * S), (18 * S, 13 * S)], fill=RED, width=2 * S)


def highlight(d):
    # a 3x2 grid of cells with two highlighted yellow
    for r in range(3):
        for c in range(2):
            x, y = 3 + c * 10, 3 + r * 7
            fill = YELLOW if (r, c) in ((0, 1), (2, 0)) else None
            d.rectangle([x * S, y * S, (x + 9) * S, (y + 6) * S], fill=fill, outline=DARK, width=S // 2)


def clear_highlight(d):
    highlight(d)
    d.line([(4 * S, 22 * S), (22 * S, 4 * S)], fill=RED, width=2 * S)


def transpose(d):
    # 3x3 grid: header row blue, first column amber; a diagonal arrow shows the swap
    for r in range(3):
        for c in range(3):
            fill = BLUE if r == 0 else (YELLOW if c == 0 else None)
            x, y = 2 + c * 7.5, 2 + r * 7.5
            d.rectangle([x * S, y * S, (x + 7) * S, (y + 7) * S], fill=fill, outline=DARK, width=S // 2)
    d.line([(9 * S, 9 * S), (21 * S, 21 * S)], fill=RED, width=2 * S)
    d.polygon([(23 * S, 23 * S), (15 * S, 21 * S), (21 * S, 15 * S)], fill=RED)


def sequence(d):
    # three stacked cells holding 1, 2, 3 dots
    for i in range(3):
        y = 2 + i * 8
        d.rectangle([4 * S, y * S, 22 * S, (y + 7) * S], fill=BLUE, outline=DARK, width=S // 2)
        for k in range(i + 1):
            x = 8 + k * 5
            d.ellipse([x * S, (y + 2.5) * S, (x + 2.4) * S, (y + 4.9) * S], fill=(255, 255, 255, 255))


def missing_seq(d):
    # a column of cells with a yellow placeholder row and a red plus
    for i, fill in enumerate((BLUE, BLUE, YELLOW, BLUE)):
        y = 2 + i * 6
        d.rectangle([3 * S, y * S, 17 * S, (y + 5) * S], fill=fill, outline=DARK, width=S // 2)
    d.line([(19.5 * S, 15 * S), (24.5 * S, 15 * S)], fill=RED, width=2 * S)
    d.line([(22 * S, 12.5 * S), (22 * S, 17.5 * S)], fill=RED, width=2 * S)


def gantt(d):
    # three staggered horizontal bars (one with a darker progress part) on a baseline
    for i, (x1, x2, done) in enumerate(((3, 14, 9), (8, 20, 12), (13, 23, 0))):
        y = 3 + i * 7
        d.rectangle([x1 * S, y * S, x2 * S, (y + 5) * S], fill=(157, 195, 255, 255), outline=DARK, width=S // 2)
        if done:
            d.rectangle([x1 * S, y * S, done * S, (y + 5) * S], fill=DARK)
    d.line([(2 * S, 24 * S), (24 * S, 24 * S)], fill=GRAY, width=S)


def remove_gantt(d):
    gantt(d)
    d.line([(4 * S, 22 * S), (22 * S, 4 * S)], fill=RED, width=2 * S)


def decimals(d):
    # a number with a ".00" tail and a small down arrow on the decimals
    d.rectangle([2 * S, 5 * S, 24 * S, 21 * S], fill=(255, 255, 255, 255), outline=DARK, width=S // 2)
    d.rectangle([4 * S, 8 * S, 10 * S, 18 * S], fill=BLUE)
    d.rectangle([14 * S, 14 * S, 17 * S, 17 * S], fill=DARK)
    d.rectangle([18 * S, 8 * S, 22 * S, 18 * S], fill=RED)


def to_words(d):
    # "123" on the left, arrow, three text lines on the right
    for i, x in enumerate((2, 6.5, 11)):
        d.rectangle([x * S, 4 * S, (x + 3.5) * S, 22 * S], fill=BLUE if i < 3 else None)
    d.polygon([(14.5 * S, 10 * S), (18 * S, 10 * S), (18 * S, 7 * S), (23 * S, 13 * S), (18 * S, 19 * S),
               (18 * S, 16 * S), (14.5 * S, 16 * S)], fill=RED)


def from_words(d):
    # mirrored: text lines on the left, arrow, number block on the right
    d.polygon([(11.5 * S, 10 * S), (8 * S, 10 * S), (8 * S, 7 * S), (3 * S, 13 * S), (8 * S, 19 * S),
               (8 * S, 16 * S), (11.5 * S, 16 * S)], fill=RED)
    for x in (13.5, 18):
        d.rectangle([x * S, 4 * S, (x + 3.5) * S, 22 * S], fill=BLUE)


def units(d):
    # a "1L" style tag: a rounded rectangle with a big L and a dot
    d.rectangle([2 * S, 4 * S, 24 * S, 22 * S], fill=(255, 255, 255, 255), outline=DARK, width=S // 2)
    d.rectangle([6 * S, 8 * S, 9 * S, 18 * S], fill=BLUE)
    d.rectangle([13 * S, 8 * S, 16 * S, 18 * S], fill=RED)
    d.rectangle([13 * S, 15 * S, 21 * S, 18 * S], fill=RED)


def functions(d):
    # "fx" tag: a bold f and x on a blue plate
    d.rectangle([2 * S, 4 * S, 24 * S, 22 * S], fill=BLUE, outline=DARK, width=S // 2)
    d.rectangle([6 * S, 8 * S, 9 * S, 19 * S], fill=(255, 255, 255, 255))
    d.rectangle([6 * S, 8 * S, 12 * S, 10.5 * S], fill=(255, 255, 255, 255))
    d.rectangle([5 * S, 12 * S, 11 * S, 14 * S], fill=(255, 255, 255, 255))
    d.line([(14 * S, 12 * S), (21 * S, 19 * S)], fill=(255, 255, 255, 255), width=2 * S)
    d.line([(21 * S, 12 * S), (14 * S, 19 * S)], fill=(255, 255, 255, 255), width=2 * S)


ICONS = {
    "calc_functions/functions": functions,
    "calc_decimals/decimals": decimals,
    "calc_numwords/to_words": to_words,
    "calc_numwords/from_words": from_words,
    "calc_units/units": units,
    "calc_gantt/gantt": gantt,
    "calc_gantt/remove_gantt": remove_gantt,
    "calc_sequence/sequence": sequence,
    "calc_missing_seq/missing_seq": missing_seq,
    "calc_transpose/transpose": transpose,
    "calc_setops/setops": setops,
    "calc_fill_blanks/fill_blanks": fill_blanks,
    "calc_remove_blanks/remove_blanks": remove_blanks,
    "calc_highlight/highlight": highlight,
    "calc_highlight/clear_highlight": clear_highlight,
    "navigation/back": lambda d: arrow(d, "left"),
    "navigation/forward": lambda d: arrow(d, "right"),
    "last_edit/last_edit": pencil_last_edit,
    "calc_visible/select_visible": select_visible,
    "calc_dedupe/dedupe": dedupe,
}


def render(draw_fn, size):
    img = Image.new("RGBA", (26 * S, 26 * S), (0, 0, 0, 0))
    draw_fn(ImageDraw.Draw(img))
    return img.resize((size, size), Image.LANCZOS)


def main():
    for name, fn in ICONS.items():
        for size in (16, 26):
            path = os.path.normpath(os.path.join(OUT, "%s_%d.png" % (name, size)))
            os.makedirs(os.path.dirname(path), exist_ok=True)
            render(fn, size).save(path)
            print("wrote", path)


if __name__ == "__main__":
    main()
