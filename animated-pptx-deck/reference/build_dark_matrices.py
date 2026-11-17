"""Build matrices.pptx — an animated English presentation about matrices.

Run:  python3 build.py
Pass 1 builds slides 1-15 and renders them through PowerPoint to get
snapshots; pass 2 builds the full deck with the zoom-out ending.
Morph moves objects that live on two slides; anim.py animates the rest.
"""
import math
import os
import re
import subprocess
import sys

from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from anim import Timeline

ROOT = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(ROOT, "build", "assets")
SNAPS = os.path.join(ROOT, "build", "snaps")
OUT = os.path.join(ROOT, "matrices.pptx")

W, H = 13.333, 7.5
OX, OY = 3.9, 3.75          # origin of the coordinate plane (inches)
U = 0.6                     # one grid step (inches)
DPI = 150

# palettes #4801 (blue) + #1447 (purple)
BG = "031725"
DEEP = "022E5B"
BLUE = "499BED"
SKY = "A1C6F6"
ICE = "DAEAFF"
PURPLE = "9657A6"
PURPLE_L = "B584C2"          # lighter tint of 9657A6 for text on dark
PLUM = "572565"

SG = "Space Grotesk"
MONO = "JetBrains Mono"


# ---------------------------------------------------------------- images

def hexrgba(h, a=255):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)) + (a,)


def make_assets():
    os.makedirs(ASSETS, exist_ok=True)
    step = U * DPI

    # faint background grid aligned with the plane origin
    img = Image.new("RGBA", (int(W * DPI), int(H * DPI)), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    col = hexrgba(BLUE, 34)
    k = -20
    while k <= 40:
        x = OX * DPI + k * step
        if 0 <= x <= img.width:
            d.line([(x, 0), (x, img.height)], fill=col, width=2)
        y = OY * DPI + k * step
        if 0 <= y <= img.height:
            d.line([(0, y), (img.width, y)], fill=col, width=2)
        k += 1
    img.save(os.path.join(ASSETS, "bg.png"))

    # the plane: brighter grid + axes, 14in square centred on the origin
    n = int(22 * DPI)
    c = n / 2
    img = Image.new("RGBA", (n, n), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    for k in range(-19, 20):
        p = c + k * step
        d.line([(p, 0), (p, n)], fill=hexrgba(SKY, 95), width=3)
        d.line([(0, p), (n, p)], fill=hexrgba(SKY, 95), width=3)
    d.line([(c, 0), (c, n)], fill=hexrgba(ICE, 230), width=5)
    d.line([(0, c), (n, c)], fill=hexrgba(ICE, 230), width=5)
    img.save(os.path.join(ASSETS, "plane.png"))

    ss = 4  # supersampling for smooth curves
    px = DPI * ss

    def canvas(half):
        n = int(2 * half * px)
        im = Image.new("RGBA", (n, n), (0, 0, 0, 0))
        return im, ImageDraw.Draw(im), n / 2

    def to_px(cx, x, y):
        return (cx + x * U * px, cx - y * U * px)

    def finish(im, name):
        im = im.resize((im.width // ss, im.height // ss), Image.LANCZOS)
        im.save(os.path.join(ASSETS, name))

    # point (2, 1) with its arrow from the origin
    im, d, cx = canvas(1.6)
    tip = to_px(cx, 2, 1)
    d.line([(cx, cx), tip], fill=hexrgba(SKY, 255), width=int(0.035 * px))
    r = 0.2 * px
    d.ellipse([tip[0] - r, tip[1] - r, tip[0] + r, tip[1] + r], fill=hexrgba(ICE, 60))
    r = 0.1 * px
    d.ellipse([tip[0] - r, tip[1] - r, tip[0] + r, tip[1] + r], fill=hexrgba(ICE, 255))
    finish(im, "point.png")

    # dashed 90° arc from (2, 1) to (1, -2), clockwise
    im, d, cx = canvas(1.6)
    rad = math.sqrt(5) * U * px
    a0 = -math.degrees(math.atan2(1, 2))       # PIL angles grow clockwise
    a1 = a0 + 90
    box = [cx - rad, cx - rad, cx + rad, cx + rad]
    a = a0 + 8
    while a < a1 - 10:
        d.arc(box, a, min(a + 5, a1 - 10), fill=hexrgba(PURPLE_L, 255), width=int(0.03 * px))
        a += 9
    # arrow head at the end
    t = math.radians(a1 - 8)
    hx, hy = cx + rad * math.cos(t), cx + rad * math.sin(t)
    tx, ty = -math.sin(t), math.cos(t)          # tangent, clockwise
    nx, ny = math.cos(t), math.sin(t)
    s = 0.13 * px
    d.polygon([(hx + tx * s, hy + ty * s),
               (hx - tx * s * 0.6 + nx * s * 0.7, hy - ty * s * 0.6 + ny * s * 0.7),
               (hx - tx * s * 0.6 - nx * s * 0.7, hy - ty * s * 0.6 - ny * s * 0.7)],
              fill=hexrgba(PURPLE_L, 255))
    finish(im, "arc.png")

    # arrow-shaped figure, drawn through its corner points
    im, d, cx = canvas(2.16)
    pts = [(0.5, 0.5), (2, 0.5), (2, 0), (3.2, 1), (2, 2), (2, 1.5), (0.5, 1.5)]
    poly = [to_px(cx, x, y) for x, y in pts]
    d.polygon(poly, fill=hexrgba(BLUE, 120))
    d.line(poly + [poly[0]], fill=hexrgba(ICE, 255), width=int(0.03 * px), joint="curve")
    r = 0.065 * px
    for x, y in poly:
        d.ellipse([x - r, y - r, x + r, y + r], fill=hexrgba(ICE, 255))
    finish(im, "shape.png")


# ---------------------------------------------------------------- pptx helpers

def rgb(h):
    return RGBColor.from_string(h)


def add_alpha(clr, alpha):
    if alpha < 100:
        a = etree.SubElement(clr, qn("a:alpha"))
        a.set("val", str(int(alpha * 1000)))


def drop_style(shape):
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def text(sl, name, x, y, w, h, paras, size, font=SG, bold=False, align="l",
         anchor="t", alpha=100, spc=None, wrap=True):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.name = "!!" + name
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    if isinstance(paras, str):
        paras = [[(paras, ICE)]]
    elif isinstance(paras[0], tuple):
        paras = [paras]
    for i, runs in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        for run in runs:
            t, color = run[0], run[1]
            r = p.add_run()
            r.text = t
            f = r.font
            f.name = font
            f.size = Pt(size)
            f.bold = run[2] if len(run) > 2 else bold
            f.color.rgb = rgb(color)
            add_alpha(r._r.find(qn("a:rPr")).find(qn("a:solidFill"))[0], alpha)
            if spc:
                r._r.find(qn("a:rPr")).set("spc", str(spc))
    return tb


def rect(sl, name, x, y, w, h, fill, alpha=100, line=None, line_w=1.5, rounded=0.0):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    s = sl.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    s.name = "!!" + name
    if rounded:
        s.adjustments[0] = rounded
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(fill)
    add_alpha(s._element.spPr.find(qn("a:solidFill"))[0], alpha)
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(line_w)
    else:
        s.line.fill.background()
    drop_style(s)
    return s


def bracket(sl, name, x, y, h, side, tick, color=ICE, width=3.0):
    e = lambda v: int(Inches(v))
    t = tick if side == "l" else -tick
    pts = [(x + t, y), (x, y), (x, y + h), (x + t, y + h)]
    b = sl.shapes.build_freeform(e(pts[0][0]), e(pts[0][1]), scale=1.0)
    b.add_line_segments([(e(px), e(py)) for px, py in pts[1:]], close=False)
    s = b.convert_to_shape()
    s.name = "!!" + name
    s.fill.background()
    s.line.color.rgb = rgb(color)
    s.line.width = Pt(width)
    s.line._get_or_add_ln().set("cap", "rnd")
    drop_style(s)
    return s



def freeform(sl, name, pts, fill):
    e = lambda v: int(Inches(v))
    b = sl.shapes.build_freeform(e(pts[0][0]), e(pts[0][1]), scale=1.0)
    b.add_line_segments([(e(px), e(py)) for px, py in pts[1:]], close=True)
    s = b.convert_to_shape()
    s.name = "!!" + name
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(fill)
    s.line.fill.background()
    drop_style(s)
    return s


def matrix(sl, pre, entries, x, y, cw, ch, size, colors=None, bcolor=ICE, bw=3.0):
    """Entries + brackets. Returns {name: shape}; brackets are '<pre>L' / '<pre>R'."""
    out = {}
    rows, cols = len(entries), len(entries[0])
    for i in range(rows):
        for j in range(cols):
            v = entries[i][j]
            if v is None:
                continue
            c = colors[i][j] if isinstance(colors, list) else (colors or ICE)
            nm = f"{pre}{i + 1}{j + 1}"
            out[nm] = text(sl, nm, x + j * cw, y + i * ch, cw, ch, [(v, c)], size, font=MONO,
                           align="c", anchor="m", wrap=False)
    pad = 0.1
    tick = min(0.22, cw * 0.25)
    out[pre + "L"] = bracket(sl, pre + "L", x - pad, y, rows * ch, "l", tick, bcolor, bw)
    out[pre + "R"] = bracket(sl, pre + "R", x + cols * cw + pad, y, rows * ch, "r", tick, bcolor, bw)
    return out


def pic(sl, name, path, cx, cy, w, h, rot=0, alpha=100, line=None):
    p = sl.shapes.add_picture(path, Inches(cx - w / 2), Inches(cy - h / 2), Inches(w), Inches(h))
    p.name = "!!" + name
    p.rotation = rot
    if alpha < 100:
        blip = p._element.find(".//" + qn("a:blip"))
        a = etree.SubElement(blip, qn("a:alphaModFix"))
        a.set("amt", str(int(alpha * 1000)))
    if line:
        p.line.color.rgb = rgb(line)
        p.line.width = Pt(1)
    return p


NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def morph_xml(dur, adv=None):
    a = f' advTm="{adv}"' if adv is not None else ""
    return (
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
        f'<p:transition xmlns:p="{NS_P}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
        f'spd="slow" p14:dur="{dur}"{a}><p159:morph option="byObject"/></p:transition>'
        '</mc:Choice><mc:Fallback>'
        f'<p:transition xmlns:p="{NS_P}" spd="slow"{a}><p:fade/></p:transition>'
        '</mc:Fallback></mc:AlternateContent>')


def load_notes():
    src = open(os.path.join(ROOT, "speaker_notes.md"), encoding="utf-8").read()
    notes = {}
    for m in re.finditer(r"^### Slide (\d+).*?\n(.*?)(?=^### |^## |\Z)", src, re.S | re.M):
        body = m.group(2)
        quote = [l[2:] for l in body.splitlines() if l.startswith("> ")]
        plain = [l for l in body.splitlines() if l.strip() and not l.startswith("*")]
        notes[int(m.group(1))] = " ".join(quote) if quote else " ".join(plain)
    return notes


# ---------------------------------------------------------------- layout constants

LETTERS = "MATRICES"
LET_PT, LET_Y = 120, 2.45
SX, MY, CW = 2.67, 2.6, 1.0                     # scene 02 row of matrices
A = [["1", "2"], ["3", "1"]]
B = [["2", "0"], ["1", "3"]]
FX, FY, FW = W / 2 - 9 * 0.6 / 2, 5.0, 0.6      # big formula slots
PX = 7.75                                       # scene 03 panel text x
R = [["0", "1"], ["−1", "0"]]
S_ = [["2", "0"], ["0", "1"]]
EQ_Y, ECW, ECH = 2.6, 0.75, 0.8
RB = dict(x=9.6, y=2.6, cw=0.9, ch=0.9, size=36)
STAG = 110                                      # default stagger, ms


def letter_boxes():
    from PIL import ImageFont
    f = ImageFont.truetype(os.path.expanduser("~/Library/Fonts/SpaceGrotesk-Bold.ttf"), 1000)
    adv = [f.getlength(c) / 1000 * LET_PT / 72 for c in LETTERS]
    track = -2 / 72
    total = sum(adv) + track * (len(adv) - 1)
    x, out = (W - total) / 2, []
    for c, a in zip(LETTERS, adv):
        out.append((c, x, a))
        x += a + track
    return out


def slot(i):
    """Centre of big-formula slot i."""
    return FX + i * FW + FW / 2, FY + 0.5


def centre(x, y, w, h):
    return x + w / 2, y + h / 2


# ---------------------------------------------------------------- deck builder

class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(W), Inches(H)
        self.slides = []            # (slide, timeline, transition xml or None)

    def new(self, transition=None):
        sl = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        sl.background.fill.solid()
        sl.background.fill.fore_color.rgb = rgb(BG)
        pic(sl, "bg", os.path.join(ASSETS, "bg.png"), W / 2, H / 2, W, H)
        tl = Timeline()
        self.slides.append((sl, tl, transition))
        return sl, tl

    def finish(self, notes):
        for n, (sl, tl, tr) in enumerate(self.slides, 1):
            if tr:
                sl._element.append(etree.fromstring(tr))
            t = tl.xml()
            if t:
                sl._element.append(etree.fromstring(t))
            if n in notes:
                sl.notes_slide.notes_text_frame.text = notes[n]
        return self.prs


def label(sl, name, t, x=0.7):
    return text(sl, name, x, 0.55, 6, 0.3, [(t, SKY)], 11, font=MONO, spc=300, alpha=80)


def title(sl, name, t, x=0.7, w=11.5, size=40, y=0.95):
    return text(sl, name, x, y, w, 1.4, t, size, bold=True)


def gone_up(sl, name, t, x=0.7, w=11.5, size=40):
    """Previous slide's title, parked above the slide so Morph flies it out."""
    title(sl, name, t, x, w, size, y=-1.9)


def group(tl, shapes, kind, at, **kw):
    for s in shapes:
        tl.add(s, kind, at=at, **kw)


def stripes(sl, state, chapter=None, tl=None):
    """Three slanted bands that wipe across the screen between scenes."""
    shift = {"left": (-19.5, -18.5, -17.5), "cover": (0, 0, 0), "right": (17.5, 16.5, 15.5)}[state]
    k, y0, y1 = 1.2, -0.3, 7.8
    bands = [("stripeB", -1.2, 2.2, BLUE), ("stripeP", 0.3, 2.2, PURPLE), ("stripeM", 1.8, 13.5, DEEP)]
    for (name, x, w, col), dx in zip(bands, shift):
        x += dx
        freeform(sl, name, [(x + k, y0), (x + w + k, y0), (x + w, y1), (x, y1)], col)
    if chapter:
        num, word, size = chapter
        dx = shift[2]
        t1 = text(sl, "chapTag", 4.3 + dx, 1.75, 3, 0.4, [("PART", SKY)], 14, font=MONO, spc=500)
        t2 = text(sl, "chapNum", 4.15 + dx, 2.0, 6, 2.6, [(num, BLUE)], 190, bold=True, spc=-400)
        t3 = text(sl, "chapWord", 4.3 + dx, 4.65, 8.6, 1.2, word, size, bold=True, spc=-100)
        if tl:
            tl.add(t2, "fly", at=0, dur=650, dx=-1.6, bounce=True)
            tl.add(t3, "rise", at=180, d=0.5)
            tl.add(t1, "fade", at=320)


def calc(a1, b1, a2, b2, res):
    return [(a1, BLUE), ("·", SKY), (b1, PURPLE_L), (" + ", SKY),
            (a2, BLUE), ("·", SKY), (b2, PURPLE_L), (" = ", SKY), (res, ICE)]


def scene2_matrices(sl):
    m = matrix(sl, "A", A, SX, MY, CW, CW, 56)
    times = text(sl, "times", SX + 2.1, MY, 0.8, 2, [("×", SKY)], 44, align="c", anchor="m")
    mb = matrix(sl, "B", B, SX + 3.0, MY, CW, CW, 56)
    eq = text(sl, "eq", SX + 5.1, MY, 0.8, 2, [("=", SKY)], 44, align="c", anchor="m")
    mc = matrix(sl, "C", [[None, None], [None, None]], SX + 6.0, MY, CW, CW, 56)
    return m, times, mb, eq, mc


def highlights(sl, collapsed=False):
    rw, ch = (0.02, 0.02) if collapsed else (1.92, 1.92)
    ln = (None, None) if collapsed else (BLUE, PURPLE_L)
    r = rect(sl, "rowHL", SX + 1.0 - rw / 2, MY + 0.08, rw, 0.84, BLUE, 18, ln[0], 1.25, 0.25)
    c = rect(sl, "colHL", SX + 3.08, MY + 1.0 - ch / 2, 0.84, ch, PURPLE, 26, ln[1], 1.25, 0.25)
    return r, c


FORMULA = [("f_a1", 0, "1", BLUE), ("f_dot1", 1, "·", SKY), ("f_b1", 2, "2", PURPLE_L),
           ("f_plus", 3, "+", SKY), ("f_a2", 4, "2", BLUE), ("f_dot2", 5, "·", SKY),
           ("f_b2", 6, "1", PURPLE_L), ("f_eq", 7, "=", SKY), ("f_res", 8, "4", ICE)]


def formula(sl, upto=7, dy=0.0):
    out = {}
    for nm, i, t, c in FORMULA[:upto]:
        out[nm] = text(sl, nm, FX + i * FW, FY + dy, FW, 1.0, [(t, c)], 44, font=MONO,
                       align="c", anchor="m", wrap=False)
    return out


def plane(sl, rot=0, stretch=False):
    return pic(sl, "plane", os.path.join(ASSETS, "plane.png"), OX, OY, 22, 44 if stretch else 22, rot)


def panel(sl):
    rect(sl, "panel", 7.3, 0, W - 7.3, H, BG, 92)
    return label(sl, "label3", "03 — A MATRIX IS A MOVE", PX)


def point_equation(sl, dx=0.0, with_r=True):
    r = matrix(sl, "R", R, 7.85 + dx, EQ_Y, ECW, ECH, 30, colors=BLUE) if with_r else None
    times = text(sl, "times", 9.45 + dx, EQ_Y, 0.4, 2 * ECH, [("×", SKY)], 30, align="c", anchor="m")
    v = matrix(sl, "V", [["2"], ["1"]], 10.05 + dx, EQ_Y, ECW, ECH, 30, colors=PURPLE_L)
    eq = text(sl, "eq", 10.95 + dx, EQ_Y, 0.4, 2 * ECH, [("=", SKY)], 30, align="c", anchor="m")
    w = matrix(sl, "W", [["1"], ["−2"]], 11.55 + dx, EQ_Y, ECW, ECH, 30)
    c1 = text(sl, "calc1", PX + dx, 4.55, 5.1, 0.45, calc("0", "2", "1", "1", "1"), 20, font=MONO)
    c2 = text(sl, "calc2", PX + dx, 5.0, 5.1, 0.45, calc("−1", "2", "0", "1", "−2"), 20, font=MONO)
    return r, times, v, eq, w, c1, c2


def two_moves(sl, s_dx=0.0):
    s = matrix(sl, "S", S_, 7.95 + s_dx, EQ_Y, ECW, ECH, 30, colors=BLUE)
    t = text(sl, "times2", 9.6 + s_dx, EQ_Y, 0.4, 2 * ECH, [("×", SKY)], 30, align="c", anchor="m")
    r = matrix(sl, "R", R, 10.2, EQ_Y, ECW, ECH, 30, colors=PURPLE_L)
    return s, t, r


# ---------------------------------------------------------------- the deck

def build(final):
    d = Deck()
    notes = load_notes()
    letters = letter_boxes()
    sub_runs = [("rows", BLUE), ("  ·  ", SKY), ("columns", PURPLE_L), ("  ·  ", SKY), ("moves", ICE)]
    shape_png = os.path.join(ASSETS, "shape.png")
    point_png = os.path.join(ASSETS, "point.png")
    MORPH = morph_xml(1000)

    # 1 — MATRICES: brackets slam in, letters drop one by one
    sl, tl = d.new()
    lab = label(sl, "label0", "UNIT 5 · ABSTRACT ALGEBRA")
    al = bracket(sl, "AL", 0.95, 1.2, 5.1, "l", 0.45, ICE, 4)
    ar = bracket(sl, "AR", W - 0.95, 1.2, 5.1, "r", 0.45, ICE, 4)
    tl.add(al, "fly", at=0, dur=700, dx=-2.5, bounce=True)
    tl.add(ar, "fly", at=0, dur=700, dx=2.5, bounce=True)
    for i, (c, x, a) in enumerate(letters):
        t = text(sl, f"L{i}", x, LET_Y, a + 0.5, 2.0, c, LET_PT, bold=True, anchor="m", wrap=False)
        tl.add(t, "drop", at=350 + i * 75, d=1.5)
    sub = text(sl, "sub", 0, 4.55, W, 0.5, sub_runs, 18, font=MONO, align="c")
    tl.add(sub, "rise", at=1250)
    tl.add(lab, "fade", at=1500)

    # 2 — a table of numbers (brackets morph into the matrix, letters fly up)
    sl, tl = d.new(MORPH)
    label(sl, "label0", "UNIT 5 · ABSTRACT ALGEBRA", x=-6)
    for i, (c, x, a) in enumerate(letters):
        text(sl, f"L{i}", x, -2.6 - i * 0.12, a + 0.5, 2.0, c, LET_PT, bold=True, anchor="m", wrap=False)
    text(sl, "sub", 0, 8.2, W, 0.5, sub_runs, 18, font=MONO, align="c")
    rh = rect(sl, "rowHL", 1.78, 2.43, 2.84, 1.14, BLUE, 18, BLUE, 1.25, 0.2)
    ch = rect(sl, "colHL", 1.84, 2.39, 1.32, 2.52, PURPLE, 26, PURPLE_L, 1.25, 0.2)
    m = matrix(sl, "A", A, 1.7, 2.35, 1.5, 1.3, 80)
    tl.add(label(sl, "label1", "01 — WHAT IS A MATRIX"), "rise", at=0, d=0.2)
    tl.add(title(sl, "t2", "A table of numbers."), "rise", at=100)
    for k, nm in enumerate(["A11", "A12", "A21", "A22"]):
        tl.add(m[nm], "pop", at=250 + k * 90)
    tl.add(rh, "stretchx", at=750)
    tl.add(text(sl, "rowTag", 5.15, 2.72, 2, 0.5, [("← row", BLUE)], 20, font=MONO), "fly", at=950, dx=0.8)
    tl.add(ch, "stretchy", at=1150)
    tl.add(text(sl, "colTag", 1.84, 5.4, 3, 0.5, [("↑ column", PURPLE_L)], 20, font=MONO), "rise", at=1350)
    tl.add(text(sl, "d2a", 7.4, 2.2, 1.4, 1.7, "2", 110, bold=True), "drop", at=1550, d=1.0)
    tl.add(text(sl, "dx", 8.36, 2.2, 1.2, 1.7, "×", 110, bold=True, align="c"), "spin", at=1700)
    tl.add(text(sl, "d2b", 9.65, 2.2, 1.4, 1.7, "2", 110, bold=True), "drop", at=1800, d=1.0)
    tl.add(text(sl, "sub22", 7.45, 4.0, 5.3, 0.5, [("2 rows", BLUE), ("  ·  ", SKY), ("2 columns", PURPLE_L)], 24),
           "rise", at=2100)
    tl.add(text(sl, "say22", 7.45, 4.6, 5.3, 0.5, [("We say: “two by two”.", SKY)], 20), "rise", at=2250)
    stripes(sl, "left")

    # 3 — chapter card 02
    sl, tl = d.new(morph_xml(650, adv=1700))
    stripes(sl, "cover", ("02", "MULTIPLY", 60), tl)

    # 4 — two matrices (stripes leave, matrices fly in from both sides)
    sl, tl = d.new(morph_xml(850))
    ma, times, mb, eq, mc = scene2_matrices(sl)
    tl.add(label(sl, "label2", "02 — MULTIPLY"), "rise", at=250, d=0.2)
    tl.add(title(sl, "t4", "How do we multiply?"), "rise", at=350)
    group(tl, ma.values(), "fly", 550, dur=750, dx=-3.5, bounce=True)
    tl.add(times, "spin", at=850)
    group(tl, mb.values(), "fly", 950, dur=750, dx=3.5, bounce=True)
    tl.add(eq, "pop", at=1250)
    group(tl, mc.values(), "pop", 1350)
    stripes(sl, "right", ("02", "MULTIPLY", 60))

    # 5 — row × column: numbers fly out of the matrices into the formula
    sl, tl = d.new(MORPH)
    gone_up(sl, "t4", "How do we multiply?")
    rh, ch = highlights(sl)
    scene2_matrices(sl)
    label(sl, "label2", "02 — MULTIPLY")
    tl.add(title(sl, "t5", "Row × column."), "rise", at=0)
    tl.add(rh, "stretchx", at=200)
    tl.add(ch, "stretchy", at=400)
    f = formula(sl)
    src = {"f_a1": (SX + 0.5, MY + 0.5), "f_b1": (SX + 3.5, MY + 0.5),
           "f_a2": (SX + 1.5, MY + 0.5), "f_b2": (SX + 3.5, MY + 1.5)}
    order = [("f_a1", 700), ("f_b1", 880), ("f_dot1", 1080), ("f_plus", 1250),
             ("f_a2", 1350), ("f_b2", 1530), ("f_dot2", 1730)]
    for nm, at in order:
        i = next(i for n, i, *_ in FORMULA if n == nm)
        if nm in src:
            tx, ty = slot(i)
            tl.add(f[nm], "fly", at=at, dur=650, dx=src[nm][0] - tx, dy=src[nm][1] - ty)
        else:
            tl.add(f[nm], "pop", at=at)

    # 6 — = 4: the answer flies into its cell
    sl, tl = d.new(MORPH)
    gone_up(sl, "t5", "Row × column.")
    highlights(sl)
    scene2_matrices(sl)
    label(sl, "label2", "02 — MULTIPLY")
    tl.add(title(sl, "t6", "First row, first column."), "rise", at=0)
    f = formula(sl, upto=9)
    tl.add(f["f_eq"], "pop", at=200)
    tl.add(f["f_res"], "pop", at=350, over=1.3)
    r11 = text(sl, "r11", SX + 6.0, MY, CW, CW, [("4", ICE)], 56, font=MONO, align="c", anchor="m", wrap=False)
    tx, ty = SX + 6.5, MY + 0.5
    sx_, sy_ = slot(8)
    tl.add(r11, "fly", at=750, dur=750, dx=sx_ - tx, dy=sy_ - ty)
    tl.add(r11, "pulse", at=1550, s=1.35)

    # 7 — every place: the other three cells pop with their sums
    sl, tl = d.new(MORPH)
    gone_up(sl, "t6", "First row, first column.")
    highlights(sl, collapsed=True)
    scene2_matrices(sl)
    label(sl, "label2", "02 — MULTIPLY")
    formula(sl, upto=9, dy=3.6)                    # drops out below the slide
    tl.add(title(sl, "t7", "Every place. Same rule."), "rise", at=0)
    cells = {"r11": (0, 0, "4"), "r12": (0, 1, "6"), "r21": (1, 0, "7"), "r22": (1, 1, "3")}
    rs = {nm: text(sl, nm, SX + 6.0 + j, MY + i, CW, CW, [(v, ICE)], 56, font=MONO, align="c",
                   anchor="m", wrap=False) for nm, (i, j, v) in cells.items()}
    lines = [calc("1", "2", "2", "1", "4"), calc("1", "0", "2", "3", "6"),
             calc("3", "2", "1", "1", "7"), calc("3", "0", "1", "3", "3")]
    for k, ln in enumerate(lines):
        c = text(sl, f"calc{k}", 0, 4.9 + k * 0.45, W, 0.45, ln, 22, font=MONO, align="c")
        tl.add(c, "rise", at=200 + k * 380, d=0.25)
        if k:
            tl.add(rs[["r11", "r12", "r21", "r22"][k]], "pop", at=200 + k * 380, over=1.25)
    group(tl, rs.values(), "pulse", 1650, s=1.18)
    stripes(sl, "left")

    # 8 — chapter card 03
    sl, tl = d.new(morph_xml(650, adv=1700))
    stripes(sl, "cover", ("03", "A MATRIX IS A MOVE", 52), tl)

    # 9 — a point: the background grid comes alive, the vector grows from 0
    sl, tl = d.new(morph_xml(850))
    tl.add(plane(sl), "grow", at=250, dur=1000, f=0.3)
    pt = pic(sl, "point", point_png, OX, OY, 3.2, 3.2)
    tl.add(pt, "pop", at=1050, dur=600, over=1.08)
    tl.add(text(sl, "ptLabel1", 5.25, 2.5, 1.8, 0.45, [("(2, 1)", ICE)], 20, font=MONO), "rise", at=1500, d=0.2)
    tl.add(panel(sl), "rise", at=200, d=0.2)
    tl.add(title(sl, "t9", "A point is two numbers.", PX, 5.1, 34), "rise", at=300)
    v = matrix(sl, "V", [["2"], ["1"]], 9.75, 2.55, 1.1, 0.95, 54)
    group(tl, v.values(), "pop", 1700)
    tl.add(text(sl, "cap9", PX, 4.75, 5.1, 0.5, [("2 steps right, 1 step up", SKY)], 20), "rise", at=1950)
    stripes(sl, "right", ("03", "A MATRIX IS A MOVE", 52))

    # 10 — multiply the point
    sl, tl = d.new(MORPH)
    plane(sl)
    pic(sl, "point", point_png, OX, OY, 3.2, 3.2)
    text(sl, "ptLabel1", 5.25, 2.5, 1.8, 0.45, [("(2, 1)", ICE)], 20, font=MONO)
    panel(sl)
    gone_up(sl, "t9", "A point is two numbers.", PX, 5.1, 34)
    text(sl, "cap9", PX, 8.2, 5.1, 0.5, [("2 steps right, 1 step up", SKY)], 20)
    tl.add(title(sl, "t10", "Multiply by a matrix.", PX, 5.1, 34), "rise", at=0)
    r, times, v, eq, w, c1, c2 = point_equation(sl)
    group(tl, r.values(), "fly", 200, dur=700, dx=-2.5, bounce=True)
    tl.add(times, "spin", at=500)
    tl.add(eq, "pop", at=650)
    tl.add(c1, "rise", at=900, d=0.25)
    group(tl, [w["W11"], w["WL"], w["WR"]], "pop", 1150)
    tl.add(c2, "rise", at=1450, d=0.25)
    tl.add(w["W21"], "pop", at=1700)

    # 11 — it turned (Morph turns the point; the arc draws itself)
    sl, tl = d.new(morph_xml(1200))
    plane(sl)
    arc = pic(sl, "arc", os.path.join(ASSETS, "arc.png"), OX, OY, 3.2, 3.2)
    pic(sl, "point", point_png, OX, OY, 3.2, 3.2, rot=90)
    tl.add(arc, "wheel", at=100, dur=900)
    tl.add(text(sl, "deg", 5.5, 3.95, 1.0, 0.45, [("90°", PURPLE_L)], 22, bold=True), "spin", at=850)
    tl.add(text(sl, "ptLabel2", 4.75, 5.05, 1.8, 0.45, [("(1, −2)", ICE)], 20, font=MONO), "pop", at=1000)
    panel(sl)
    gone_up(sl, "t10", "Multiply by a matrix.", PX, 5.1, 34)
    tl.add(title(sl, "t11", "It turned!", PX, 5.1, 34), "drop", at=300, d=0.9)
    point_equation(sl)

    # 12 — a shape (point shrinks back to 0, the equation slides out right)
    sl, tl = d.new(MORPH)
    plane(sl)
    pic(sl, "point", point_png, OX, OY, 0.05, 0.05, rot=90)
    pic(sl, "arc", os.path.join(ASSETS, "arc.png"), OX, OY, 0.05, 0.05)
    tl.add(pic(sl, "shape", shape_png, OX, OY, 4.32, 4.32), "pop", at=300, dur=650, over=1.06)
    panel(sl)
    gone_up(sl, "t11", "It turned!", PX, 5.1, 34)
    tl.add(title(sl, "t12", "A shape is many points.", PX, 5.1, 34), "rise", at=0)
    point_equation(sl, dx=6.5, with_r=False)
    matrix(sl, "R", R, RB["x"], RB["y"], RB["cw"], RB["ch"], RB["size"], colors=BLUE)
    tl.add(text(sl, "cap12", PX, 4.75, 5.1, 0.5, [("Same matrix. Every point.", SKY)], 20), "rise", at=700)

    # 13 — everything turns
    sl, tl = d.new(morph_xml(1300))
    plane(sl, rot=90)
    pic(sl, "shape", shape_png, OX, OY, 4.32, 4.32, rot=90)
    panel(sl)
    gone_up(sl, "t12", "A shape is many points.", PX, 5.1, 34)
    tl.add(title(sl, "t13", "Everything turns.", PX, 5.1, 34), "rise", at=0)
    matrix(sl, "R", R, RB["x"], RB["y"], RB["cw"], RB["ch"], RB["size"], colors=BLUE)
    text(sl, "cap12", PX, 8.2, 5.1, 0.5, [("Same matrix. Every point.", SKY)], 20)
    tl.add(text(sl, "cap13", PX, 4.75, 5.1, 1.0, [("Games do this millions of times per second.", SKY)], 20),
           "rise", at=250)
    matrix(sl, "S", S_, 14.5, EQ_Y, ECW, ECH, 30, colors=BLUE)     # waits off-screen

    # 14 — turn, then stretch (S flies in while everything stretches)
    sl, tl = d.new(morph_xml(1300))
    plane(sl, rot=90, stretch=True)
    pic(sl, "shape", shape_png, OX, OY, 4.32, 8.64, rot=90)
    panel(sl)
    gone_up(sl, "t13", "Everything turns.", PX, 5.1, 34)
    text(sl, "cap13", PX, 8.2, 5.1, 1.0, [("Games do this millions of times per second.", SKY)], 20)
    tl.add(title(sl, "t14", "Turn, then stretch.", PX, 5.1, 34), "rise", at=0)
    s, t2, r = two_moves(sl)
    tl.add(t2, "spin", at=200)
    tl.add(text(sl, "tagS", 7.85, 4.3, 1.7, 0.4, [("stretch", SKY)], 16, font=MONO, align="c"), "rise", at=400, d=0.2)
    tl.add(text(sl, "tagR", 10.1, 4.3, 1.7, 0.4, [("turn", SKY)], 16, font=MONO, align="c"), "rise", at=520, d=0.2)

    # 15 — two moves, one matrix
    sl, tl = d.new(MORPH)
    plane(sl, rot=90, stretch=True)
    pic(sl, "shape", shape_png, OX, OY, 4.32, 8.64, rot=90)
    panel(sl)
    gone_up(sl, "t14", "Turn, then stretch.", PX, 5.1, 34)
    tl.add(title(sl, "t15", "Two moves. One matrix.", PX, 5.1, 34), "rise", at=0)
    two_moves(sl)
    text(sl, "tagS", 7.85, 4.3, 1.7, 0.4, [("stretch", SKY)], 16, font=MONO, align="c")
    text(sl, "tagR", 10.1, 4.3, 1.7, 0.4, [("turn", SKY)], 16, font=MONO, align="c")
    tl.add(text(sl, "eq2", PX, 5.0, 0.4, 2 * ECH, [("=", SKY)], 30, align="c", anchor="m"), "pop", at=250)
    p = matrix(sl, "P", [["0", "2"], ["−1", "0"]], 8.45, 5.0, ECW, ECH, 30)
    group(tl, [p["PL"], p["PR"]], "pop", 400)
    for k, nm in enumerate(["P11", "P12", "P21", "P22"]):
        tl.add(p[nm], "pop", at=500 + k * 120, over=1.3)
    tl.add(text(sl, "capP", 10.2, 5.55, 2.6, 0.5, [("turn + stretch", ICE)], 18, font=MONO), "fly", at=1100, dx=0.8)
    group(tl, [p[n] for n in ("P11", "P12", "P21", "P22")], "pulse", 1500, s=1.2)

    if final:
        cw_, ch_, gap = 3.0, 3.0 * 9 / 16, 0.22
        gw, gh = 3 * cw_ + 2 * gap, 3 * ch_ + 2 * gap
        gx0, gy0 = (W - gw) / 2, (H - gh) / 2
        order = [1, 2, 8, 5, 15, 11, 7, 13, 14]
        zf = cw_ / W
        cells = []
        for k, n in enumerate(order):
            i, j = divmod(k, 3)
            cells.append((n, gx0 + j * (cw_ + gap) + cw_ / 2, gy0 + i * (ch_ + gap) + ch_ / 2))
        snap = lambda n: os.path.join(SNAPS, f"s-{n:02d}.png")

        def zoom(v, c, f):
            return c + (v - c) / f

        def grid_brackets(sl, f):
            bx0, bx1 = gx0 - 0.32, gx0 + gw + 0.32
            by, bh = gy0 - 0.15, gh + 0.3
            bracket(sl, "bigL", zoom(bx0, W / 2, f), zoom(by, H / 2, f), bh / f, "l", 0.35 / f, ICE, 4)
            bracket(sl, "bigR", zoom(bx1, W / 2, f), zoom(by, H / 2, f), bh / f, "r", 0.35 / f, ICE, 4)

        # 16 — frozen copy of slide 15, rest of the grid waiting off-screen
        sl, tl = d.new(f'<p:transition xmlns:p="{NS_P}" spd="fast" advTm="0"/>')
        for n, cx, cy in cells:
            pic(sl, f"cell{n}", snap(n), zoom(cx, W / 2, zf), zoom(cy, H / 2, zf), W, H, line=DEEP)
        grid_brackets(sl, zf)

        # 17 — zoom out
        sl, tl = d.new(morph_xml(2200))
        for n, cx, cy in cells:
            pic(sl, f"cell{n}", snap(n), cx, cy, cw_, ch_, line=DEEP)
        grid_brackets(sl, 1)

        # 18 — MATRICES again
        sl, tl = d.new(morph_xml(900))
        for n, cx, cy in cells:
            pic(sl, f"cell{n}", snap(n), cx, cy, cw_, ch_, alpha=15, line=DEEP)
        grid_brackets(sl, 1)
        for i, (c, x, a) in enumerate(letters):
            t = text(sl, f"L{i}", x, LET_Y, a + 0.5, 2.0, c, LET_PT, bold=True, anchor="m", wrap=False)
            tl.add(t, "drop", at=100 + i * 75, d=1.5)
        tl.add(text(sl, "sub", 0, 4.55, W, 0.5, sub_runs, 18, font=MONO, align="c"), "rise", at=1000)

    return d.finish(notes)


def render(pptx, pdf):
    script = f'''tell application "Microsoft PowerPoint"
open POSIX file "{pptx}"
delay 3
save active presentation in POSIX file "{pdf}" as save as PDF
close active presentation saving no
end tell'''
    subprocess.run(["osascript", "-e", script], check=True, timeout=180)


def main():
    make_assets()
    os.makedirs(SNAPS, exist_ok=True)
    if "--skip-snaps" not in sys.argv:
        p1 = os.path.join(ROOT, "build", "pass1.pptx")
        build(final=False).save(p1)
        pdf = os.path.join(ROOT, "build", "pass1.pdf")
        render(p1, pdf)
        for f in os.listdir(SNAPS):
            os.remove(os.path.join(SNAPS, f))
        subprocess.run(["pdftoppm", "-png", "-r", "144", pdf, os.path.join(SNAPS, "s")], check=True)
    build(final=True).save(OUT)
    print("saved", OUT)


if __name__ == "__main__":
    main()
