"""Build every_photo_is_a_matrix.pptx — animated English presentation (light theme).

Run:  python3 build.py               full build, renders snapshots via PowerPoint
      python3 build.py --skip-snaps  reuse snapshots from the last full build

Morph moves objects that live on two slides; anim.py animates whatever
enters or leaves. Look: warm paper, printed photos with soft shadows,
serif display type (Source Serif 4 Display), Inter for text, JetBrains
Mono for numbers. Palette romanuke #4802.
"""
import os
import re
import subprocess
import sys

import numpy as np
from lxml import etree
from PIL import Image, ImageDraw, ImageFilter, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from anim import Timeline

ROOT = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(ROOT, "build", "assets")
SNAPS = os.path.join(ROOT, "build", "snaps")
OUT = os.path.join(ROOT, "every_photo_is_a_matrix.pptx")
FONTS = os.path.expanduser("~/Library/Fonts")

W, H = 13.333, 7.5

# palette romanuke #4802 (+ tints derived from it)
PAPER = "FAF6EE"      # canvas: lightest tint of F5E9D4
CREAM = "F5E9D4"
WHITE = "FFFDF8"      # photo print border
INK = "010A13"
TEAL = "497683"
SAGE = "8AB0A5"
GREY = "82797A"
GREY_D = "5E5758"     # darker grey for small text on paper
RED = "C47F74"        # only for the "red" layer label

SERIF = "Source Serif 4 Display Light"
SANS = "Inter"
MONO = "JetBrains Mono"

PIX_W, PIX_H = 16, 12
REG_C, REG_R = 3, 2                 # zoom region: columns 3-6, rows 2-5
Z_X, Z_Y, Z_B = 0.95, 1.85, 1.1     # zoom region on the slide, block size
RX = 6.45                           # right-hand column (pixel scene)
FX, FY, FW, FH = 0.9, 1.95, 6.0, 4.5  # photo on the filter slides
MX = 7.95                           # right-hand column (filter scene)
# prints under the title photo: (image, dx, dy, rotation, start ms)
STACK = [("blur.png", -0.45, 0.3, -12, 250), ("pix.png", 0.5, -0.15, 8, 450)]


# ---------------------------------------------------------------- images

def hexrgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


GRADIENT = [(0.0, "0C1A26"), (0.35, "2F5563"), (0.6, TEAL), (0.8, SAGE), (1.0, CREAM)]


def colorize(L):
    xs = [s for s, _ in GRADIENT]
    out = np.zeros(L.shape + (3,), np.uint8)
    for ch in range(3):
        ys = [hexrgb(c)[ch] for _, c in GRADIENT]
        out[..., ch] = np.interp(np.clip(L, 0, 1), xs, ys).astype(np.uint8)
    return Image.fromarray(out, "RGB")


def landscape(w=1200, h=900):
    """Luminance 0..1 of a small mountain scene: sun on the left, tree on the right."""
    yy, xx = np.mgrid[0:h, 0:w].astype(float)
    L = 0.28 + 0.42 * (yy / (0.62 * h)).clip(0, 1) ** 1.3
    sx, sy, sr = 0.27 * w, 0.33 * h, 0.085 * w
    d = np.hypot(xx - sx, yy - sy)
    L += 0.16 * np.exp(-(d ** 2) / (2 * (0.16 * w) ** 2))
    L = np.where(d < sr, 0.98, L)
    x = np.arange(w) / w

    def ridge(base, peaks):
        y = np.full(w, base)
        for px, ph, pw in peaks:
            y -= ph * np.clip(1 - np.abs(x - px) / pw, 0, 1)
        return (y + np.sin(x * 37) * 0.01) * h

    for ry, val in [(ridge(0.60, [(0.15, 0.10, 0.25), (0.45, 0.14, 0.2), (0.9, 0.08, 0.3)]), 0.52),
                    (ridge(0.70, [(0.62, 0.34, 0.28), (0.25, 0.08, 0.2)]), 0.34),
                    (ridge(0.84, [(0.8, 0.06, 0.4), (0.1, 0.04, 0.3)]), 0.14)]:
        L = np.where(yy > ry[None, :], val, L)
    tx = 0.83 * w
    for top, half in [(0.52, 0.05), (0.58, 0.065), (0.64, 0.08)]:
        base = top + 0.1
        inside = (yy > top * h) & (yy < base * h) & \
                 (np.abs(xx - tx) < half * w * (yy - top * h) / ((base - top) * h))
        L = np.where(inside, 0.07, L)
    L = np.where((np.abs(xx - tx) < 0.008 * w) & (yy > 0.7 * h) & (yy < 0.8 * h), 0.07, L)
    return L


def make_assets():
    os.makedirs(ASSETS, exist_ok=True)
    # paper grain: a whisper of ink noise over the canvas
    rng = np.random.default_rng(7)
    gw, gh = int(W * 120), int(H * 120)
    grain = np.zeros((gh, gw, 4), np.uint8)
    grain[..., :3] = hexrgb(INK)
    grain[..., 3] = (rng.random((gh, gw)) ** 6 * 26).astype(np.uint8)
    Image.fromarray(grain, "RGBA").save(os.path.join(ASSETS, "grain.png"))

    L = landscape()
    colorize(L).save(os.path.join(ASSETS, "photo.png"))
    colorize(L[:, ::-1]).save(os.path.join(ASSETS, "photo_mirror.png"))
    colorize(L + 50 / 255).save(os.path.join(ASSETS, "bright.png"))
    blur = np.asarray(Image.fromarray((L.clip(0, 1) * 255).astype(np.uint8))
                      .filter(ImageFilter.GaussianBlur(22))) / 255
    colorize(blur).save(os.path.join(ASSETS, "blur.png"))

    small = np.asarray(Image.fromarray((L.clip(0, 1) * 255).astype(np.uint8))
                       .resize((PIX_W, PIX_H), Image.BOX)).astype(int)
    s = 100
    pix = colorize(np.kron(small / 255, np.ones((s, s))))
    d = ImageDraw.Draw(pix)
    for c in range(1, PIX_W):
        d.line([(c * s, 0), (c * s, PIX_H * s)], fill=hexrgb(WHITE), width=3)
    for r in range(1, PIX_H):
        d.line([(0, r * s), (PIX_W * s, r * s)], fill=hexrgb(WHITE), width=3)
    pix.save(os.path.join(ASSETS, "pix.png"))
    pix.transpose(Image.FLIP_LEFT_RIGHT).save(os.path.join(ASSETS, "pix_mirror.png"))
    for c in range(PIX_W):
        pix.crop((c * s, 0, (c + 1) * s, PIX_H * s)).save(os.path.join(ASSETS, f"strip{c:02d}.png"))
    for i in range(4):
        for j in range(4):
            r, c = REG_R + i, REG_C + j
            pix.crop((c * s, r * s, (c + 1) * s, (r + 1) * s)).save(os.path.join(ASSETS, f"block{i}{j}.png"))
    colorize(np.tile(np.linspace(0, 1, 1200), (60, 1))).save(os.path.join(ASSETS, "bar.png"))
    return small


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


def shadow(shape, blur=0.22, dist=0.08, alpha=22):
    sp_pr = shape._element.spPr
    sp_pr.append(etree.fromstring(
        '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:outerShdw blurRad="{int(Inches(blur))}" dist="{int(Inches(dist))}" dir="5400000" '
        f'algn="t" rotWithShape="0"><a:srgbClr val="{INK}"><a:alpha val="{alpha * 1000}"/>'
        '</a:srgbClr></a:outerShdw></a:effectLst>'))


def text(sl, name, x, y, w, h, paras, size, font=SANS, bold=False, align="l",
         anchor="t", spc=None, wrap=True, italic=False, line=None):
    """paras: str | [(text, color[, bold[, italic]])] | [[runs], [runs]]."""
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.name = "!!" + name
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    if isinstance(paras, str):
        paras = [[(paras, INK)]]
    elif isinstance(paras[0], tuple):
        paras = [paras]
    for i, runs in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        if line:
            p.line_spacing = line
        for run in runs:
            r = p.add_run()
            r.text = run[0]
            f = r.font
            f.name = font
            f.size = Pt(size)
            f.bold = run[2] if len(run) > 2 else bold
            f.italic = run[3] if len(run) > 3 else italic
            f.color.rgb = rgb(run[1])
            if spc:
                r._r.find(qn("a:rPr")).set("spc", str(spc))
    return tb


def rect(sl, name, x, y, w, h, fill=None, alpha=100, line=None, line_w=1.5, rounded=0.0, rot=0):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    s = sl.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    s.name = "!!" + name
    s.rotation = rot
    if rounded:
        s.adjustments[0] = rounded
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
        add_alpha(s._element.spPr.find(qn("a:solidFill"))[0], alpha)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(line_w)
    else:
        s.line.fill.background()
    drop_style(s)
    return s


def oval(sl, name, cx, cy, w, h, fill=None, alpha=100, line=None, line_w=2.0, rot=0):
    s = sl.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - w / 2), Inches(cy - h / 2), Inches(w), Inches(h))
    s.name = "!!" + name
    s.rotation = rot
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
        add_alpha(s._element.spPr.find(qn("a:solidFill"))[0], alpha)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(line_w)
    else:
        s.line.fill.background()
    drop_style(s)
    return s


def polyline(sl, name, pts, color, width):
    e = lambda v: int(Inches(v))
    b = sl.shapes.build_freeform(e(pts[0][0]), e(pts[0][1]), scale=1.0)
    b.add_line_segments([(e(px), e(py)) for px, py in pts[1:]], close=False)
    s = b.convert_to_shape()
    s.name = "!!" + name
    s.fill.background()
    s.line.color.rgb = rgb(color)
    s.line.width = Pt(width)
    s.line._get_or_add_ln().set("cap", "sq")
    drop_style(s)
    return s


def bracket(sl, name, x, y, h, side, tick, color=INK, width=2.5):
    t = tick if side == "l" else -tick
    return polyline(sl, name, [(x + t, y), (x, y), (x, y + h), (x + t, y + h)], color, width)


def viewfinder(sl, x, y, w, h, a=0.38, color=TEAL, width=2.25):
    """Four camera-viewfinder corners around a box."""
    return [polyline(sl, "vf0", [(x, y + a), (x, y), (x + a, y)], color, width),
            polyline(sl, "vf1", [(x + w - a, y), (x + w, y), (x + w, y + a)], color, width),
            polyline(sl, "vf2", [(x + w, y + h - a), (x + w, y + h), (x + w - a, y + h)], color, width),
            polyline(sl, "vf3", [(x + a, y + h), (x, y + h), (x, y + h - a)], color, width)]


def matrix(sl, pre, entries, x, y, cw, ch, size, colors=None, bw=2.0, names=None, bold=False):
    out = {}
    rows, cols = len(entries), len(entries[0])
    for i in range(rows):
        for j in range(cols):
            v = entries[i][j]
            if v is None:
                continue
            c = colors[i][j] if isinstance(colors, list) else (colors or INK)
            nm = names[i][j] if names else f"{pre}{i + 1}{j + 1}"
            out[nm] = text(sl, nm, x + j * cw, y + i * ch, cw, ch, [(v, c, bold)], size, font=MONO,
                           align="c", anchor="m", wrap=False)
    tick = min(0.18, cw * 0.22)
    out[pre + "L"] = bracket(sl, pre + "L", x - 0.1, y, rows * ch, "l", tick, INK, bw)
    out[pre + "R"] = bracket(sl, pre + "R", x + cols * cw + 0.1, y, rows * ch, "r", tick, INK, bw)
    return out


def pic(sl, name, path, x, y, w, h, alpha=100, line=None, line_w=0.75, crop=None, rot=0):
    p = sl.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    p.name = "!!" + name
    p.rotation = rot
    if crop:
        p.crop_left, p.crop_top, p.crop_right, p.crop_bottom = crop
    if alpha < 100:
        blip = p._element.find(".//" + qn("a:blip"))
        a = etree.SubElement(blip, qn("a:alphaModFix"))
        a.set("amt", str(int(alpha * 1000)))
    if line:
        p.line.color.rgb = rgb(line)
        p.line.width = Pt(line_w)
    return p


def photo_print(sl, path, cx, cy, w, h, rot=0, pad=0.16, crop=None, name="photo", frame="frame"):
    """A printed photo: white border with a soft shadow, picture on top."""
    fr = rect(sl, frame, cx - w / 2 - pad, cy - h / 2 - pad, w + 2 * pad, h + 2 * pad, WHITE, rot=rot)
    shadow(fr)
    ph = pic(sl, name, path, cx - w / 2, cy - h / 2, w, h, crop=crop, rot=rot)
    return fr, ph


NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def morph_xml(dur=1000, adv=None):
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


class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(W), Inches(H)
        self.slides = []

    def new(self, transition=None):
        sl = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        sl.background.fill.solid()
        sl.background.fill.fore_color.rgb = rgb(PAPER)
        pic(sl, "grain", os.path.join(ASSETS, "grain.png"), 0, 0, W, H)
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


def label(sl, name, t, x=0.8):
    return text(sl, name, x, 0.55, 7, 0.3, [(t, GREY_D)], 10.5, font=MONO, spc=300)


def title(sl, name, runs, y=0.9, size=44, x=0.8, w=11.8):
    if isinstance(runs, str):
        runs = [(runs, INK)]
    return text(sl, name, x, y, w, 0.9, runs, size, font=SERIF, spc=-60)


def gone_up(sl, name, runs, **kw):
    title(sl, name, runs, y=-1.9, **kw)


def group(tl, shapes, kind, at, **kw):
    for s in shapes:
        tl.add(s, kind, at=at, **kw)


def slider(sl, tl, at, dur):
    """Before/after handle that travels with a left-to-right wipe on the filter photo."""
    x = FX + FW
    ln = polyline(sl, "sldLine", [(x, FY - 0.05), (x, FY + FH + 0.05)], WHITE, 2.5)
    knob = oval(sl, "sldKnob", x, FY + FH / 2, 0.42, 0.42, WHITE)
    shadow(knob, 0.1, 0.03, 30)
    arrows = text(sl, "sldArr", x - 0.21, FY + FH / 2 - 0.21, 0.42, 0.42, [("↔", TEAL, True)], 14, font=MONO,
                  align="c", anchor="m", wrap=False)
    for s_ in (ln, knob, arrows):
        tl.add(s_, "slide", at=at, dur=dur, dx=-FW)
        tl.add(s_, "out", at=at + dur + 250, dur=350)


def headline_words(lines, pt, x0=None, ys=(2.0, 3.25)):
    """Word boxes for the big serif headline. lines: [[(word, italic)], ...]."""
    fonts = {False: ImageFont.truetype(os.path.join(FONTS, "SourceSerif4Display-Light.ttf"), 1000),
             True: ImageFont.truetype(os.path.join(FONTS, "SourceSerif4Display-LightIt.ttf"), 1000)}
    space = fonts[False].getlength(" ") / 1000 * pt / 72
    out = []
    for words, y in zip(lines, ys):
        widths = [fonts[it].getlength(wd) / 1000 * pt / 72 for wd, it in words]
        total = sum(widths) + space * (len(words) - 1)
        x = x0 if x0 is not None else (W - total) / 2
        for (wd, it), wdt in zip(words, widths):
            out.append((wd, it, x, y, wdt))
            x += wdt + space
    return out


# ---------------------------------------------------------------- slides

def build(final, small):
    d = Deck()
    notes = load_notes()
    A = lambda n: os.path.join(ASSETS, n)
    MORPH = morph_xml(1000)
    HEAD = [[("Every", False), ("photo", False)], [("is", False), ("a", False), ("matrix.", True)]]
    head_left = headline_words(HEAD, 80, x0=0.8)
    head_mid = headline_words(HEAD, 80, ys=(2.35, 3.6))

    def headline(sl, words, tl=None, at=0, up=False):
        for k, (wd, it, x, y, wdt) in enumerate(words):
            col = TEAL if it else INK
            t = text(sl, f"w{k}", x, -2.2 if up else y, wdt + 0.6, 1.3, [(wd, col, False, it)], 80,
                     font=SERIF, anchor="m", wrap=False)
            if tl:
                tl.add(t, "mask", at=at + k * 110, dur=750, d=0.45)

    crop = (REG_C / PIX_W, REG_R / PIX_H, 1 - (REG_C + 4) / PIX_W, 1 - (REG_R + 4) / PIX_H)
    zc = (Z_X + 2 * Z_B, Z_Y + 2 * Z_B)           # centre of the zoom region
    zs = 4 * Z_B

    # 1 — title: headline rises line by line, a photo print drops onto the table
    sl, tl = d.new()
    tl.add(label(sl, "label0", "UNIT 5 · APPLICATIONS OF MATRICES"), "fade", at=0)
    headline(sl, head_left, tl, at=150)
    tl.add(text(sl, "sub", 0.85, 4.75, 5.8, 0.5, [("A short story about pixels, numbers and filters.", GREY_D)],
                18), "mask", at=800)
    tl.add(text(sl, "foot", 0.8, 6.75, 5, 0.3, [("ENGLISH · UNIT 5 · 2026", GREY)], 10, font=MONO, spc=300),
           "fade", at=1100)
    for k, (img, dx, dy, rot, at) in enumerate(STACK):
        f2, p2 = photo_print(sl, A(img), 9.85 + dx, 3.9 + dy, 4.2, 3.15, rot=rot, name=f"stackP{k}",
                             frame=f"stackF{k}")
        group(tl, [f2, p2], "fly", at, dur=850, dx=dx * 4, dy=-6, bounce=True)
    fr, ph = photo_print(sl, A("photo.png"), 9.85, 3.9, 4.2, 3.15, rot=-5)
    tl.add(fr, "fly", at=650, dur=900, dy=-6, bounce=True)
    tl.add(ph, "fade", at=1450, dur=1600)           # the print develops
    group(tl, viewfinder(sl, 7.35, 1.75, 5.0, 4.3), "pop", 1700, over=1.2)

    # 2 — look at this photo (print straightens, camera flash)
    sl, tl = d.new(MORPH)
    label(sl, "label0", "UNIT 5 · APPLICATIONS OF MATRICES", x=-7.5)
    headline(sl, head_left, up=True)
    text(sl, "sub", 0.85, 8.0, 5.8, 0.5, [("A short story about pixels, numbers and filters.", GREY_D)], 18)
    for k, (img, dx, dy, rot, at) in enumerate(STACK):
        photo_print(sl, A(img), (-4 if k == 0 else W + 4), 9.5, 4.2, 3.15, rot=rot * 3, name=f"stackP{k}",
                    frame=f"stackF{k}")
    photo_print(sl, A("photo.png"), W / 2, 4.35, 6.4, 4.8, pad=0.18)
    vf = viewfinder(sl, W / 2 - 3.6, 1.72, 7.2, 5.26)
    tl.add(label(sl, "label1", "01 — PIXELS"), "fade", at=0)
    tl.add(title(sl, "t2", "Look at this photo."), "mask", at=150)
    tl.add(text(sl, "img", 7.25, 1.2, 3.12, 0.3, [("IMG_0412.JPG", GREY_D)], 10, font=MONO, align="r", spc=200),
           "fade", at=700)
    group(tl, vf, "pulse", 250, s=1.06)          # viewfinder "focuses" on the photo

    # 3 — zoom in: Morph crops into the sun, pixels appear
    sl, tl = d.new(morph_xml(1400))
    photo_print(sl, A("photo.png"), *zc, zs, zs, pad=0.14, crop=crop)
    tl.add(pic(sl, "pix", A("pix.png"), Z_X, Z_Y, zs, zs, crop=crop), "fade", at=150, dur=700)
    viewfinder(sl, Z_X - 0.35, Z_Y - 0.35, zs + 0.7, zs + 0.7)
    label(sl, "label1", "01 — PIXELS")
    gone_up(sl, "t2", "Look at this photo.")
    tl.add(title(sl, "t3", [("Zoom in: ", INK), ("pixels.", TEAL, False, True)]), "mask", at=0)
    right3 = [("big1612", RX, 1.75, 6.2, 1.5, [("16 × 12", INK)], 96, SERIF),
              ("cap3a", RX + 0.05, 3.3, 6.2, 0.5, [("pixels in this little photo", GREY_D)], 20, SANS),
              ("cap3b", RX + 0.05, 4.25, 6.2, 0.4, [("A photo on your phone has about", GREY_D)], 20, SANS)]
    for k, (nm, x, y, w, h, runs, s, f) in enumerate(right3):
        tl.add(text(sl, nm, x, y, w, h, runs, s, font=f), "mask", at=750 + k * 180)
    counts = ["0", "3 000", "250 000", "4 000 000", "12 000 000"]
    for k, v in enumerate(counts):
        c = text(sl, f"cnt{k}", RX, 4.62, 6.2, 1.1, [(v, TEAL, False, True)], 60, font=SERIF)
        t0 = 1400 + k * 150
        tl.add(c, "fade", at=t0, dur=40)
        if k < len(counts) - 1:
            tl.add(c, "out", at=t0 + 150, dur=40)
        else:
            tl.add(c, "pulse", at=t0 + 100, s=1.08)
    tl.add(text(sl, "cap3c", RX + 0.05, 5.7, 6.2, 0.4, [("pixels.", GREY_D)], 20), "fade", at=2150)

    def region_numbers(sl, on_blocks=True, size=21):
        out = {}
        for i in range(4):
            for j in range(4):
                v = small[REG_R + i, REG_C + j]
                c = (INK if v > 150 else CREAM) if on_blocks else INK
                out[(i, j)] = text(sl, f"n{i}{j}", Z_X + j * Z_B, Z_Y + i * Z_B, Z_B, Z_B, [(str(v), c, on_blocks)],
                                   size, font=MONO, align="c", anchor="m", wrap=False)
        return out

    # 4 — every pixel is a number
    sl, tl = d.new(MORPH)
    photo_print(sl, A("photo.png"), *zc, zs, zs, pad=0.14, crop=crop)
    pic(sl, "pix", A("pix.png"), Z_X, Z_Y, zs, zs, crop=crop)
    label(sl, "label1", "01 — PIXELS")
    gone_up(sl, "t3", [("Zoom in: ", INK), ("pixels.", TEAL, False, True)])
    tl.add(title(sl, "t4", [("Every pixel is a ", INK), ("number.", TEAL, False, True)]), "mask", at=0)
    for (i, j), t in region_numbers(sl).items():
        tl.add(t, "pop", at=330 + i * 275 + j * 45, over=1.25)
    for nm, x, y, w, h, runs, s, f in right3:
        text(sl, nm, x + 7.5, y, w, h, runs, s, font=f)
    text(sl, "cnt4", RX + 7.5, 4.62, 6.2, 1.1, [("12 000 000", TEAL, False, True)], 60, font=SERIF)
    text(sl, "cap3c", RX + 7.55, 5.7, 6.2, 0.4, [("pixels.", GREY_D)], 20)
    # a scan line sweeps down; each row of numbers appears as it passes
    scan = polyline(sl, "scan", [(Z_X - 0.3, Z_Y + zs), (Z_X + zs + 0.3, Z_Y + zs)], TEAL, 2.5)
    tl.add(scan, "slide", at=200, dur=1100, dy=-zs)
    tl.add(scan, "out", at=1350, dur=300)
    bar = pic(sl, "bar", A("bar.png"), RX + 0.05, 2.3, 5.8, 0.55, line=INK, line_w=0.5)
    tl.add(bar, "stretchx", at=900, dur=700)
    scale = [("b0", RX + 0.05, [("0", INK, True), ("  dark", GREY_D)], "l"),
             ("b255", RX + 3.35, [("light  ", GREY_D), ("255", INK, True)], "r")]
    for k, (nm, x, runs, al) in enumerate(scale):
        tl.add(text(sl, nm, x, 3.0, 2.5, 0.4, runs, 18, font=MONO, align=al), "mask", at=1300 + k * 150)
    tl.add(text(sl, "cap4", RX + 0.05, 3.9, 5.8, 0.9, [("Every shade has its own number.", GREY_D)], 20),
           "mask", at=1650)

    # 5 — it's a matrix: pixels fall away one by one, the numbers stay on the print
    sl, tl = d.new(morph_xml(900))
    rect(sl, "frame", Z_X - 0.14, Z_Y - 0.14, zs + 0.28, zs + 0.28, WHITE)
    shadow(sl.shapes[-1])
    order = [(i, j) for i in range(4) for j in range(4)]
    rng = np.random.default_rng(3)
    rng.shuffle(order)
    for k, (i, j) in enumerate(order):
        b = pic(sl, f"blk{i}{j}", A(f"block{i}{j}.png"), Z_X + j * Z_B, Z_Y + i * Z_B, Z_B, Z_B)
        tl.add(b, "fallout", at=250 + k * 70, dur=650, d=1.2)
    region_numbers(sl, on_blocks=False)
    label(sl, "label1", "01 — PIXELS")
    gone_up(sl, "t4", [("Every pixel is a ", INK), ("number.", TEAL, False, True)])
    tl.add(title(sl, "t5", [("It's a ", INK), ("matrix!", TEAL, False, True)]), "mask", at=0)
    bl = bracket(sl, "brL", Z_X + 0.1, Z_Y + 0.1, zs - 0.2, "l", 0.2, INK, 2.5)
    br = bracket(sl, "brR", Z_X + zs - 0.1, Z_Y + 0.1, zs - 0.2, "r", 0.2, INK, 2.5)
    group(tl, [bl, br], "pop", 1500, over=1.1)
    pic(sl, "bar", A("bar.png"), RX + 7.55, 2.3, 5.8, 0.55, line=INK, line_w=0.5)
    for nm, x, runs, al in scale:
        text(sl, nm, x + 7.5, 3.0, 2.5, 0.4, runs, 18, font=MONO, align=al)
    text(sl, "cap4", RX + 7.55, 3.9, 5.8, 0.9, [("Every shade has its own number.", GREY_D)], 20)
    eq = text(sl, "eq5", RX, 2.0, 6.5, 1.3, [("photo", INK), (" = ", GREY), ("matrix", TEAL, False, True)], 66,
              font=SERIF)
    tl.add(eq, "mask", at=1700, d=0.4)
    tl.add(text(sl, "cap5", RX + 0.05, 3.5, 6, 0.5, [("rows and columns of numbers", GREY_D)], 20), "mask", at=1950)

    # 6 — colour photo = three matrices
    sl, tl = d.new(MORPH)
    cx, cy, cs = 1.25, 2.72, 3.9
    cards = []
    for k, (nm, col, lab_) in enumerate([("cardB", TEAL, "BLUE"), ("cardG", SAGE, "GREEN")]):
        off = (2 - k) * 0.42
        c = rect(sl, nm, cx + off, cy - off, cs, cs, WHITE, 100, col, 1.25)
        shadow(c, 0.18, 0.05, 16)
        t = text(sl, nm + "t", cx + off + 0.18, cy - off + 0.1, 2, 0.3, [(lab_, col)], 10, font=MONO, spc=300)
        tl.add(c, "fly", at=300, dur=650, dx=-off, dy=off)
        tl.add(t, "fade", at=800)
    fr = rect(sl, "frame", cx, cy, cs, cs, WHITE, 100, RED, 1.25)
    shadow(fr, 0.18, 0.05, 16)
    tl.add(text(sl, "cardRt", cx + 0.18, cy + 0.1, 2, 0.3, [("RED", RED)], 10, font=MONO, spc=300), "fade", at=800)
    b = 0.78
    for i in range(4):
        for j in range(4):
            text(sl, f"n{i}{j}", cx + 0.39 + j * b, cy + 0.5 + i * b, b, b, [(str(small[REG_R + i, REG_C + j]), INK)],
                 16, font=MONO, align="c", anchor="m", wrap=False)
    bracket(sl, "brL", cx + 0.3, cy + 0.5, 4 * b, "l", 0.16, INK, 2)
    bracket(sl, "brR", cx + 0.48 + 4 * b, cy + 0.5, 4 * b, "r", 0.16, INK, 2)
    label(sl, "label1", "01 — PIXELS")
    text(sl, "eq5", RX + 7.5, 2.0, 6.5, 1.3, [("photo", INK), (" = ", GREY), ("matrix", TEAL, False, True)], 66,
         font=SERIF)
    text(sl, "cap5", RX + 7.55, 3.5, 6, 0.5, [("rows and columns of numbers", GREY_D)], 20)
    gone_up(sl, "t5", [("It's a ", INK), ("matrix!", TEAL, False, True)])
    tl.add(title(sl, "t6", [("A colour photo = ", INK), ("three", TEAL, False, True), (" matrices.", INK)]),
           "mask", at=0)
    for k, (t, c) in enumerate([("Red", RED), ("Green", SAGE), ("Blue", TEAL)]):
        tl.add(oval(sl, f"dot{k}", 7.15, 2.76 + k * 0.85, 0.22, 0.22, c), "pop", at=950 + k * 170)
        tl.add(text(sl, f"rgb{k}", 7.5, 2.45 + k * 0.85, 5, 0.7, [(t, INK)], 34, font=SERIF),
               "mask", at=950 + k * 170)
    tl.add(text(sl, "cap6", 7.05, 5.15, 5.5, 0.9, [("Each colour has its own matrix of numbers.", GREY_D)], 18),
           "mask", at=1500)
    oval(sl, "iris", W / 2, H / 2, 0.02, 0.02, TEAL, 0)

    # 7 — chapter: a shutter-like circle closes over the screen
    sl, tl = d.new(morph_xml(800, adv=1900))
    oval(sl, "iris", W / 2, H / 2, 16, 16, TEAL)
    tl.add(text(sl, "chapTag", 0, 2.35, W, 0.4, [("PART 02", CREAM)], 13, font=MONO, align="c", spc=600),
           "fade", at=0)
    tl.add(text(sl, "chapHead", 0, 2.85, W, 1.5, [("Filters are ", CREAM), ("math.", CREAM, False, True)], 84,
                font=SERIF, align="c", anchor="m", spc=-80), "mask", at=120, dur=800, d=0.5)
    tl.add(text(sl, "chapSub", 0, 4.45, W, 0.5, [("brighter  ·  mirror  ·  blur", CREAM)], 18, font=MONO,
                align="c"), "fade", at=600)

    M = [["40", "90", "150"], ["70", "120", "180"], ["20", "60", "200"]]
    Mb = [[str(int(v) + 50) for v in row] for row in M]
    fcx, fcy = FX + FW / 2, FY + FH / 2

    # 8 — brighter = +50 (circle opens again)
    sl, tl = d.new(morph_xml(900))
    oval(sl, "iris", fcx, fcy, 0.02, 0.02, TEAL, 0)
    photo_print(sl, A("photo.png"), fcx, fcy, FW, FH)
    tl.add(pic(sl, "bright", A("bright.png"), FX, FY, FW, FH), "wipe", at=1700, dur=1100, dir="left")
    slider(sl, tl, 1700, 1100)
    tl.add(label(sl, "label2", "02 — FILTERS"), "fade", at=200)
    tl.add(title(sl, "t8", [("Brighter", INK), (" = +50", TEAL, False, True)]), "mask", at=250)
    m = matrix(sl, "M", M, MX + 0.3, 1.95, 0.95, 0.55, 22)
    group(tl, m.values(), "fly", 450, dur=650, dx=2.5, bounce=True)
    chip = rect(sl, "chip", MX + 1.12, 3.8, 1.4, 0.55, TEAL, 100, None, 1, 0.5)
    tl.add(chip, "spin", at=850)
    tl.add(text(sl, "chipT", MX + 1.12, 3.8, 1.4, 0.55, [("+ 50", CREAM, True)], 20, font=MONO, align="c",
                anchor="m"), "spin", at=850)
    mb = matrix(sl, "Mb", Mb, MX + 0.3, 4.6, 0.95, 0.55, 22, colors=TEAL, bold=True)
    group(tl, [mb["MbL"], mb["MbR"]], "pop", 1100)
    for i in range(3):
        group(tl, [mb[f"Mb{i + 1}{j + 1}"] for j in range(3)], "pop", 1200 + i * 150, over=1.3)
    tl.add(text(sl, "cap8", FX, 6.75, FW, 0.35, [("BIGGER NUMBERS → LIGHTER PIXELS", GREY_D)], 10, font=MONO,
                spc=200), "fade", at=2400)

    N = [["10", "50", "90"], ["20", "60", "100"], ["30", "70", "110"]]
    NW, NH, NX, NY = 1.1, 0.75, MX + 0.3, 2.2
    sw = FW / PIX_W

    def strips(sl, mirrored, x0=FX):
        out = []
        for c in range(PIX_W):
            pos = PIX_W - 1 - c if mirrored else c
            out.append(pic(sl, f"s{c:02d}", A(f"strip{c:02d}.png"), x0 + pos * sw, FY, sw, FH))
        return out

    def col_hl(sl, first_at, last_at, dx=0.0, dy_first=0.0, dy_last=0.0):
        a = rect(sl, "hlFirst", NX + first_at * NW + 0.08 + dx, NY - 0.12 + dy_first, NW - 0.16, 3 * NH + 0.24,
                 SAGE, 30, TEAL, 1.25, 0.2)
        b = rect(sl, "hlLast", NX + last_at * NW + 0.08 + dx, NY - 0.12 + dy_last, NW - 0.16, 3 * NH + 0.24,
                 CREAM, 100, TEAL, 1.25, 0.2)
        return a, b

    def pix_print(sl, mirrored=False, x=None, w=FW, h=FH):
        cx_ = fcx if x is None else x
        photo_print(sl, A("photo_mirror.png" if mirrored else "photo.png"), cx_, fcy, w, h)
        return pic(sl, "pixph", A("pix_mirror.png" if mirrored else "pix.png"), cx_ - w / 2, fcy - h / 2, w, h)

    # 9 — mirror: the photo turns into 16 pixel columns
    sl, tl = d.new(MORPH)
    photo_print(sl, A("photo.png"), fcx, fcy, FW, FH)
    tl.add(pic(sl, "pixph", A("pix.png"), FX, FY, FW, FH), "fade", at=150, dur=500)
    label(sl, "label2", "02 — FILTERS")
    gone_up(sl, "t8", [("Brighter", INK), (" = +50", TEAL, False, True)])
    tl.add(title(sl, "t9", [("Mirror", INK), (" = swap the columns.", TEAL, False, True)]), "mask", at=0)
    matrix(sl, "M", M, MX + 7.3, 1.95, 0.95, 0.55, 22)
    rect(sl, "chip", MX + 8.12, 3.8, 1.4, 0.55, TEAL, 100, None, 1, 0.5)
    text(sl, "chipT", MX + 8.12, 3.8, 1.4, 0.55, [("+ 50", CREAM, True)], 20, font=MONO, align="c", anchor="m")
    matrix(sl, "Mb", Mb, MX + 7.3, 4.6, 0.95, 0.55, 22, colors=TEAL, bold=True)
    text(sl, "cap8", FX, 8.2, FW, 0.35, [("BIGGER NUMBERS → LIGHTER PIXELS", GREY_D)], 10, font=MONO, spc=200)
    a, b = col_hl(sl, 0, 2)
    tl.add(a, "stretchy", at=900)
    tl.add(b, "stretchy", at=1050)
    group(tl, matrix(sl, "N", N, NX, NY, NW, NH, 28).values(), "pop", 450)

    # 10 — (auto) half-way through the flip: the print turns edge-on like a card,
    # the first and last columns swing over and under the matrix
    sl, tl = d.new(morph_xml(450, adv=0))
    pix_print(sl, w=0.06, h=FH * 1.05)
    label(sl, "label2", "02 — FILTERS")
    title(sl, "t9", [("Mirror", INK), (" = swap the columns.", TEAL, False, True)])
    lift = 2.2
    col_hl(sl, 1, 1, dy_first=-lift, dy_last=lift)
    mid = [[f"N{i + 1}{j + 1}" for j in range(3)] for i in range(3)]
    for i in range(3):
        for j, dy in [(0, -lift), (1, 0), (2, lift)]:
            text(sl, mid[i][j], NX + NW, NY + i * NH + dy, NW, NH, [(N[i][j], INK)], 28 if j == 1 else 24, font=MONO,
                 align="c", anchor="m", wrap=False)
    bracket(sl, "NL", NX - 0.1, NY, 3 * NH, "l", 0.18, INK, 2.0)
    bracket(sl, "NR", NX + 3 * NW + 0.1, NY, 3 * NH, "r", 0.18, INK, 2.0)

    # 11 — flip! the print turns back with the mirrored picture
    sl, tl = d.new(morph_xml(650))
    pix_print(sl, mirrored=True)
    label(sl, "label2", "02 — FILTERS")
    gone_up(sl, "t9", [("Mirror", INK), (" = swap the columns.", TEAL, False, True)])
    tl.add(title(sl, "t10", [("Flip!", INK)]), "drop", at=100, d=0.9)
    col_hl(sl, 2, 0)
    swapped = [[f"N{i + 1}{3 - j}" for j in range(3)] for i in range(3)]
    matrix(sl, "N", [row[::-1] for row in N], NX, NY, NW, NH, 28, names=swapped)
    tl.add(text(sl, "cap10", NX - 0.1, 4.75, 4.5, 0.5, [("first ↔ last", TEAL)], 22, font=MONO), "mask", at=400)

    K = [["90", "90", "90"], ["90", "0", "90"], ["90", "90", "90"]]
    KX, KY, KW, KH = 1.5, 2.15, 1.3, 1.05

    # 11 — blur = average of the neighbours
    sl, tl = d.new(MORPH)
    pix_print(sl, mirrored=True, x=fcx - 8)
    label(sl, "label2", "02 — FILTERS")
    gone_up(sl, "t10", [("Flip!", INK)])
    tl.add(title(sl, "t11", [("Blur", INK), (" = the average of neighbours.", TEAL, False, True)]), "mask", at=0)
    col_hl(sl, 2, 0, dx=7.3)
    matrix(sl, "N", [row[::-1] for row in N], NX + 7.3, NY, NW, NH, 28, names=swapped)
    text(sl, "cap10", NX + 7.2, 4.75, 4.5, 0.5, [("first ↔ last", TEAL)], 22, font=MONO)
    k = matrix(sl, "K", K, KX, KY, KW, KH, 40)
    group(tl, [v for nm_, v in k.items() if nm_ != "K22"], "fly", 250, dur=650, dx=-3, bounce=True)
    tl.add(k["K22"], "pop", at=500, over=1.3)
    win = rect(sl, "win", KX - 0.02, KY - 0.02, 3 * KW + 0.04, 3 * KH + 0.04, None, 100, TEAL, 3, 0.08)
    tl.add(win, "pop", at=900)
    tl.add(text(sl, "k1", MX - 0.2, 2.35, 5.2, 0.6, [("90 × 8 + 0 = ", GREY_D), ("720", INK, True)], 26, font=MONO),
           "mask", at=1200)
    tl.add(text(sl, "k2", MX - 0.2, 3.15, 5.2, 0.6, [("720 ÷ 9 = ", GREY_D), ("80", TEAL, True)], 26, font=MONO),
           "mask", at=1500)
    tl.add(k["K22"], "out", at=1900, dur=250)
    k80 = text(sl, "K22b", KX + KW, KY + KH, KW, KH, [("80", TEAL, True)], 40, font=MONO, align="c", anchor="m",
               wrap=False)
    tl.add(k80, "pop", at=2100, over=1.4)
    tl.add(text(sl, "cap11", MX - 0.2, 4.25, 5, 0.9, [("The black pixel becomes softer.", GREY_D)], 20),
           "mask", at=2400)

    # 12 — the window visits every pixel
    sl, tl = d.new(morph_xml(1100))
    photo_print(sl, A("photo.png"), fcx, fcy, FW, FH)
    tl.add(pic(sl, "blur", A("blur.png"), FX, FY, FW, FH), "wipe", at=700, dur=1600, dir="left")
    label(sl, "label2", "02 — FILTERS")
    gone_up(sl, "t11", [("Blur", INK), (" = the average of neighbours.", TEAL, False, True)])
    tl.add(title(sl, "t12", [("The computer does it for ", INK), ("every", TEAL, False, True), (" pixel.", INK)]),
           "mask", at=0)
    matrix(sl, "K", [["90", "90", "90"], ["90", None, "90"], ["90", "90", "90"]], MX + 0.3, 2.0, 0.8, 0.6, 22)
    text(sl, "K22b", MX + 0.3 + 0.8, 2.6, 0.8, 0.6, [("80", TEAL, True)], 22, font=MONO, align="c", anchor="m")
    text(sl, "k1", MX + 7, 2.35, 5.2, 0.6, [("90 × 8 + 0 = ", GREY_D), ("720", INK, True)], 26, font=MONO)
    text(sl, "k2", MX + 7, 3.15, 5.2, 0.6, [("720 ÷ 9 = ", GREY_D), ("80", TEAL, True)], 26, font=MONO)
    text(sl, "cap11", MX + 7, 4.25, 5, 0.9, [("The black pixel becomes softer.", GREY_D)], 20)
    win = rect(sl, "win", FX + FW - 0.35, FY + FH / 2 - 0.35, 0.7, 0.7, None, 100, CREAM, 2.5, 0.08)
    tl.add(win, "slide", at=700, dur=1600, dx=-FW)
    tl.add(win, "out", at=2450, dur=350)
    tl.add(text(sl, "cap12", MX + 0.2, 4.3, 5, 0.9, [("→ a soft, blurry photo", TEAL, False, True)], 30,
                font=SERIF), "mask", at=2200)

    if final:
        cw_, ch_, gap = 3.0, 3.0 * 9 / 16, 0.3
        gw, gh = 3 * cw_ + 2 * gap, 3 * ch_ + 2 * gap
        gx0, gy0 = (W - gw) / 2, (H - gh) / 2
        order = [1, 4, 7, 5, 13, 8, 6, 11, 9]
        zf = cw_ / W
        cells = []
        for idx, n_ in enumerate(order):
            i, j = divmod(idx, 3)
            cells.append((idx, n_, gx0 + j * (cw_ + gap) + cw_ / 2, gy0 + i * (ch_ + gap) + ch_ / 2))
        snap = lambda n_: os.path.join(SNAPS, f"s-{n_:02d}.png")
        zoom = lambda v, c, f: c + (v - c) / f
        ccx, ccy = cells[4][2], cells[4][3]

        def cell_pics(sl, f=1.0, alpha=100):
            out = []
            for idx, n_, cx_, cy_ in cells:
                w_, h_ = cw_ / f, ch_ / f
                x_, y_ = zoom(cx_, ccx, f), zoom(cy_, ccy, f)
                p = pic(sl, f"cell{n_}", snap(n_), x_ - w_ / 2, y_ - h_ / 2, w_, h_, alpha=alpha,
                        line=GREY, line_w=0.5)
                out.append(p)
            return out

        # 13 — frozen copy of slide 12 (centre cell), others waiting off-screen
        sl, tl = d.new(f'<p:transition xmlns:p="{NS_P}" spd="fast" advTm="0"/>')
        cell_pics(sl, f=zf)

        # 14 — zoom out to a contact sheet, the favourite frame gets circled
        sl, tl = d.new(morph_xml(2200))
        cell_pics(sl)
        for idx, n_, cx_, cy_ in cells:
            tl.add(text(sl, f"fn{idx}", cx_ - cw_ / 2, cy_ + ch_ / 2 + 0.04, 1.5, 0.25,
                        [(f"{idx + 1:02d}", GREY_D)], 9, font=MONO, spc=200), "fade", at=300 + idx * 60)
        pick = oval(sl, "pick", ccx, ccy, cw_ + 0.75, ch_ + 0.6, None, 100, TEAL, 2.5, rot=-4)
        tl.add(pick, "wheel", at=1100, dur=900)
        tl.add(text(sl, "t14", 0.8, 6.95, 11.8, 0.4, [("Brighter · mirror · blur — all made with matrices.", GREY_D)],
                    14, font=SANS, align="c"), "fade", at=1700)

        # 15 — final headline over the faded sheet
        sl, tl = d.new(morph_xml(900))
        cell_pics(sl, alpha=10)
        headline(sl, head_mid, tl, at=150)
        tl.add(text(sl, "sub", 0, 5.1, W, 0.5, [("Thank you for listening.", GREY_D)], 18, align="c"),
               "mask", at=900)

    return d.finish(notes)


def render(pptx, pdf):
    name = os.path.basename(pptx)
    script = f'''tell application "Microsoft PowerPoint"
open POSIX file "{pptx}"
delay 3
set p to presentation "{name}"
save p in POSIX file "{pdf}" as save as PDF
close p saving no
end tell'''
    subprocess.run(["perl", "-e", "alarm 180; exec @ARGV", "osascript", "-e", script], check=True)


def main():
    small = make_assets()
    os.makedirs(SNAPS, exist_ok=True)
    if "--skip-snaps" not in sys.argv:
        p1 = os.path.join(ROOT, "build", "pass1.pptx")
        build(False, small).save(p1)
        pdf = os.path.join(ROOT, "build", "pass1.pdf")
        render(p1, pdf)
        for f in os.listdir(SNAPS):
            os.remove(os.path.join(SNAPS, f))
        subprocess.run(["pdftoppm", "-png", "-r", "144", pdf, os.path.join(SNAPS, "s")], check=True)
    build(True, small).save(OUT)
    print("saved", OUT)


if __name__ == "__main__":
    main()
