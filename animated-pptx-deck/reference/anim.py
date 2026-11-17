"""Slide animation timeline (p:timing) written as raw OOXML.

python-pptx has no animation API, so effects are generated here. Every
effect lives in one group that starts by itself right after the slide
transition, so the presenter still clicks only once per slide.

Kinds (entrance unless noted):
  fade   plain fade
  rise   fades in while moving up by `d` inches
  fly    comes from an offset (dx, dy) inches, optional overshoot
  drop   falls from `d` inches above with a small bounce
  pop    scales from zero with overshoot
  grow   scales from factor `f` to full size
  spin   pop + turn from `r` degrees
  stretchx / stretchy  grows from its left / top edge to full width / height
  wipe   directional reveal, `dir` = left/right/up/down
  wheel  clockwise circular reveal around the shape centre
  pulse  emphasis: quick scale up and back
  slide  linear move from an offset (dx, dy), no fade — sliders, scan lines
  mask   text line rises out from behind an invisible edge
  out    exit: fades away
  fallout exit: drops down while fading
"""
import itertools

W, H = 13.333, 7.5


class Timeline:
    def __init__(self):
        self.effects = []   # (at_ms, spid, is_sp, kind, dur_ms, opts)

    def add(self, shape, kind, at=0, dur=None, **opts):
        is_sp = shape._element.tag.endswith("}sp")
        default = {"fade": 400, "rise": 600, "fly": 700, "drop": 750, "pop": 480,
                   "grow": 800, "spin": 600, "wipe": 500, "wheel": 800, "pulse": 200,
                   "stretchx": 550, "stretchy": 550, "out": 400,
                   "mask": 650, "fallout": 700, "slide": 1000}
        self.effects.append((int(at), shape.shape_id, is_sp, kind, int(dur or default[kind]), opts))
        return self

    def xml(self):
        if not self.effects:
            return None
        ids = itertools.count(5)
        pars, bld, used = [], [], {}
        for k, (at, spid, is_sp, kind, dur, o) in enumerate(sorted(self.effects, key=lambda e: e[0])):
            # every effect on the same shape needs its own build group, otherwise
            # PowerPoint silently drops all but the first one
            grp = used.get(spid, 0)
            used[spid] = grp + 1
            if is_sp:
                bld.append(f'<p:bldP spid="{spid}" grpId="{grp}" animBg="1"/>')
            eid = next(ids)
            cls, pid, beh = _behaviours(kind, spid, dur, o, ids)
            node = "afterEffect" if k == 0 else "withEffect"
            rev = ' autoRev="1"' if kind == "pulse" else ""
            pars.append(
                f'<p:par><p:cTn id="{eid}" presetID="{pid}" presetClass="{cls}" presetSubtype="0"{rev} '
                f'fill="hold" grpId="{grp}" nodeType="{node}"><p:stCondLst><p:cond delay="{at}"/>'
                f'</p:stCondLst><p:childTnLst>{beh}</p:childTnLst></p:cTn></p:par>')
        bld_xml = f"<p:bldLst>{''.join(bld)}</p:bldLst>" if bld else ""
        return (
            '<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
            '<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
            '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            '<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
            '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
            '<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst>{"".join(pars)}</p:childTnLst></p:cTn></p:par>'
            '</p:childTnLst></p:cTn></p:par>'
            '</p:childTnLst></p:cTn>'
            '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
            '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
            f'{bld_xml}</p:timing>')


def _behaviours(kind, spid, dur, o, ids):
    tgt = f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'

    def show():
        return (f'<p:set><p:cBhvr><p:cTn id="{next(ids)}" dur="1" fill="hold"><p:stCondLst>'
                f'<p:cond delay="0"/></p:stCondLst></p:cTn>{tgt}<p:attrNameLst>'
                '<p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
                '<p:to><p:strVal val="visible"/></p:to></p:set>')

    def filt(name, d):
        return (f'<p:animEffect transition="in" filter="{name}"><p:cBhvr>'
                f'<p:cTn id="{next(ids)}" dur="{int(d)}"/>{tgt}</p:cBhvr></p:animEffect>')

    def anim(attr, keys, d, ease=' decel="100000"'):
        tavs = "".join(f'<p:tav tm="{int(t * 100000)}"><p:val><p:strVal val="{v}"/></p:val></p:tav>'
                       for t, v in keys)
        return (f'<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
                f'<p:cTn id="{next(ids)}" dur="{int(d)}" fill="hold"{ease}/>{tgt}'
                f'<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst></p:cBhvr>'
                f'<p:tavLst>{tavs}</p:tavLst></p:anim>')

    def rot(frm, d):
        return (f'<p:anim calcmode="lin" valueType="num"><p:cBhvr>'
                f'<p:cTn id="{next(ids)}" dur="{int(d)}" fill="hold" decel="100000"/>{tgt}'
                '<p:attrNameLst><p:attrName>style.rotation</p:attrName></p:attrNameLst></p:cBhvr>'
                f'<p:tavLst><p:tav tm="0"><p:val><p:fltVal val="{frm}"/></p:val></p:tav>'
                '<p:tav tm="100000"><p:val><p:fltVal val="0"/></p:val></p:tav></p:tavLst></p:anim>')

    def scale_keys(attr, over=1.12):
        return [(0, "0"), (0.68, f"#{attr}*{over}"), (1, f"#{attr}")]

    def off(axis, inches):
        frac = inches / (W if axis == "x" else H)
        return f"#ppt_{axis}{frac:+.4f}"

    if kind == "fade":
        return "entr", 10, show() + filt("fade", dur)
    if kind == "rise":
        d = o.get("d", 0.35)
        return "entr", 42, (show() + filt("fade", dur * 0.7)
                            + anim("ppt_y", [(0, off("y", d)), (1, "#ppt_y")], dur))
    if kind == "fly":
        dx, dy = o.get("dx", 0), o.get("dy", 0)
        if o.get("bounce"):
            kx = [(0, off("x", dx)), (0.72, off("x", -dx * 0.04)), (1, "#ppt_x")]
            ky = [(0, off("y", dy)), (0.72, off("y", -dy * 0.04)), (1, "#ppt_y")]
            ease = ' decel="60000"'
        else:
            kx = [(0, off("x", dx)), (1, "#ppt_x")]
            ky = [(0, off("y", dy)), (1, "#ppt_y")]
            ease = ' decel="100000"'
        return "entr", 2, (show() + filt("fade", min(250, dur))
                           + anim("ppt_x", kx, dur, ease) + anim("ppt_y", ky, dur, ease))
    if kind == "drop":
        d = o.get("d", 1.2)
        keys = [(0, off("y", -d)), (0.62, off("y", d * 0.05)), (0.82, off("y", -d * 0.015)), (1, "#ppt_y")]
        return "entr", 2, show() + filt("fade", 220) + anim("ppt_y", keys, dur, ' accel="40000"')
    if kind in ("pop", "spin"):
        over = o.get("over", 1.12)
        beh = (show() + filt("fade", 200)
               + anim("ppt_w", scale_keys("ppt_w", over), dur, ' decel="50000"')
               + anim("ppt_h", scale_keys("ppt_h", over), dur, ' decel="50000"'))
        if kind == "spin":
            beh += rot(o.get("r", -180), dur)
        return "entr", 53, beh
    if kind == "grow":
        f = o.get("f", 0.4)
        return "entr", 53, (show() + filt("fade", dur * 0.6)
                            + anim("ppt_w", [(0, f"#ppt_w*{f}"), (1, "#ppt_w")], dur)
                            + anim("ppt_h", [(0, f"#ppt_h*{f}"), (1, "#ppt_h")], dur))
    if kind == "stretchx":
        return "entr", 22, (show() + filt("fade", 150)
                            + anim("ppt_w", [(0, "0"), (1, "#ppt_w")], dur)
                            + anim("ppt_x", [(0, "#ppt_x-#ppt_w/2"), (1, "#ppt_x")], dur))
    if kind == "stretchy":
        return "entr", 22, (show() + filt("fade", 150)
                            + anim("ppt_h", [(0, "0"), (1, "#ppt_h")], dur)
                            + anim("ppt_y", [(0, "#ppt_y-#ppt_h/2"), (1, "#ppt_y")], dur))
    if kind == "wipe":
        return "entr", 22, show() + filt(f"wipe({o.get('dir', 'right')})", dur)
    if kind == "wheel":
        return "entr", 21, show() + filt("wheel(1)", dur)
    if kind == "pulse":
        s = int(o.get("s", 1.2) * 100000)
        return "emph", 6, (f'<p:animScale><p:cBhvr><p:cTn id="{next(ids)}" dur="{dur}" '
                           f'fill="hold" decel="100000"/>{tgt}</p:cBhvr><p:by x="{s}" y="{s}"/></p:animScale>')
    if kind == "slide":
        dx, dy = o.get("dx", 0), o.get("dy", 0)
        return "entr", 2, (show() + anim("ppt_x", [(0, off("x", dx)), (1, "#ppt_x")], dur, "")
                           + anim("ppt_y", [(0, off("y", dy)), (1, "#ppt_y")], dur, ""))
    if kind == "mask":
        # line of text rising out from behind an invisible edge
        d = o.get("d", 0.25)
        return "entr", 22, (show() + filt("wipe(up)", dur * 0.85)
                            + anim("ppt_y", [(0, off("y", d)), (1, "#ppt_y")], dur))
    if kind == "fallout":
        d = o.get("d", 1.5)
        hide = (f'<p:set><p:cBhvr><p:cTn id="{next(ids)}" dur="1" fill="hold"><p:stCondLst>'
                f'<p:cond delay="{dur - 1}"/></p:stCondLst></p:cTn>{tgt}<p:attrNameLst>'
                '<p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
                '<p:to><p:strVal val="hidden"/></p:to></p:set>')
        return "exit", 2, (f'<p:animEffect transition="out" filter="fade"><p:cBhvr>'
                           f'<p:cTn id="{next(ids)}" dur="{dur}"/>{tgt}</p:cBhvr></p:animEffect>'
                           + anim("ppt_y", [(0, "#ppt_y"), (1, off("y", d))], dur, ' accel="100000"')
                           + hide)
    if kind == "out":
        hide = (f'<p:set><p:cBhvr><p:cTn id="{next(ids)}" dur="1" fill="hold"><p:stCondLst>'
                f'<p:cond delay="{dur - 1}"/></p:stCondLst></p:cTn>{tgt}<p:attrNameLst>'
                '<p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
                '<p:to><p:strVal val="hidden"/></p:to></p:set>')
        return "exit", 10, (f'<p:animEffect transition="out" filter="fade"><p:cBhvr>'
                            f'<p:cTn id="{next(ids)}" dur="{dur}"/>{tgt}</p:cBhvr></p:animEffect>' + hide)
    raise ValueError(kind)
