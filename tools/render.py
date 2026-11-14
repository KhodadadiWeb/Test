"""Rough preview renderer for the generated pptx (no LibreOffice available).

Approximates geometry, fills, rounded rects, pictures and RTL text so layout
problems (overflow, overlap, spacing) can be spotted visually.
"""
import io
import os
import sys
import arabic_reshaper
from bidi.algorithm import get_display
from PIL import Image, ImageDraw, ImageFont
from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
DPI = 110
EMU = 914400

FONTS = {
    (False, False): '/tmp/vazirmatn-master/fonts/ttf/Vazirmatn-Regular.ttf',
    (True, False): '/tmp/vazirmatn-master/fonts/ttf/Vazirmatn-Bold.ttf',
    (False, True): '/tmp/vazirmatn-master/fonts/ttf/Vazirmatn-Medium.ttf',
    (True, True): '/tmp/vazirmatn-master/fonts/ttf/Vazirmatn-Bold.ttf',
}
_fc = {}


def font(size_pt, bold):
    px = max(6, int(round(size_pt * DPI / 72)))
    key = (px, bold)
    if key not in _fc:
        _fc[key] = ImageFont.truetype(FONTS[(bold, False)], px)
    return _fc[key]


def px(emu):
    return emu / EMU * DPI


def shape_color(el):
    """find srgbClr of a fill element"""
    if el is None:
        return None, None
    s = el.find('.//{%s}srgbClr' % A)
    if s is not None:
        a = 255
        al = s.find('{%s}alpha' % A)
        if al is not None:
            a = int(int(al.get('val')) / 100000 * 255)
        return s.get('val'), a
    return None, None


def fill_of(spPr):
    for tag in ('noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'):
        e = spPr.find(qn('a:' + tag))
        if e is None:
            continue
        if tag == 'noFill':
            return None
        if tag == 'solidFill':
            c, a = shape_color(e)
            return ('solid', c, a)
        if tag == 'gradFill':
            stops = []
            for gs in e.findall('.//' + qn('a:gs')):
                c, a = shape_color(gs)
                stops.append((int(gs.get('pos')) / 100000, c, a))
            lin = e.find(qn('a:lin'))
            ang = int(lin.get('ang')) / 60000 if lin is not None else 0
            return ('grad', stops, ang)
    return None


def line_of(spPr):
    ln = spPr.find(qn('a:ln'))
    if ln is None:
        return None
    if ln.find(qn('a:noFill')) is not None:
        return None
    sf = ln.find(qn('a:solidFill'))
    if sf is None:
        return None
    c, a = shape_color(sf)
    w = ln.get('w')
    return (c, max(1, int(round(int(w or 12700) / 12700 * DPI / 72))))


def rgb(h, a=255):
    if h is None:
        return None
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), min(255, a))


def prepare(text):
    try:
        return get_display(arabic_reshaper.reshape(text))
    except Exception:
        return text


def wrap(draw, text, fnt, maxw):
    words = text.split(' ')
    lines, cur = [], ''
    for w in words:
        test = (cur + ' ' + w).strip()
        if draw.textlength(prepare(test), font=fnt) <= maxw or not cur:
            cur = test
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def draw_text(draw, tf, box, base):
    x, y, w, h = box
    bodyPr = tf.find(qn('a:bodyPr'))
    anchor = 't'
    if bodyPr is not None:
        anchor = bodyPr.get('anchor', 't')
    # measure
    blocks = []
    for p in tf.findall(qn('a:p')):
        pPr = p.find(qn('a:pPr'))
        algn = pPr.get('algn', 'l') if pPr is not None else 'l'
        marL = int(pPr.get('marL', 0)) if pPr is not None else 0
        rtl = pPr.get('rtl', '0') == '1' if pPr is not None else False
        runs = []
        for r in p.findall(qn('a:r')):
            rPr = r.find(qn('a:rPr'))
            size = int(rPr.get('sz', '1800')) / 100 if rPr is not None else 18
            bold = rPr is not None and rPr.get('b') == '1'
            col = INK_DEFAULT
            if rPr is not None:
                c, a = shape_color(rPr.find(qn('a:solidFill')))
                if c:
                    col = rgb(c, a)
            t = r.find(qn('a:t')).text or ''
            runs.append((t, size, bold, col))
        if not runs:
            runs = [('', 18, False, INK_DEFAULT)]
        spc_b = 0
        if pPr is not None:
            sb = pPr.find(qn('a:spcBef'))
            if sb is not None:
                pts = sb.findall('.//' + qn('a:spcPts'))
                if pts:
                    spc_b = int(pts[0].get('val')) / 100
            sa = pPr.find(qn('a:spcAft'))
            if sa is not None:
                pts = sa.findall('.//' + qn('a:spcPts'))
                if pts:
                    spc_b_sa = int(pts[0].get('val')) / 100
            else:
                spc_b_sa = 0
        else:
            spc_b_sa = 0
        lnspc = 1.0
        if pPr is not None:
            ls = pPr.find(qn('a:lnSpc'))
            if ls is not None:
                pct = ls.find(qn('a:spcPct'))
                if pct is not None:
                    lnspc = int(pct.get('val')) / 100000
        if rtl:
            runs = runs[::-1]
        size = max(r[1] for r in runs)
        blocks.append(dict(runs=runs, size=size, algn=algn, rtl=rtl,
                           spc_b=spc_b, spc_a=spc_b_sa, lnspc=lnspc,
                           marL=px(marL)))
    total = 0
    for b in blocks:
        f = font(b['size'], any(r[2] for r in b['runs']))
        text = ''.join(r[0] for r in b['runs'])
        lines = wrap(draw, text, f, max(10, w - b['marL']))
        b['lines'] = lines
        b['font'] = f
        b['h'] = len(lines) * b['size'] * b['lnspc'] * DPI / 72 * 1.02
        total += b['h'] + (b['spc_b'] + b['spc_a']) * DPI / 72
    ty = y
    if anchor == 'ctr':
        ty = y + (h - total) / 2
    elif anchor == 'b':
        ty = y + h - total
    for b in blocks:
        ty += b['spc_b'] * DPI / 72
        for ln in b['lines']:
            lw = draw.textlength(prepare(ln), font=b['font'])
            if b['algn'] == 'ctr':
                tx = x + (w - lw) / 2
            elif b['algn'] == 'r':
                tx = x + w - lw - b['marL']
            else:
                tx = x + b['marL']
            draw.text((tx, ty), prepare(ln), font=b['font'], fill=b['runs'][0][3])
            ty += b['size'] * b['lnspc'] * DPI / 72 * 1.02
        ty += b['spc_a'] * DPI / 72
    return total


INK_DEFAULT = (0, 0, 0, 255)


def overlay(img, box, color, radius=0):
    x0, y0, x1, y1 = [int(v) for v in box]
    if x1 <= x0 or y1 <= y0:
        return
    lay = Image.new('RGBA', (x1 - x0, y1 - y0), (0, 0, 0, 0))
    d2 = ImageDraw.Draw(lay)
    if radius:
        d2.rounded_rectangle((0, 0, x1 - x0 - 1, y1 - y0 - 1), radius=radius, fill=color)
    else:
        d2.rectangle((0, 0, x1 - x0, y1 - y0), fill=color)
    img.alpha_composite(lay, (x0, y0))


def rounded(draw, box, r, fill, outline, width=1):
    x0, y0, x1, y1 = box
    r = max(0, min(r, (x1 - x0) / 2, (y1 - y0) / 2))
    draw.rounded_rectangle((x0, y0, x1, y1), radius=r, fill=fill,
                           outline=outline, width=width)


def grad_fill(size, stops, ang):
    """Approximate a linear gradient (angle in pptx degrees, clockwise from
    +x axis) by sampling bands across the diagonal."""
    import math
    import numpy as np
    w, h = size
    a = math.radians(ang)
    dx, dy = math.cos(a), math.sin(a)
    xs = np.arange(w)[None, :]
    ys = np.arange(h)[:, None]
    cx, cy = w / 2, h / 2
    proj = (xs - cx) * dx + (ys - cy) * dy
    rng = abs(w * dx) + abs(h * dy)
    t = (proj + rng / 2) / max(rng, 1e-6)
    t = np.clip(t, 0, 1)
    pos = np.array([s[0] for s in stops], dtype=float)
    cols = np.array([rgb(s[1], s[2] if s[2] is not None else 255)
                     for s in stops], dtype=float)
    out = np.zeros((h, w, 4), dtype=np.uint8)
    for ch in range(4):
        out[:, :, ch] = np.interp(t, pos, cols[:, ch]).astype(np.uint8)
    return Image.fromarray(out, 'RGBA')


def render(path, out_dir, only=None):
    prs = Presentation(path)
    W = int(px(prs.slide_width))
    H = int(px(prs.slide_height))
    os.makedirs(out_dir, exist_ok=True)
    outs = []
    for idx, slide in enumerate(prs.slides, 1):
        if only and idx not in only:
            continue
        img = Image.new('RGBA', (W, H), (255, 255, 255, 255))
        draw = ImageDraw.Draw(img)
        # background
        sld = slide._element
        bg = sld.find(qn('p:cSld') + '/' + qn('p:bg'))
        if bg is not None:
            sf = bg.find('.//' + qn('a:solidFill'))
            gf = bg.find('.//' + qn('a:gradFill'))
            if gf is not None:
                stops = []
                for gs in gf.findall('.//' + qn('a:gs')):
                    c, a = shape_color(gs)
                    stops.append((int(gs.get('pos')) / 100000, c, a))
                lin = gf.find(qn('a:lin'))
                ang = int(lin.get('ang')) / 60000 if lin is not None else 0
                g = grad_fill((W, H), stops, ang)
                lay = Image.new('RGBA', (W, H), (0, 0, 0, 0))
                lay.paste(g, (0, 0))
                img.alpha_composite(lay)
            elif sf is not None:
                c, a = shape_color(sf)
                img.paste(Image.new('RGBA', (W, H), rgb(c, a)), (0, 0))
        for sh in slide.shapes:
            el = sh._element
            if el.tag == qn('p:pic'):
                blip = el.find('.//' + qn('a:blip'))
                rid = blip.get('{%s}embed' % ns_r())
                try:
                    blob = slide.part.related_part(rid).blob
                    pim = Image.open(io.BytesIO(blob)).convert('RGBA')
                except Exception:
                    continue
                cl, ct, cr, cb = (sh.crop_left or 0, sh.crop_top or 0,
                                  sh.crop_right or 0, sh.crop_bottom or 0)
                iw, ih = pim.size
                box = (int(cl * iw), int(ct * ih), int(iw * (1 - cr)), int(ih * (1 - cb)))
                pim = pim.crop(box)
                x, y = px(sh.left), px(sh.top)
                w, h = px(sh.width), px(sh.height)
                pim = pim.resize((max(1, int(w)), max(1, int(h))), Image.LANCZOS)
                spPr = el.find(qn('p:spPr'))
                geom = spPr.find(qn('a:prstGeom'))
                rad = 0
                if geom is not None and geom.get('prst') in ('roundRect', 'ellipse'):
                    if geom.get('prst') == 'ellipse':
                        m = Image.new('L', pim.size, 0)
                        ImageDraw.Draw(m).ellipse((0, 0, pim.size[0], pim.size[1]), fill=255)
                        img.paste(pim, (int(x), int(y)), m)
                        continue
                    gd = geom.find('.//' + qn('a:gd'))
                    rad = float(gd.get('fmla').split()[1]) / 100000 * min(pim.size)
                m = Image.new('L', pim.size, 0)
                ImageDraw.Draw(m).rectangle((0, 0, pim.size[0], pim.size[1]), fill=255)
                img.paste(pim, (int(x), int(y)), m)
                ln = line_of(spPr)
                if ln:
                    draw.rounded_rectangle((x, y, x + w, y + h), radius=rad,
                                           outline=rgb(ln[0]), width=ln[1])
                continue
            if el.tag == qn('p:graphicFrame'):
                tbl_el = el.find('.//' + qn('a:tbl'))
                if tbl_el is not None:
                    ox, oy = px(sh.left), px(sh.top)
                    yy = oy
                    rowsh = []
                    for r in tbl_el.findall(qn('a:tr')):
                        rowsh.append(px(int(r.get('h'))))
                    # recompute heights to fit text
                    newh = []
                    ci_w = [px(int(c.get('w'))) for c in tbl_el.find(qn('a:tblGrid'))]
                    for ri, r in enumerate(tbl_el.findall(qn('a:tr'))):
                        need = rowsh[ri]
                        for ci, tc in enumerate(r.findall(qn('a:tc'))):
                            tcb = tc.find(qn('a:txBody'))
                            if tcb is None:
                                continue
                            p_ = tcb.find(qn('a:p'))
                            rp = p_.find(qn('a:r'))
                            sz = int(rp.find(qn('a:rPr')).get('sz', '1700')) / 100 if rp is not None else 17
                            txt = p_.find(qn('a:r')).find(qn('a:t')).text or ''
                            f = font(sz, ri == 0)
                            lines = wrap(draw, txt, f, ci_w[ci] - 2 * px(0.12 * EMU))
                            need = max(need, len(lines) * sz * 1.25 * DPI / 72 + px(0.2 * EMU))
                        newh.append(need)
                    for ri, r in enumerate(tbl_el.findall(qn('a:tr'))):
                        xx = ox
                        for ci, tc in enumerate(r.findall(qn('a:tc'))):
                            tcPr = tc.find(qn('a:tcPr'))
                            col = (255, 255, 255, 255)
                            if tcPr is not None:
                                c, a = shape_color(tcPr.find(qn('a:solidFill')))
                                if c:
                                    col = rgb(c, a)
                            draw.rectangle((xx, yy, xx + ci_w[ci], yy + newh[ri]), fill=col,
                                           outline=rgb('D8E3F3'), width=1)
                            tb = tc.find(qn('a:txBody'))
                            if tb is not None:
                                draw_text(draw, tb, (xx + px(0.1 * EMU), yy + px(0.03 * EMU),
                                                     ci_w[ci] - 2 * px(0.1 * EMU),
                                                     newh[ri]), None)
                            xx += ci_w[ci]
                        yy += newh[ri]
                    continue
            if el.tag not in (qn('p:sp'),):
                continue
            spPr = el.find(qn('p:spPr'))
            if spPr is None:
                continue
            xfrm = spPr.find(qn('a:xfrm'))
            if xfrm is None:
                continue
            off = xfrm.find(qn('a:off'))
            ext = xfrm.find(qn('a:ext'))
            x, y = px(int(off.get('x'))), px(int(off.get('y')))
            w, h = px(int(ext.get('cx'))), px(int(ext.get('cy')))
            if w <= 0 or h <= 0:
                continue
            geom = spPr.find(qn('a:prstGeom'))
            prst = geom.get('prst') if geom is not None else 'rect'
            f = fill_of(spPr)
            ln = line_of(spPr)
            oc = rgb(ln[0]) if ln else None
            ow = (ln[1] if ln else 0) or 1
            solid = f[0] == 'solid' if f else False
            fc = rgb(f[1], f[2]) if solid else None
            rad = 0
            if prst == 'roundRect':
                gd = geom.find('.//' + qn('a:gd'))
                adj = float(gd.get('fmla').split()[1]) / 100000 if gd is not None else 0.05
                rad = adj * min(w, h)
            if f and f[0] == 'grad' and prst not in ('roundRect', 'ellipse'):
                g = grad_fill((max(1, int(w)), max(1, int(h))), f[1], f[2])
                lay = Image.new('RGBA', img.size, (0, 0, 0, 0))
                lay.paste(g, (int(x), int(y)))
                img.alpha_composite(lay)
            if prst == 'roundRect':
                if solid:
                    overlay(img, (x, y, x + w, y + h), fc, rad) if fc[3] < 255 else \
                        draw.rounded_rectangle((x, y, x + w, y + h), radius=rad, fill=fc)
                if ln:
                    draw.rounded_rectangle((x, y, x + w, y + h), radius=rad,
                                           outline=oc, width=ow)
            elif prst == 'ellipse':
                if solid:
                    draw.ellipse((x, y, x + w, y + h), fill=fc,
                                 outline=oc if ln else None, width=ow)
                elif ln:
                    draw.ellipse((x, y, x + w, y + h), outline=oc, width=ow)
            elif prst == 'leftArrow' or prst == 'rightArrow':
                if solid:
                    if prst == 'rightArrow':
                        pts = [(x, y + h * .25), (x + w * .62, y + h * .25),
                               (x + w * .62, y), (x + w, y + h / 2),
                               (x + w * .62, y + h), (x + w * .62, y + h * .75),
                               (x, y + h * .75)]
                    else:
                        pts = [(x + w, y + h * .25), (x + w * .38, y + h * .25),
                               (x + w * .38, y), (x, y + h / 2),
                               (x + w * .38, y + h), (x + w * .38, y + h * .75),
                               (x + w, y + h * .75)]
                    draw.polygon(pts, fill=fc)
            else:
                if solid:
                    draw.rectangle((x, y, x + w, y + h), fill=fc)
                if ln:
                    draw.rectangle((x, y, x + w, y + h), outline=oc, width=ow)
            tx = el.find(qn('p:txBody'))
            if tx is not None and el.find('.//' + qn('a:t')) is not None:
                bl = px(0)
                draw_text(draw, tx, (x + bl, y + bl, w - 2 * bl, h - 2 * bl), None)
        out = os.path.join(out_dir, 'slide%02d.png' % idx)
        img.convert('RGB').save(out, quality=90)
        outs.append(out)
    return outs


def ns_r():
    return 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


if __name__ == '__main__':
    src = sys.argv[1]
    out = sys.argv[2]
    only = set(int(a) for a in sys.argv[3:]) if len(sys.argv) > 3 else None
    files = render(src, out, only)
    print('\n'.join(files))
