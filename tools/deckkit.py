"""Small helper toolkit on top of python-pptx used to build the deck.

Design goals: light/clean academic look, consistent grid, RTL Persian text,
soft shadows, uniform fade transitions + fade-in animations.
"""
import copy
from lxml import etree
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

# ---------------------------------------------------------------- palette ---
INK      = RGBColor(0x14, 0x21, 0x3D)   # main text
NAVY     = RGBColor(0x18, 0x2C, 0x55)   # headings
BLUE     = RGBColor(0x21, 0x59, 0xA8)   # primary accent
BLUE_2   = RGBColor(0x3E, 0x7B, 0xD6)
BLUE_PALE= RGBColor(0xE8, 0xF0, 0xFC)
TEAL     = RGBColor(0x10, 0x8C, 0x92)
GOLD     = RGBColor(0xD9, 0x9A, 0x22)
RED      = RGBColor(0xB0, 0x48, 0x45)
MUTED    = RGBColor(0x5A, 0x6A, 0x85)
BORDER   = RGBColor(0xD8, 0xE3, 0xF3)
CARD     = RGBColor(0xF5, 0xF8, 0xFE)
CARD_2   = RGBColor(0xED, 0xF3, 0xFD)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
DARK     = RGBColor(0x0E, 0x1A, 0x2F)   # panels that hold dark screenshots

FONT = "Tahoma"

L, R = "left", "right"
SIDES = {'L': 'a:lnL', 'R': 'a:lnR', 'T': 'a:lnT', 'B': 'a:lnB'}
TCPR_ORDER = ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB', 'a:lnTlToBr', 'a:lnBlToTr',
              'a:cell3D', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill',
              'a:pattFill', 'a:grpFill', 'a:headers', 'a:extLst']

FA_DIGITS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")


def fa(s):
    return str(s).translate(FA_DIGITS)


# ------------------------------------------------------------------ utils ---
def _sub(parent, tag, **attrs):
    el = etree.SubElement(parent, qn(tag))
    for k, v in attrs.items():
        el.set(k, str(v))
    return el


def _insert_ordered(parent, el, order):
    tag = el.tag
    names = [qn(t) for t in order]
    idx = names.index(tag)
    for child in parent:
        if child.tag in names and names.index(child.tag) > idx:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


def set_grad(shape, stops, angle=0.0, angle_deg=None):
    """stops: list of (pos 0..1, RGBColor, alpha 0..100 or None)"""
    spPr = shape._element.spPr
    for tag in ('a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill',
                'a:pattFill', 'a:grpFill'):
        for e in spPr.findall(qn(tag)):
            spPr.remove(e)
    gs = "".join(
        '<a:gs pos="%d"><a:srgbClr val="%02X%02X%02X">%s</a:srgbClr></a:gs>' % (
            int(p * 100000), c[0], c[1], c[2],
            '<a:alpha val="%d"/>' % int(a * 1000) if a is not None else '')
        for p, c, a in stops)
    xml = ('<a:gradFill xmlns:a="%s" rotWithShape="1"><a:gsLst>%s</a:gsLst>'
           '<a:lin ang="%d" scaled="1"/></a:gradFill>'
           % (A, gs, int((angle_deg if angle_deg is not None else angle) * 60000)))
    _insert_ordered(spPr, parse_xml(xml),
                    ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill',
                     'a:pattFill', 'a:grpFill', 'a:ln', 'a:effectLst',
                     'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst'])


def set_line(shape, color=None, width_pt=1.0, dash=None):
    ln = shape.line
    if color is None:
        ln.fill.background()
    else:
        ln.color.rgb = color
        ln.width = Pt(width_pt)
        if dash:
            ln.dash_style = dash


def set_shadow(shape, blur=14, dist=3, direction=90, alpha=14, color=NAVY):
    spPr = shape._element.spPr
    for e in spPr.findall(qn('a:effectLst')):
        spPr.remove(e)
    xml = ('<a:effectLst xmlns:a="%s"><a:outerShdw blurRad="%d" dist="%d" '
           'dir="%d" rotWithShape="0"><a:srgbClr val="%02X%02X%02X">'
           '<a:alpha val="%d"/></a:srgbClr></a:outerShdw></a:effectLst>'
           % (A, int(blur * 12700), int(dist * 12700), int(direction * 60000),
              color[0], color[1], color[2], int(alpha * 1000)))
    _insert_ordered(spPr, parse_xml(xml),
                    ['a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d',
                     'a:sp3d', 'a:extLst'])


def no_shadow(shape):
    spPr = shape._element.spPr
    for e in spPr.findall(qn('a:effectLst')):
        spPr.remove(e)
    spPr.append(parse_xml('<a:effectLst xmlns:a="%s"/>' % A))


def style_run(run, size=20, bold=False, color=INK, font=FONT, italic=False, spc=None):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.name = font
    f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rPr.set('lang', 'fa-IR')
    rPr.set('dirty', '0')
    if spc is not None:
        rPr.set('spc', str(int(spc * 100)))
    for e in rPr.findall(qn('a:cs')):
        rPr.remove(e)
    cs = etree.SubElement(rPr, qn('a:cs'))
    cs.set('typeface', font)


def para(p, align='right', rtl=True, space_before=0, space_after=0,
         line_spacing=1.35, indent=0.0):
    pPr = p._p.get_or_add_pPr()
    pPr.set('rtl', '1' if rtl else '0')
    pPr.set('algn', {'right': 'r', 'left': 'l', 'center': 'ctr',
                     'justify': 'just'}[align])
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    p.line_spacing = line_spacing
    if indent:
        pPr.set('marR', str(int(indent * 914400)))
        pPr.set('indent', '0')


def textbox(slide, x, y, w, h, anchor='top'):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {'top': MSO_ANCHOR.TOP, 'middle': MSO_ANCHOR.MIDDLE,
                          'bottom': MSO_ANCHOR.BOTTOM}[anchor]
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    for t in ('a:normAutofit', 'a:spAutoFit', 'a:noAutofit'):
        for e in bodyPr.findall(qn(t)):
            bodyPr.remove(e)
    etree.SubElement(bodyPr, qn('a:noAutofit'))
    return tb


def add_text(slide, x, y, w, h, lines, anchor='top'):
    """lines: list of dicts: t, size, bold, color, align, space_after,
    line_spacing, rtl, runs=[(text, opts)]"""
    tb = textbox(slide, x, y, w, h, anchor)
    tf = tb.text_frame
    first = True
    for spec in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        para(p, align=spec.get('align', 'right'), rtl=spec.get('rtl', True),
             space_before=spec.get('space_before', 0),
             space_after=spec.get('space_after', 0),
             line_spacing=spec.get('line_spacing', 1.35))
        runs = spec.get('runs') or [(spec.get('t', ''), {})]
        for text, o in runs:
            r = p.add_run()
            r.text = text
            style_run(r, size=o.get('size', spec.get('size', 20)),
                      bold=o.get('bold', spec.get('bold', False)),
                      color=o.get('color', spec.get('color', INK)),
                      italic=o.get('italic', False),
                      spc=o.get('spc', spec.get('spc')))
    return tb


def shape(slide, kind, x, y, w, h, fill=None, line=None, line_w=1.0,
          radius=None, shadow=False, shadow_alpha=14, grad=None, angle=0,
          name=None):
    s = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius is not None:
        try:
            s.adjustments[0] = radius
        except Exception:
            pass
    if grad:
        set_grad(s, [(p, c, None) for p, c in grad], angle)
    elif fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        set_line(s, None)
    else:
        set_line(s, line, line_w)
    if shadow:
        set_shadow(s, alpha=shadow_alpha)
    else:
        no_shadow(s)
    if name:
        s.name = name
    # clear default text
    s.text_frame.word_wrap = True
    return s


def card(slide, x, y, w, h, fill=WHITE, radius=0.055, line=BORDER,
         shadow=True, grad=None, angle=0, shadow_alpha=13):
    return shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=fill,
                 line=line, line_w=1.0, radius=radius, shadow=shadow,
                 grad=grad, angle=angle, shadow_alpha=shadow_alpha)


def round_pic(pic, adj=0.05):
    spPr = pic._element.spPr
    for g in spPr.findall(qn('a:prstGeom')):
        spPr.remove(g)
    geom = parse_xml('<a:prstGeom xmlns:a="%s" prst="roundRect">'
                     '<a:avLst><a:gd name="adj" fmla="val %d"/></a:avLst>'
                     '</a:prstGeom>' % (A, int(adj * 100000)))
    xfrm = spPr.find(qn('a:xfrm'))
    if xfrm is not None:
        xfrm.addnext(geom)
    else:
        spPr.insert(0, geom)


def add_pic(slide, path, x, y, w, h, radius=0.04, border=BORDER, shadow=True,
            crop=None, bloom=False):
    pic = slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius:
        round_pic(pic, radius)
    if border:
        set_line(pic, border, 1.0)
    else:
        set_line(pic, None)
    if shadow:
        set_shadow(pic, blur=16, dist=3, alpha=16)
    else:
        no_shadow(pic)
    if crop:
        l, t, r, b = crop
        pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = l, t, r, b
    return pic


def pic_size(path):
    from PIL import Image
    with Image.open(path) as im:
        return im.size


def fit(path, x, y, w, h, mode='contain'):
    pw, ph = pic_size(path)
    ar = pw / ph
    if mode == 'contain':
        if w / h > ar:
            nh = h
            nw = h * ar
        else:
            nw = w
            nh = w / ar
        return x + (w - nw) / 2, y + (h - nh) / 2, nw, nh
    else:  # cover -> returns crop fractions
        if w / h > ar:
            nw = w
            nh = w / ar
            over = (nh - h) / nh
            return x, y, w, h, (0, over / 2, 0, over / 2)
        nh = h
        nw = h * ar
        over = (nw - w) / nw
        return x, y, w, h, (over / 2, 0, over / 2, 0)


# ----------------------------------------------------------------- tables ---
def _cell_border(cell, side, color, width_pt):
    tcPr = cell._tc.get_or_add_tcPr()
    tag = SIDES[side]
    for e in tcPr.findall(qn(tag)):
        tcPr.remove(e)
    xml = ('<%s xmlns:a="%s" w="%d" cap="flat" cmpd="sng" algn="ctr">'
           '<a:solidFill><a:srgbClr val="%02X%02X%02X"/></a:solidFill>'
           '<a:prstDash val="solid"/></%s>'
           % (tag, A, int(width_pt * 12700), color[0], color[1], color[2], tag))
    _insert_ordered(tcPr, parse_xml(xml), TCPR_ORDER)


def make_table(slide, data, x, y, w, col_w, row_h=0.55, head_h=0.62,
               head_fill=NAVY, head_color=WHITE, body_size=17, head_size=18,
               alt=CARD, radius=True):
    rows, cols = len(data), len(data[0])
    h = head_h + row_h * (rows - 1)
    gf = slide.shapes.add_table(rows, cols, Inches(x), Inches(y),
                                Inches(w), Inches(h))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    tblPr.set('rtl', '1')
    tblPr.set('firstRow', '1')
    tblPr.set('bandRow', '0')
    for i, cw in enumerate(col_w):
        tbl.columns[i].width = Inches(cw)
    tbl.rows[0].height = Inches(head_h)
    for i in range(1, rows):
        tbl.rows[i].height = Inches(row_h)
    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.margin_left = Inches(0.14)
            cell.margin_right = Inches(0.14)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.05)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if ri == 0:
                cell.fill.fore_color.rgb = head_fill
            else:
                cell.fill.fore_color.rgb = WHITE if ri % 2 else alt
            for side in 'LRTB':
                _cell_border(cell, side, BORDER, 0.75)
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            para(p, align='center', rtl=True, line_spacing=1.2)
            r = p.add_run()
            r.text = val
            style_run(r, size=head_size if ri == 0 else body_size,
                      bold=(ri == 0), color=head_color if ri == 0 else INK)
    return gf


# ----------------------------------------------------- transitions & anims ---
def add_transition(slide, kind='fade'):
    sld = slide._element
    for e in sld.findall(qn('p:transition')):
        sld.remove(e)
    xml = ('<p:transition xmlns:p="%s" xmlns:a="%s" spd="med"><p:fade/>'
           '</p:transition>' % (P, A))
    sld.append(parse_xml(xml))


def add_fade_anim(slide, shape_ids, dur=450, stagger=110, start=0):
    """Simple, uniform fade-in for the given shape ids (auto after slide
    appears, staggered)."""
    sld = slide._element
    for e in sld.findall(qn('p:timing')):
        sld.remove(e)
    ids = list(shape_ids)
    if not ids:
        return
    nid = [10]

    def nxt():
        nid[0] += 1
        return nid[0]

    effects = []
    for i, sid in enumerate(ids):
        t = nxt()
        d = start + i * stagger
        container = 'afterEffect' if i == 0 else 'withEffect'
        effects.append(
            '<p:par><p:cTn id="%d" presetID="10" presetClass="entr" '
            'presetSubtype="0" fill="hold" nodeType="%s">'
            '<p:stCondLst><p:cond delay="%d"/></p:stCondLst>'
            '<p:childTnLst>'
            '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold">'
            '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>style.visibility</p:attrName>'
            '</p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to>'
            '</p:set>'
            '<p:animEffect transition="in" filter="fade"><p:cBhvr>'
            '<p:cTn id="%d" dur="%d"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '</p:cBhvr></p:animEffect>'
            '</p:childTnLst></p:cTn></p:par>'
            % (t, container, d, nxt(), sid, nxt(), dur, sid))
    body = "".join(effects)
    xml = (
        '<p:timing xmlns:p="{p}" xmlns:a="{a}"><p:tnLst><p:par>'
        '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        '<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
        '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        '<p:par><p:cTn id="3" fill="hold"><p:stCondLst>'
        '<p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
        '<p:par><p:cTn id="4" fill="hold"><p:stCondLst>'
        '<p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        '{body}'
        '</p:childTnLst></p:cTn></p:par>'
        '</p:childTnLst></p:cTn></p:par>'
        '</p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl>'
        '<p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl>'
        '<p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
    ).format(p=P, a=A, body=body)
    sld.append(parse_xml(xml))


def enable_smooth_no(slide):
    pass


def set_solid(shape, color, alpha=None):
    """solid fill with optional transparency (alpha 0..100 = opacity)"""
    spPr = shape._element.spPr
    for tag in ('a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill',
                'a:pattFill', 'a:grpFill'):
        for e in spPr.findall(qn(tag)):
            spPr.remove(e)
    a = '<a:alpha val="%d"/>' % int(alpha * 1000) if alpha is not None else ''
    xml = ('<a:solidFill xmlns:a="%s"><a:srgbClr val="%02X%02X%02X">%s'
           '</a:srgbClr></a:solidFill>' % (A, color[0], color[1], color[2], a))
    _insert_ordered(spPr, parse_xml(xml),
                    ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill',
                     'a:pattFill', 'a:grpFill', 'a:ln', 'a:effectLst',
                     'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst'])
    return shape


def set_bg(slide, color=None, grad=None, angle=90):
    """slide background"""
    sld = slide._element
    csld = sld.find(qn('p:cSld'))
    for e in csld.findall(qn('p:bg')):
        csld.remove(e)
    if grad:
        inner = ('<a:gradFill rotWithShape="1"><a:gsLst>%s</a:gsLst>'
                 '<a:lin ang="%d" scaled="1"/></a:gradFill>' % (
                     ''.join('<a:gs pos="%d"><a:srgbClr val="%02X%02X%02X"/></a:gs>'
                             % (int(p * 100000), c[0], c[1], c[2]) for p, c in grad),
                     int(angle * 60000)))
    else:
        inner = ('<a:solidFill><a:srgbClr val="%02X%02X%02X"/></a:solidFill>'
                 % (color[0], color[1], color[2]))
    xml = ('<p:bg xmlns:p="%s" xmlns:a="%s"><p:bgPr>%s<a:effectLst/>'
           '</p:bgPr></p:bg>' % (P, A, inner))
    csld.insert(0, parse_xml(xml))


def icon_circle(slide, x, y, d, fill, glyph=None, color=WHITE, size=18,
                line=None, font=FONT):
    c = shape(slide, MSO_SHAPE.OVAL, x, y, d, d, fill=fill, line=line)
    if glyph:
        tf = c.text_frame
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        para(p, align='center', rtl=False, line_spacing=1.0)
        r = p.add_run()
        r.text = glyph
        style_run(r, size=size, bold=True, color=color, font=font)
    return c


def arrow(slide, x, y, w, h, color=BLUE_2, left=True, thick=0.55):
    """RTL flow arrow"""
    kind = MSO_SHAPE.LEFT_ARROW if left else MSO_SHAPE.RIGHT_ARROW
    s = shape(slide, kind, x, y, w, h, fill=color)
    try:
        s.adjustments[0] = thick
        s.adjustments[1] = 0.5
    except Exception:
        pass
    return s


def hline(slide, x, y, w, color=BORDER, width=1.0):
    s = shape(slide, MSO_SHAPE.RECTANGLE, x, y, w, 0.012, fill=color)
    return s


def chevron_tag(slide, x, y, w, h, text, fill=BLUE_PALE, color=BLUE, size=13):
    s = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=fill,
              radius=0.45)
    tf = s.text_frame
    tf.margin_left = tf.margin_right = Inches(0.06)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    para(p, align='center', rtl=True, line_spacing=1.1)
    r = p.add_run()
    r.text = text
    style_run(r, size=size, bold=True, color=color)
    return s
