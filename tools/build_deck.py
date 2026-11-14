#!/usr/bin/env python3
"""Rebuild + redesign of 'neural-network-inversion.pptx'.

Content is preserved (all Persian text, tables, screenshots); the visual layer
(layout, palette, typography, spacing, cards, shadows, transitions, simple
uniform fade animations) is new.
"""
import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import *   # noqa
import deckkit as K

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MEDIA = os.path.join(HERE, 'assets', 'media')
OUT = os.path.join(HERE, 'neural-network-inversion-upgraded.pptx')

SW, SH = 13.333, 7.5
M = 0.72                      # outer margin
CW = SW - 2 * M               # content width  (11.89)

SLIDES_ANIM = {}              # slide-id -> [shape ids] collected for animations
ANIM_SLIDES = {}              # slide-id -> slide


# --------------------------------------------------------------- helpers ---
def new_deck():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    return prs


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def light_bg(slide):
    set_bg(slide, grad=[(0, RGBColor(0xFF, 0xFF, 0xFF)),
                        (1, RGBColor(0xEE, 0xF3, 0xFB))], angle=115)


def subtle_dots(slide, x, y, w, h, step=0.34, color=RGBColor(0xC9, 0xDC, 0xF3)):
    """very faint dot grid made of tiny squares (cheap, prints fine)"""
    cols = int(w / step)
    rows = int(h / step)
    for r in range(rows):
        for c in range(cols):
            shape(slide, MSO_SHAPE.OVAL, x + c * step, y + r * step, 0.035, 0.035,
                  fill=color)


def header(slide, title, kicker=None, num=None, sub=None):
    """standard slide header (RTL): kicker + title + accent rule on the right"""
    ty = 0.44
    if kicker:
        add_text(slide, M, 0.40, CW, 0.30,
                 [dict(t=kicker, size=13, bold=True, color=BLUE_2, align='right')])
        ty = 0.70
    add_text(slide, M, ty, CW, 0.66,
             [dict(t=title, size=30, bold=True, color=NAVY, align='right',
                   line_spacing=1.15)])
    # accent rule under the title, right aligned
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SW - M - 1.55, ty + 0.72, 1.55,
          0.055, fill=BLUE, radius=0.5)
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, SW - M - 2.05, ty + 0.72, 0.35,
          0.055, fill=GOLD, radius=0.5)
    if num is not None:
        add_text(slide, M, 0.42, 0.9, 0.3,
                 [dict(t=num, size=12, bold=True, color=RGBColor(0xA8, 0xB8, 0xCC),
                       align='left', rtl=False)])
    if sub:
        add_text(slide, M, ty + 0.86, CW, 0.34,
                 [dict(t=sub, size=15, color=MUTED, align='right')])
        return ty + 1.24
    return ty + 0.95


def bullets(slide, x, y, w, lines, size=20, gap=0.13, color=INK, rtl=True,
            dot=BLUE_2, indent=0.30):
    """RTL bullet list; returns total height used"""
    total = 0
    for ln in lines:
        tb = add_text(slide, x, y, w, 3.0,
                      [dict(t=ln, size=size, color=color, align='right', rtl=rtl,
                            line_spacing=1.4, runs=ln if isinstance(ln, list) else None)])
        h = tb.text_frame.paragraphs[0]
        # measure: estimate height from python-pptx-less heuristic
        est = est_lines(ln if isinstance(ln, str) else ''.join(x[0] for x in ln),
                        size, w - indent)
        hh = est * size * 1.4 / 72
        shape(slide, MSO_SHAPE.OVAL, x + w - indent * 0.55, y + size / 72 * 0.42,
              0.085, 0.085, fill=dot)
        tb.left = Inches(x + indent * 0.0)
        total += hh + gap
        y += hh + gap
    return total


def est_lines(text, size, width_in, factor=1.0):
    """rough Persian text metrics: avg char width ~0.52em"""
    if not text:
        return 1
    per_line = max(6, int(width_in / (size * 0.52 * factor / 72)))
    n = 0
    for part in text.split('\n'):
        n += max(1, -(-len(part) // per_line))
    return n


def footer(slide, i, note='معکوس‌سازی شبکه‌های عصبی'):
    hline(slide, M, SH - 0.62, CW, color=RGBColor(0xDD, 0xE7, 0xF5), width=1.0)
    add_text(slide, M, SH - 0.55, 5.0, 0.28,
             [dict(t=note, size=11, color=RGBColor(0x93, 0xA5, 0xBE), align='left')])
    add_text(slide, SW - M - 1.2, SH - 0.55, 1.2, 0.28,
             [dict(t=str(i).zfill(2), size=11, bold=True,
                   color=RGBColor(0x93, 0xA5, 0xBE), align='right', rtl=False)])


def stat_pill(slide, x, y, w, h, label, value, fill=WHITE, accent=BLUE,
              vsize=26, lsize=13):
    c = card(slide, x, y, w, h, fill=fill, line=BORDER, radius=0.09)
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x + w - 0.055, y + 0.12, 0.055,
          h - 0.24, fill=accent, radius=0.5)
    add_text(slide, x + 0.14, y + 0.10, w - 0.30, h - 0.20,
             [dict(t=value, size=vsize, bold=True, color=NAVY, align='right',
                   line_spacing=1.05),
              dict(t=label, size=lsize, color=MUTED, align='right',
                   space_before=2, line_spacing=1.2)])
    return c


def anim(slide, ids):
    bucket = SLIDES_ANIM.setdefault(id(slide), [])
    for i in ids:
        if isinstance(i, int) and i > 0 and i not in bucket:
            bucket.append(i)
    ANIM_SLIDES[id(slide)] = slide


ART = os.path.join(HERE, 'assets')


def diamond(slide, x, y, size, color):
    return shape(slide, MSO_SHAPE.DIAMOND, x, y, size, size, fill=color)
def img(n):
    return os.path.join(MEDIA, n)


def image_card(slide, path, x, y, w, h, pad=0.16, dark=False, radius=0.05,
               shadow=True, mode='contain'):
    """picture inside a framed card (consistent gutter around photos)"""
    outer = card(slide, x, y, w, h, fill=DARK if dark else WHITE,
                 radius=radius, line=BORDER, shadow=shadow)
    if mode == 'cover':
        add_pic(slide, path, x + pad, y + pad, w - 2 * pad, h - 2 * pad,
                radius=0.035, border=None, shadow=False,
                crop=fit(path, 0, 0, w - 2 * pad, h - 2 * pad, 'cover')[4])
        return outer
    ix, iy, iw, ih = fit(path, x + pad, y + pad, w - 2 * pad, h - 2 * pad,
                         'contain')
    add_pic(slide, path, ix, iy, iw, ih, radius=0.02,
            border=None if dark else BORDER, shadow=False)
    return outer


def cover_card(slide, path, x, y, w, h, radius=0.05, shadow=True, pad=0.05):
    card(slide, x, y, w, h, fill=WHITE, radius=radius, line=BORDER, shadow=shadow)
    add_pic(slide, path, x + pad, y + pad, w - 2 * pad, h - 2 * pad,
            radius=0.04, border=None, shadow=False,
            crop=fit(path, 0, 0, w - 2 * pad, h - 2 * pad, 'cover')[4])


# =============================================================== slide 1 ===
def s01(prs, i):
    s = blank(prs)
    set_bg(s, RGBColor(0xFF, 0xFF, 0xFF))
    add_pic(s, os.path.join(ART, 'cover_bg.png'), 0, 0, SW, SH, radius=None,
            border=None, shadow=False)
    # top ornament
    add_text(s, 1.2, 1.62, SW - 2.4, 0.5,
             [dict(t='بسمه تعالی', size=16, color=GOLD, align='center', bold=True)])
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, SW / 2 - 0.75, 2.16, 1.5, 0.045,
          fill=GOLD, radius=0.5)
    diamond(s, SW / 2 - 0.075, 2.13, 0.15, GOLD)
    add_text(s, 1.2, 2.52, SW - 2.4, 1.20,
             [dict(t='به نام خالق صفر و یک', size=52, bold=True, color=NAVY,
                   align='center', line_spacing=1.2)])
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, SW / 2 - 1.35, 3.92, 2.7, 0.045,
          fill=BLUE_2, radius=0.5)
    add_text(s, 2.4, 4.20, SW - 4.8, 0.5,
             [dict(t='دانشگاه شهید باهنر شیراز  ·  درس تجزیه و تحلیل سیستم‌ها',
                   size=17, color=MUTED, align='center')])
    chip = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, SW / 2 - 1.9, 4.78, 3.8, 0.48,
                 fill=WHITE, radius=0.4, line=BORDER)
    add_text(s, SW / 2 - 1.8, 4.87, 3.6, 0.32,
             [dict(t='ارائه دانشگاهی  ·  هوش مصنوعی', size=13, bold=True,
                   color=BLUE, align='center')])
    footer_note = 'علیرضا خدادادی  ·  ۰۴۲۱۱۰۸۵۳۳۴۰۳۲'
    add_text(s, 2.4, 5.62, SW - 4.8, 0.4,
             [dict(t=footer_note, size=14, color=MUTED, align='center')])
    return s


# =============================================================== slide 2 ===
def s02(prs, i):
    s = blank(prs)
    set_bg(s, grad=[(0, RGBColor(0xF7, 0xFA, 0xFF)), (1, RGBColor(0xE3, 0xEE, 0xFB))],
           angle=120)
    # left visual panel
    px_, py, pw, ph = 0.55, 0.55, 5.15, SH - 1.10
    card(s, px_, py, pw, ph, fill=WHITE, radius=0.045, shadow=True, shadow_alpha=22)
    add_pic(s, os.path.join(ART, 'panel_art.png'), px_ + 0.04, py + 0.04,
            pw - 0.08, ph - 0.08, radius=0.045, border=None, shadow=False,
            crop=fit(os.path.join(ART, 'panel_art.png'), 0, 0, pw - 0.08,
                     ph - 0.08, 'cover')[4])
    tag = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, px_ + 0.30, py + 0.34,
                pw - 0.60, 0.44, fill=WHITE, radius=0.22)
    set_solid(tag, WHITE, alpha=82)
    add_text(s, px_ + 0.35, py + 0.42, pw - 0.70, 0.30,
             [dict(t='شبکه عصبی  |  معکوس‌سازی  |  داده اصلی', size=12, bold=True,
                   color=NAVY, align='center')])

    # right text column
    rx = 6.10
    rw = SW - rx - 0.78
    chevron_tag(s, rx + rw - 1.80, 0.80, 1.80, 0.36, 'ارائه دانشگاهی', size=12)
    add_text(s, rx, 1.24, rw, 0.62,
             [dict(t='معکوس‌سازی شبکه‌های عصبی', size=33, bold=True, color=NAVY,
                   align='right', line_spacing=1.25)])
    add_text(s, rx, 1.98, rw, 0.40,
             [dict(t='از شبکه‌های عصبی تا معکوس کردن آن‌ها', size=18, color=BLUE,
                   align='right')])
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, rx + rw - 2.0, 2.52, 2.0, 0.04,
          fill=GOLD, radius=0.5)

    info = [('دانشجو', 'علیرضا خدادادی'),
            ('کد دانشجویی', '۰۴۲۱۱۰۸۵۳۳۴۰۳۲'),
            ('دانشگاه', 'شهید باهنر شیراز'),
            ('درس', 'تجزیه و تحلیل سیستم ها'),
            ('استاد', 'محمد رحمان ستایش شیرازی')]
    cy = 2.80
    for k, v in info:
        c = card(s, rx, cy, rw, 0.62, fill=WHITE, radius=0.10, line=BORDER,
                 shadow=False)
        add_text(s, rx + 0.20, cy + 0.15, rw - 0.40, 0.34,
                 [dict(runs=[(v, dict(size=16, bold=True, color=INK)),
                             ('      ' + k + ' :', dict(size=14, color=MUTED))],
                       align='right', line_spacing=1.1)])
        cy += 0.76
    return s


# =============================================================== slide 3 ===
def s03(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'فهرست مطالب', kicker='سرفصل‌های ارائه', num='03')
    items = ['سیستم عصبی انسان و نورون‌های بیولوژیکی',
             'شبکه‌های عصبی مصنوعی',
             'معکوس‌سازی شبکه‌های عصبی',
             'کاربردها و ابزارهای واقعی',
             'چالش‌ها و آینده',
             'جمع‌بندی',
             'منابع',
             'پرسش و پاسخ']
    cw_, gap = 5.72, 0.44
    rh, vgap = 1.00, 0.26
    for idx, it in enumerate(items):
        col = idx // 4
        row = idx % 4
        x = SW - M - cw_ - col * (cw_ + gap)
        y = top + 0.18 + row * (rh + vgap)
        c = card(s, x, y, cw_, rh, fill=WHITE, radius=0.10, line=BORDER)
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + cw_ - 0.78, y + (rh - 0.60) / 2,
              0.60, 0.60, fill=BLUE_PALE, radius=0.24)
        add_text(s, x + cw_ - 0.78, y + (rh - 0.60) / 2 + 0.14, 0.60, 0.32,
                 [dict(t=str(idx + 1).zfill(2), size=15, bold=True, color=BLUE,
                       align='center', rtl=False)])
        add_text(s, x + 0.28, y, cw_ - 1.30, rh,
                 [dict(t=it, size=17, bold=True, color=NAVY, align='right')],
                 anchor='middle')
        anim(s, [c.shape_id])
    footer(s, i)
    return s


# =============================================================== slide 4 ===
def s04(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'مقدمه و انگیزه', kicker='چرا این موضوع مهم است؟', num='04')
    items = [('کاربرد گسترده', 'شبکه‌های عصبی امروزه در تشخیص چهره، پزشکی، خودروهای خودران و بسیاری از سیستم‌های مهم استفاده می‌شوند.'),
             ('داده‌های حساس', 'این شبکه‌ها معمولاً با داده‌های خصوصی و حساس آموزش می‌بینند.'),
             ('پرسش کلیدی', 'آیا می‌توان از روی خروجی یک مدل، به ورودی‌های اصلی آن پی برد؟'),
             ('دو لبه یک شمشیر', 'معکوس‌سازی به این سوال پاسخ می‌دهد؛ هم تهدید حریم خصوصی است و هم کاربرد علمی دارد.')]
    cw_, ch = 5.80, 1.55
    for k, (t, txt) in enumerate(items):
        col, row = k % 2, k // 2
        x = SW - M - cw_ - col * (cw_ + 0.30)
        y = top + 0.10 + row * (ch + 0.28)
        c = card(s, x, y, cw_, ch, fill=WHITE, radius=0.06, line=BORDER)
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + cw_ - 0.055, y + 0.20, 0.055,
              ch - 0.40, fill=[BLUE, TEAL, GOLD, NAVY][k], radius=0.5)
        add_text(s, x + 0.30, y + 0.22, cw_ - 0.65, 0.32,
                 [dict(t=t, size=16, bold=True, color=NAVY, align='right')])
        add_text(s, x + 0.30, y + 0.62, cw_ - 0.65, ch - 0.78,
                 [dict(t=txt, size=15, color=INK, align='right', line_spacing=1.4)])
        anim(s, [c.shape_id])
    footer(s, i)
    return s


# =============================================================== slide 5 ===
def s05(prs, i):
    s = blank(prs)
    set_bg(s, grad=[(0, RGBColor(0xFB, 0xFD, 0xFF)), (1, RGBColor(0xEA, 0xF1, 0xFA))],
           angle=110)
    # right column: heading + short intro, left: wide image
    add_text(s, 7.05, 0.85, 5.55, 0.32,
             [dict(t='بخش اول', size=13, bold=True, color=GOLD, align='right')])
    add_text(s, 7.05, 1.20, 5.55, 1.45,
             [dict(t='سیستم عصبی انسان و نورون‌های بیولوژیکی', size=30, bold=True,
                   color=NAVY, align='right', line_spacing=1.25)])
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 11.05, 2.78, 1.55, 0.055, fill=BLUE,
          radius=0.5)
    add_text(s, 7.05, 3.05, 5.55, 0.80,
             [dict(t='نگاهی ساده به ساختار مغز', size=19, color=BLUE,
                   align='right', line_spacing=1.3)])
    for k, txt in enumerate(['میلیاردها نورون در شبکه‌ای پیچیده',
                            'ارتباط الکتریکی و شیمیایی بین سلول‌ها',
                            'الهام‌بخش ساخت شبکه‌های عصبی مصنوعی']):
        y = 3.95 + k * 0.76
        shape(s, MSO_SHAPE.OVAL, 12.32, y + 0.16, 0.09, 0.09, fill=BLUE_2)
        add_text(s, 7.05, y, 5.15, 0.40,
                 [dict(t=txt, size=16, color=MUTED, align='right')])
    image_card(s, img('image3.png'), 0.62, 1.62, 6.05, 3.55, pad=0.06,
               dark=True, radius=0.05, mode='cover')
    # stat strip under the picture
    stats = [('۸۶ میلیارد', 'نورون در مغز'), ('۱۰۰ تریلیون', 'اتصال سیناپسی'),
             ('~۲۰ وات', 'مصرف انرژی')]
    for k, (v, lab) in enumerate(stats):
        x = 0.62 + k * 2.02
        c = card(s, x, 5.42, 1.90, 0.92, fill=WHITE, radius=0.10, line=BORDER,
                 shadow=False)
        add_text(s, x + 0.10, 5.50, 1.70, 0.36,
                 [dict(t=v, size=17, bold=True, color=BLUE, align='center')])
        add_text(s, x + 0.10, 5.86, 1.70, 0.30,
                 [dict(t=lab, size=12, color=MUTED, align='center')])
    footer(s, i)
    return s


# =============================================================== slide 6 ===
def s06(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'نورون بیولوژیکی چگونه کار می‌کند؟',
                 kicker='دندریت  ·  جسم سلولی  ·  آکسون', num='06')
    image_card(s, img('image4.jpg'), M, top + 0.12, 6.15, 3.60, pad=0.05,
               radius=0.05)
    add_text(s, M, top + 3.80, 6.15, 0.32,
             [dict(t='ساختار یک نورون بیولوژیکی و اجزای اصلی آن', size=13,
                   color=MUTED, align='center')])
    x = 7.25
    w = SW - x - M
    card(s, x, top + 0.12, w, 4.28, fill=WHITE, radius=0.05, line=BORDER)
    cy = top + 0.42
    add_text(s, x + 0.32, cy, w - 0.64, 0.34,
             [dict(t='نورون واحد اصلی سیستم عصبی است', size=17, bold=True,
                   color=NAVY, align='right')])
    cy += 0.52
    parts = [('دندریت‌ها', 'دریافت پیام از نورون‌های دیگر', BLUE),
             ('جسم سلولی', 'پردازش و جمع‌بندی اطلاعات', TEAL),
             ('آکسون', 'ارسال پیام به نورون‌های بعدی', GOLD)]
    for name, desc, col in parts:
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + w - 0.30, cy + 0.10, 0.10,
              0.52, fill=col, radius=0.5)
        add_text(s, x + 0.30, cy, w - 0.70, 0.34,
                 [dict(t=name, size=16, bold=True, color=col, align='right')])
        add_text(s, x + 0.30, cy + 0.34, w - 0.70, 0.32,
                 [dict(t=desc, size=14, color=MUTED, align='right')])
        cy += 0.86
    hline(s, x + 0.32, cy - 0.06, w - 0.64, color=BORDER)
    add_text(s, x + 0.32, cy + 0.10, w - 0.64, 0.80,
             [dict(t='نورون‌ها از طریق اتصال به هم، شبکه‌های پیچیده‌ای می‌سازند که اساس تفکر، حافظه و تصمیم‌گیری هستند.',
                   size=15, color=INK, align='right', line_spacing=1.45)])
    footer(s, i)
    return s


# =============================================================== slide 7 ===
def s07(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'از مغز انسان تا شبکه‌های عصبی مصنوعی',
                 kicker='الهام زیستی، ساده‌سازی مهندسی', num='07')
    w = 5.35
    h = 3.50
    y = top + 0.15
    c1 = image_card(s, img('image5.jpg'), SW - M - w, y, w, h, pad=0.05, dark=True)
    c2 = image_card(s, img('image6.png'), M, y, w, h, pad=0.05, dark=True)
    add_text(s, SW - M - w, y + h + 0.14, w, 0.34,
             [dict(t='مغز انسان', size=16, bold=True, color=NAVY, align='center')])
    add_text(s, M, y + h + 0.14, w, 0.34,
             [dict(t='شبکه عصبی مصنوعی', size=16, bold=True, color=NAVY,
                   align='center')])
    # arrow between
    ax = M + w + 0.42
    aw = SW - 2 * M - 2 * w - 0.84
    shape(s, MSO_SHAPE.LEFT_ARROW, ax, y + h / 2 - 0.28, aw, 0.56, fill=BLUE_2)
    add_text(s, ax, y + h / 2 - 0.72, aw, 0.34,
             [dict(t='الهام', size=13, bold=True, color=BLUE_2, align='center')])
    # callout
    call = card(s, M + 0.6, y + h + 0.62, SW - 2 * M - 1.2, 0.62, fill=CARD_2,
                radius=0.14, line=BORDER, shadow=False)
    add_text(s, M + 0.85, y + h + 0.75, SW - 2 * M - 1.7, 0.40,
             [dict(t='دانشمندان با الهام از ساختار مغز، مدل ساده‌شده‌ای به نام شبکه عصبی مصنوعی ساختند!',
                   size=17, bold=True, color=BLUE, align='center')])
    footer(s, i)
    return s



def art_band(slide, path, width=4.35):
    """Left decorative band: network motif faded out toward the text side."""
    add_pic(slide, path, 0.0, 0.0, width, SH, radius=None, border=None,
            shadow=False, crop=fit(path, 0, 0, width, SH, 'cover')[4])
    for stop, alpha, w_, x_ in [(0.00, 100, 0.85, 0.00), (0.42, 100, 0.0, 0.0)]:
        pass
    # a soft vertical fade is enough: white overlay, opaque at the seam
    veil = shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, width, SH, fill=WHITE)
    set_grad(veil, [(0.0, WHITE, 0), (0.55, WHITE, 0), (0.80, WHITE, 40),
                    (0.94, WHITE, 88), (1.0, WHITE, 100)], angle=0)
    # thin accent seam
    shape(slide, MSO_SHAPE.RECTANGLE, width - 0.03, 1.30, 0.03, SH - 2.60,
          fill=BLUE_2)
    set_solid(slide.shapes[-1], BLUE_2, alpha=22)


def section_slide(prs, i, num_fa, kicker, title, sub, art='section_art.png'):
    s = blank(prs)
    set_bg(s, grad=[(0, RGBColor(0xFB, 0xFD, 0xFF)),
                    (1, RGBColor(0xE6, 0xEF, 0xFA))], angle=110)
    art_band(s, os.path.join(ART, art), 4.35)
    # content on the right
    x = 5.05
    w = SW - x - M
    add_text(s, x, 2.05, w, 0.42,
             [dict(t=kicker, size=14, bold=True, color=GOLD, align='right')])
    add_text(s, x, 2.55, w, 1.65,
             [dict(t=title, size=38, bold=True, color=NAVY, align='right',
                   line_spacing=1.25)])
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, SW - M - 1.9, 4.32, 1.9, 0.055,
          fill=BLUE, radius=0.5)
    add_text(s, x, 4.55, w, 0.75,
             [dict(t=sub, size=18, color=MUTED, align='right', line_spacing=1.4)])
    chip = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, SW - M - 0.85, 6.10, 0.85,
                 0.85, fill=WHITE, radius=0.5, line=BORDER)
    add_text(s, SW - M - 0.85, 6.32, 0.85, 0.45,
             [dict(t=num_fa, size=22, bold=True, color=RGBColor(0xC2, 0xD2, 0xE8),
                   align='center', rtl=False)])
    return s


# =============================================================== slide 8 ===
def s08(prs, i):
    s = section_slide(prs, i, '۰۲', 'بخش دوم', 'شبکه‌های عصبی مصنوعی',
                      'مفاهیم پایه + مثال‌های کلیدی')
    return s


# =============================================================== slide 9 ===
def s09(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'شبکه عصبی مصنوعی چیست؟',
                 kicker='مدلی محاسباتی با الهام از مغز', num='09')
    x = 0.72
    w = 5.55
    card(s, x, top + 0.15, w, 4.05, fill=WHITE, radius=0.05, line=BORDER)
    cy = top + 0.58
    for k, (t, d) in enumerate([
            ('مدل محاسباتی الهام‌گرفته از مغز',
             'لایه‌هایی از نورون‌های مصنوعی که به هم متصل‌اند'),
            ('مثال کلاسیک: تشخیص اعداد دست‌نویس',
             'همان «Hello World» دنیای هوش مصنوعی'),
            ('تصویر ۲۸×۲۸ پیکسل = ۷۸۴ عدد',
             'هر پیکسل یک عدد بین ۰ تا ۱ است'),
            ('هر پیکسل ← یک نورون در لایه ورودی',
             'خروجی، احتمال هر رقم را نشان می‌دهد')]):
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + w - 0.34, cy + 0.06, 0.10,
              0.66, fill=[BLUE, TEAL, GOLD, NAVY][k], radius=0.5)
        add_text(s, x + 0.30, cy, w - 0.80, 0.32,
                 [dict(t=t, size=15, bold=True, color=NAVY, align='right')])
        add_text(s, x + 0.30, cy + 0.32, w - 0.80, 0.30,
                 [dict(t=d, size=13, color=MUTED, align='right')])
        cy += 0.88
    image_card(s, img('image8.png'), 6.55, top + 0.15, SW - 6.55 - M, 4.05,
               pad=0.05, dark=True)
    add_text(s, 6.55, top + 4.28, SW - 6.55 - M, 0.32,
             [dict(t='تصویر ۲۸×۲۸ پیکسل به ۷۸۴ عدد تبدیل می‌شود', size=13,
                   color=MUTED, align='center')])
    footer(s, i)
    return s


# ============================================================== slide 10 ===
def s10(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'اجزای اصلی یک شبکه عصبی',
                 kicker='سه جزء کلیدی در هر نورون مصنوعی', num='10')
    # left: neuron diagram card
    iw, ih = pic_size(img('image9.png'))
    dw = 5.00
    dh = dw * ih / iw + 0.14
    dy = top + 0.26 + (4.10 - dh) / 2
    image_card(s, img('image9.png'), M, dy, dw, dh, pad=0.06)
    add_text(s, M, dy + dh + 0.18, dw, 0.34,
             [dict(t='ساختار یک نورون مصنوعی', size=12, color=MUTED,
                   align='center')])
    # right: three component cards
    items = [('وزن', 'Weight', 'میزان تأثیر هر ورودی روی نورون بعدی', BLUE),
             ('بایاس', 'Bias',
              'مقداری که به مجموع وزن‌دار اضافه می‌شود تا آستانه فعال شدن نورون تنظیم شود',
              TEAL),
             ('تابع فعال‌سازی', 'Activation Function',
              'خروجی نورون را در یک بازه مشخص (۰ تا ۱) فشرده می‌کند و غیرخطی بودن را وارد مدل می‌کند',
              GOLD)]
    x = M + dw + 0.42
    w = SW - x - M
    y = top + 0.24
    for t, en, dsc, col in items:
        c = card(s, x, y, w, 1.28, fill=WHITE, radius=0.06, line=BORDER)
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + w - 0.055, y + 0.24, 0.055,
              0.80, fill=col, radius=0.5)
        add_text(s, x + 0.28, y + 0.16, w - 0.58, 0.34,
                 [dict(runs=[(t, dict(size=17, bold=True, color=NAVY)),
                             ('   ' + en, dict(size=11, color=col))],
                       align='right', line_spacing=1.2, rtl=False)])
        add_text(s, x + 0.28, y + 0.56, w - 0.58, 0.60,
                 [dict(t=dsc, size=13, color=INK, align='right',
                       line_spacing=1.4)])
        anim(s, [c.shape_id])
        y += 1.42
    footer(s, i)
    return s


# ============================================================== slide 11 ===
def s11(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'لایه‌های شبکه عصبی',
                 kicker='از داده خام تا نتیجه نهایی', num='11')
    image_card(s, img('image10.png'), M, top + 0.18, 6.35, 4.05, pad=0.05,
               dark=True)
    x = 7.40
    w = SW - x - M
    layers = [('لایه ورودی', 'داده‌های خام را دریافت می‌کند (مثلاً پیکسل‌های تصویر)', BLUE),
              ('لایه‌های پنهان', 'محاسبات و استخراج ویژگی‌ها را انجام می‌دهند', TEAL),
              ('لایه خروجی', 'نتیجه نهایی را تولید می‌کند (مثلاً گربه یا سگ بودن عکس)', GOLD)]
    y = top + 0.18
    for t, d, col in layers:
        c = card(s, x, y, w, 1.24, fill=WHITE, radius=0.06, line=BORDER)
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + w - 0.06, y + 0.20, 0.06,
              0.84, fill=col, radius=0.5)
        add_text(s, x + 0.28, y + 0.20, w - 0.55, 0.34,
                 [dict(t=t, size=17, bold=True, color=NAVY, align='right')])
        add_text(s, x + 0.28, y + 0.58, w - 0.55, 0.56,
                 [dict(t=d, size=13, color=MUTED, align='right',
                       line_spacing=1.35)])
        y += 1.40
    footer(s, i)
    return s


# ============================================================== slide 12 ===
def s12(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'نورون‌های لایه بعد چه چیزی می‌بینند؟',
                 kicker='وزن‌ها به شکل تصویر', num='12')
    cols = [(img('image11.png'), img('image14.png'),
             'هر نورون در لایه بعدی به تمام نورون‌های لایه قبلی متصل است.'),
            (img('image12.png'), img('image15.png'),
             'وزن‌های این اتصالات را می‌توان به شکل یک تصویر جدید دید.'),
            (img('image13.png'), img('image16.png'),
             'در واقع هر نورون لایه بعد، به دنبال یک الگوی خاص در تصویر ورودی می‌گردد.')]
    cw_ = 3.72
    ch = 4.35
    for k, (diag, patt, txt) in enumerate(cols):
        x = SW - M - cw_ - k * (cw_ + 0.36)
        y = top + 0.14
        c = card(s, x, y, cw_, ch, fill=WHITE, radius=0.05, line=BORDER)
        image_card(s, diag, x + 0.16, y + 0.16, 1.72, 2.90, pad=0.04,
                   dark=True, shadow=False)
        image_card(s, patt, x + 1.94, y + 0.16, 1.62, 2.90, pad=0.04,
                   dark=True, shadow=False)
        add_text(s, x + 0.22, y + 3.18, cw_ - 0.44, 1.00,
                 [dict(t=txt, size=13, color=INK, align='right',
                       line_spacing=1.5)])
        anim(s, [c.shape_id])
    footer(s, i)
    return s


# ============================================================== slide 13 ===
def s13(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'شباهت‌ها و تفاوت‌ها با مغز انسان',
                 kicker='مقایسه دو سیستم', num='13')
    data = [['مغز انسان', 'شبکه عصبی مصنوعی'],
            ['میلیاردها نورون واقعی', 'تعداد محدودی نورون مصنوعی'],
            ['یادگیری بسیار پیچیده', 'یادگیری با تنظیم وزن‌ها'],
            ['مصرف انرژی بسیار کم', 'نیاز به توان محاسباتی بالا'],
            ['انعطاف‌پذیری بسیار زیاد', 'تخصصی برای یک وظیفه خاص']]
    make_table(s, data, M + 1.05, top + 0.34, CW - 2.10, [5.55, 5.55],
               row_h=0.82, head_h=0.70, body_size=18, head_size=18)
    add_text(s, M + 1.05, top + 4.58, CW - 2.10, 0.34,
             [dict(t='شبکه عصبی مصنوعی، ساده‌سازی مهندسی‌شده‌ای از مغز است و نه کپی آن.',
                   size=14, color=MUTED, align='center')])
    footer(s, i)
    return s


# ============================================================== slide 14 ===
def s14(prs, i):
    s = blank(prs)
    set_bg(s, grad=[(0, RGBColor(0xFB, 0xFD, 0xFF)),
                    (1, RGBColor(0xE6, 0xEF, 0xFA))], angle=110)
    art_band(s, os.path.join(ART, 'section_art.png'), 4.35)
    x = 5.05
    w = SW - x - M
    add_text(s, x, 1.55, w, 0.42,
             [dict(t='بخش سوم', size=14, bold=True, color=GOLD, align='right')])
    add_text(s, x, 2.00, w, 1.30,
             [dict(t='معکوس‌سازی شبکه‌های عصبی', size=36, bold=True, color=NAVY,
                   align='right', line_spacing=1.25)])
    add_text(s, x, 3.15, w, 0.45,
             [dict(t='از خروجی به ورودی', size=21, color=BLUE, align='right')])
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, SW - M - 1.9, 3.78, 1.9, 0.055,
          fill=BLUE, radius=0.5)
    # the forward / backward propagation diagram
    iw, ih = pic_size(img('image24.png'))
    dw = min(w, 4.55)
    dh = dw / (iw / ih) + 0.24
    dx = x + (w - dw) / 2
    image_card(s, img('image24.png'), dx, 4.05, dw, dh, pad=0.10, shadow=True)
    add_text(s, dx, 4.05 + dh + 0.14, dw, 0.30,
             [dict(t='مسیر رو به جلو و پس‌رو در شبکه عصبی', size=12,
                   color=MUTED, align='center')])
    shape(s, MSO_SHAPE.OVAL, SW - M - 0.85, 6.10, 0.85, 0.85, fill=WHITE,
          line=BORDER)
    add_text(s, SW - M - 0.85, 6.32, 0.85, 0.45,
             [dict(t='۰۳', size=22, bold=True, color=RGBColor(0xC2, 0xD2, 0xE8),
                   align='center', rtl=False)])
    return s


# ============================================================== slide 15 ===
def s15(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'معکوس‌سازی چیست؟',
                 kicker='از خروجی به ورودی', num='15')
    panels = [('شبکه عصبی معمولی', 'ورودی', 'خروجی', BLUE,
               'مسیر رو به جلو (Forward): از داده به پیش‌بینی'),
              ('معکوس‌سازی', 'خروجی', 'ورودی', TEAL,
               'مسیر برگشتی (Inverse): از پیش‌بینی به داده')]
    cw_ = 5.85
    for k, (title, a, b, col, cap) in enumerate(panels):
        x = SW - M - cw_ - k * (cw_ + 0.28)
        y = top + 0.10
        c = card(s, x, y, cw_, 2.60, fill=WHITE, radius=0.05, line=BORDER)
        add_text(s, x + 0.30, y + 0.26, cw_ - 0.60, 0.36,
                 [dict(t=title, size=21, bold=True, color=NAVY, align='right')])
        hline(s, x + 0.30, y + 0.74, cw_ - 0.60, color=BORDER)
        bw, bh = 1.32, 0.76
        by = y + 1.06
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + cw_ - 0.30 - bw, by, bw, bh,
              fill=BLUE_PALE, line=None, radius=0.18)
        add_text(s, x + cw_ - 0.30 - bw, by + 0.20, bw, 0.36,
                 [dict(t=a, size=16, bold=True, color=NAVY, align='center')])
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + 0.30, by, bw, bh,
              fill=DARK if k else CARD_2, line=None, radius=0.18)
        add_text(s, x + 0.30, by + 0.20, bw, 0.36,
                 [dict(t=b, size=16, bold=True,
                       color=WHITE if k else NAVY, align='center')])
        shape(s, MSO_SHAPE.LEFT_ARROW, x + 0.30 + bw + 0.22, by + 0.21,
              cw_ - 2 * bw - 1.14, 0.34, fill=col)
        add_text(s, x + 0.30, y + 1.98, cw_ - 0.60, 0.34,
                 [dict(t=cap, size=13, color=MUTED, align='center')])
        anim(s, [c.shape_id])
    add_text(s, SW - M - 8.6, top + 2.96, 8.6, 0.60,
             [dict(t='یعنی تلاش برای پیدا کردن اینکه چه ورودی باعث این خروجی شده است؟',
                   size=20, bold=True, color=BLUE, align='right',
                   line_spacing=1.3)])
    eq = card(s, M, top + 2.88, 3.05, 0.84, fill=CARD_2, radius=0.10,
              line=BORDER, shadow=False)
    add_text(s, M + 0.10, top + 3.02, 2.85, 0.56,
             [dict(t='y = f(x)   →   x̂ = f⁻¹(y)', size=18, bold=True,
                   color=NAVY, align='center', rtl=False)])
    footer(s, i)
    return s


COL_HEAD = [NAVY, NAVY]


# ============================================================== slide 16 ===
def s16(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'چرا معکوس کردن شبکه سخت است؟',
                 kicker='سه دلیل اصلی', num='16')
    w = 4.55
    image_card(s, img('image18.jpg'), M, top + 0.10, w, 4.05, pad=0.05,
               mode='cover')
    add_text(s, M, top + 4.28, w, 0.32,
             [dict(t='معکوس‌سازی دقیق، در عمل یک مسئله دشوار است', size=12,
                   color=MUTED, align='center')])
    x = M + w + 0.35
    rw = SW - x - M
    reasons = [('اطلاعات فشرده می‌شود',
                'شبکه‌های عصبی معمولاً اطلاعات را فشرده و خلاصه می‌کنند.', BLUE),
               ('چند ورودی، یک خروجی',
                'ممکن است چندین ورودی مختلف، یک خروجی یکسان بسازند.', TEAL),
               ('پاسخ یکتا وجود ندارد',
                'در نتیجه معکوس دقیق و یکتا وجود ندارد و باید تقریبی عمل کنیم.', GOLD)]
    cy = top + 0.10
    for t, d_, col in reasons:
        c = card(s, x, cy, rw, 1.30, fill=WHITE, radius=0.06, line=BORDER)
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + rw - 0.055, cy + 0.22, 0.055,
              0.86, fill=col, radius=0.5)
        add_text(s, x + 0.30, cy + 0.22, rw - 0.60, 0.34,
                 [dict(t=t, size=17, bold=True, color=NAVY, align='right')])
        add_text(s, x + 0.30, cy + 0.62, rw - 0.60, 0.56,
                 [dict(t=d_, size=14, color=MUTED, align='right',
                       line_spacing=1.4)])
        anim(s, [c.shape_id])
        cy += 1.44
    footer(s, i)
    return s


# ============================================================== slide 17 ===
def s17(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'دو رویکرد اصلی', kicker='مدل یا روش؟', num='17')
    items = [('معکوس‌سازی مدل‌های معمولی', 'روش',
              'مدل از قبل ساخته شده و ما با روش‌هایی سعی می‌کنیم آن را برعکس کنیم.',
              'تقریبی', TEAL, CARD_2, 'مثال: حملات معکوس‌سازی مدل'),
             ('شبکه‌های ذاتاً معکوس‌پذیر (INNs)', 'مدل',
              'از اول شبکه را طوری طراحی می‌کنیم که دقیقاً قابل برگشت باشد.',
              'دقیق', BLUE, WHITE, 'مثال: حل مسائل معکوس در فیزیک')]
    cw_ = 5.85
    for k, (t, kind, d_, acc, col, bgc, ex) in enumerate(items):
        x = SW - M - cw_ - k * (cw_ + 0.28)
        y = top + 0.26
        c = card(s, x, y, cw_, 3.55, fill=bgc, radius=0.05,
                 line=BORDER if k == 0 else BLUE_2)
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + cw_ - 1.20, y + 0.30, 0.95,
              0.34, fill=col, radius=0.5)
        add_text(s, x + cw_ - 1.20, y + 0.36, 0.95, 0.24,
                 [dict(t=kind, size=12, bold=True, color=WHITE, align='center')])
        add_text(s, x + 0.30, y + 0.72, cw_ - 0.60, 1.00,
                 [dict(t=t, size=19, bold=True, color=NAVY, align='right',
                       line_spacing=1.35)])
        hline(s, x + 0.30, y + 1.86, cw_ - 0.60, color=BORDER)
        add_text(s, x + 0.30, y + 2.02, cw_ - 0.60, 0.90,
                 [dict(t=d_, size=15, color=INK, align='right',
                       line_spacing=1.5)])
        add_text(s, x + 0.30, y + 2.94, cw_ - 0.60, 0.30,
                 [dict(t=ex, size=12, color=MUTED, align='right')])
        pill = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + cw_ - 2.28, y + 0.31,
                     0.95, 0.32, fill=col, radius=0.5)
        add_text(s, x + cw_ - 2.28, y + 0.36, 0.95, 0.24,
                 [dict(t='دقت ' + acc, size=11, bold=True, color=WHITE,
                       align='center')])
        anim(s, [c.shape_id])
    footer(s, i)
    return s


# ============================================================== slide 18 ===
def s18(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'مقایسه دو رویکرد', kicker='جمع‌بندی تفاوت‌ها', num='18')
    data = [['شبکه معکوس‌پذیر', 'مدل معمولی', 'مورد'],
            ['مدل / ساختار', 'روش', 'نوع'],
            ['دقیق', 'تقریبی', 'دقت معکوس'],
            ['تقریباً ندارد', 'دارد', 'از دست رفتن اطلاعات'],
            ['مسائل علمی و تولیدی', 'حمله و تفسیر', 'کاربرد اصلی']]
    make_table(s, data, M + 0.75, top + 0.28, CW - 1.50, [4.20, 4.20, 3.19],
               row_h=0.78, head_h=0.62, body_size=17, head_size=17)
    footer(s, i)
    return s


# ============================================================== slide 19 ===
def s19(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'مثال معروف: بازسازی چهره',
                 kicker='حریم خصوصی در عمل', num='19')
    # left: image
    image_card(s, img('image19.png'), M, top + 0.06, 5.30, 4.45, pad=0.05,
               mode='cover')
    add_text(s, M, top + 4.63, 5.30, 0.30,
             [dict(t='نمونه‌ای از بازسازی تصویر چهره از خروجی مدل', size=12,
                   color=MUTED, align='center')])
    # right: steps
    x = M + 5.30 + 0.42
    w = SW - x - M
    steps = ['فرض کنید یک مدل تشخیص هویت چهره داریم',
             'مهاجم فقط به خروجی مدل دسترسی دارد',
             'با معکوس‌سازی می‌تواند تصویری نزدیک به چهره واقعی فرد را بازسازی کند',
             'این مثال یکی از مهم‌ترین دلایل اهمیت موضوع از نظر حریم خصوصی است']
    cy = top + 0.06
    for n, t in enumerate(steps, 1):
        h = 0.98 if n < 4 else 1.08
        c = card(s, x, cy, w, h, fill=WHITE, radius=0.08, line=BORDER)
        shape(s, MSO_SHAPE.OVAL, x + w - 0.62, cy + (h - 0.44) / 2, 0.44, 0.44,
              fill=BLUE_PALE)
        add_text(s, x + w - 0.62, cy + (h - 0.44) / 2 + 0.09, 0.44, 0.28,
                 [dict(t=fa(n), size=14, bold=True, color=BLUE, align='center')])
        add_text(s, x + 0.25, cy, w - 1.00, h,
                 [dict(t=t, size=14, color=INK, align='right',
                       line_spacing=1.45)], anchor='middle')
        anim(s, [c.shape_id])
        cy += h + 0.16
    footer(s, i)
    return s


# ============================================================== slide 20 ===
def s20(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'کاربردهای معکوس‌سازی شبکه‌های عصبی',
                 kicker='فرصت یا تهدید؟', num='20')
    data = [['جنبه مثبت (فرصت)', 'جنبه منفی (تهدید)'],
            ['تفسیر و درک بهتر مدل‌ها', 'حملات حریم خصوصی (Model Inversion Attack)'],
            ['ویرایش تصویر در ابزارهای هوش مصنوعی', 'بازسازی داده‌های آموزشی خصوصی (مانند چهره)'],
            ['حل مسائل معکوس در پزشکی', 'نشت اطلاعات حساس از مدل‌های آموزش‌دیده'],
            ['کاربرد در ژئوفیزیک و تصویربرداری زیرزمینی', 'خطر برای سیستم‌های تشخیص هویت'],
            ['مدل‌سازی احتمالاتی و عدم قطعیت', 'امکان سوءاستفاده در جعل هویت و کلاهبرداری'],
            ['استفاده در فیزیک، هوافضا و طراحی مهندسی', 'تضعیف اعتماد عمومی به سیستم‌های هوش مصنوعی']]
    make_table(s, data, M + 0.55, top + 0.20, CW - 1.10, [5.40, 5.40],
               row_h=0.63, head_h=0.58, body_size=14, head_size=16)
    footer(s, i)
    return s


# ============================================================== slide 21 ===
def s21(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'چالش‌ها و آینده', kicker='پرسشی بزرگ', num='21')
    add_text(s, SW - M - 8.9, top - 0.04, 8.9, 0.46,
             [dict(t='آیا هوش مصنوعی در نهایت بر هوش طبیعی غلبه می‌کند؟', size=17,
                   bold=True, color=BLUE, align='right')])
    imgs = [(img('image23.jpg'), 'همکاری انسان و ماشین'),
            (img('image20.jpg'), 'هوش مصنوعی و هوش انسانی'),
            (img('image21.jpg'), 'مغز مصنوعی و مغز زیستی'),
            (img('image22.jpg'), 'آینده: رقابت یا هم‌افزایی؟')]
    cw_, ch = 2.86, 1.78
    for k, (p_, cap) in enumerate(imgs):
        col, row = k % 2, k // 2
        x = SW - M - cw_ - col * (cw_ + 0.24)
        y = top + 0.44 + row * (ch + 0.44)
        image_card(s, p_, x, y, cw_, ch, pad=0.04, mode='cover', shadow=True)
        add_text(s, x, y + ch + 0.10, cw_, 0.32,
                 [dict(t=cap, size=12, color=MUTED, align='center')])
    # takeaway panel
    panel = card(s, M, top + 0.44, 3.30, 3.86, fill=WHITE, radius=0.05,
                 line=BORDER)
    add_text(s, M + 0.28, top + 0.70, 2.75, 0.36,
             [dict(t='جمع‌بندی چالش‌ها', size=15, bold=True, color=NAVY,
                   align='right')])
    hline(s, M + 0.28, top + 1.12, 2.75, color=BORDER)
    for k, t in enumerate(['حریم خصوصی داده‌ها', 'اعتماد به مدل‌ها',
                           'شفافیت و تفسیرپذیری', 'قوانین و اخلاق']):
        y = top + 1.30 + k * 0.68
        shape(s, MSO_SHAPE.OVAL, M + 2.72, y + 0.13, 0.09, 0.09, fill=GOLD)
        add_text(s, M + 0.28, y, 2.35, 0.36,
                 [dict(t=t, size=13, color=INK, align='right')])
    footer(s, i)
    return s


# ============================================================== slide 22 ===
def s22(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'جمع‌بندی', kicker='چهار نکته کلیدی', num='22')
    pts = [('معکوس‌سازی چیست؟',
            'تلاش برای پیدا کردن ورودی از روی خروجی شبکه.'),
           ('دو رویکرد اصلی',
            'معکوس کردن مدل‌های معمولی (تقریبی) و طراحی شبکه‌های ذاتاً معکوس‌پذیر (دقیق).'),
           ('یک موضوع دووجهی',
            'هم تهدید جدی برای حریم خصوصی است و هم کاربردهای علمی و عملی مهمی دارد.'),
           ('چرا مهم است؟',
            'درک معکوس‌سازی کمک می‌کند مدل‌ها را بهتر بفهمیم و ایمن‌تر طراحی کنیم.')]
    cw_, ch = 5.80, 1.55
    for k, (t, d_) in enumerate(pts):
        col, row = k % 2, k // 2
        x = SW - M - cw_ - col * (cw_ + 0.30)
        y = top + 0.10 + row * (ch + 0.26)
        c = card(s, x, y, cw_, ch, fill=WHITE, radius=0.06, line=BORDER)
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + cw_ - 0.055, y + 0.22, 0.055,
              ch - 0.44, fill=[BLUE, TEAL, GOLD, NAVY][k], radius=0.5)
        add_text(s, x + 0.28, y + 0.22, cw_ - 0.62, 0.34,
                 [dict(t=t, size=17, bold=True, color=NAVY, align='right')])
        add_text(s, x + 0.28, y + 0.62, cw_ - 0.62, ch - 0.78,
                 [dict(t=d_, size=14, color=INK, align='right',
                       line_spacing=1.45)])
        anim(s, [c.shape_id])
    add_text(s, M, top + 3.72, CW, 0.40,
             [dict(t='معکوس‌سازی تنها یک حمله نیست؛ ابزاری برای فهم و طراحی بهتر شبکه‌هاست.',
                   size=15, bold=True, color=BLUE, align='center')])
    footer(s, i)
    return s


# ============================================================== slide 23 ===
def s23(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'اصطلاحات کلیدی', kicker='واژه‌نامه سریع', num='23')
    terms = [('Weight', 'وزن'), ('Bias', 'بایاس'),
             ('Activation Function', 'تابع فعال‌سازی'),
             ('Forward Pass', 'عبور رو به جلو'), ('Inversion', 'معکوس‌سازی'),
             ('Latent Space', 'فضای نهان'),
             ('Model Inversion Attack', 'حمله معکوس‌سازی مدل'),
             ('Invertible Neural Network (INN)', 'شبکه عصبی معکوس‌پذیر')]
    cw_, ch = 5.80, 0.94
    for k, (en, faf) in enumerate(terms):
        col, row = k % 2, k // 2
        x = SW - M - cw_ - col * (cw_ + 0.30)
        y = top + 0.16 + row * (ch + 0.20)
        c = card(s, x, y, cw_, ch, fill=WHITE, radius=0.08, line=BORDER)
        add_text(s, x + 0.30, y + 0.10, cw_ - 0.60, 0.40,
                 [dict(t=faf, size=17, bold=True, color=NAVY, align='right')],
                 anchor='middle')
        add_text(s, x + 0.30, y + 0.52, cw_ - 0.60, 0.30,
                 [dict(t=en, size=11, color=BLUE_2, align='right', rtl=False)],
                 anchor='middle')
        anim(s, [c.shape_id])
    footer(s, i)
    return s


# ============================================================== slide 24 ===
def s24(prs, i):
    s = blank(prs)
    light_bg(s)
    top = header(s, 'منابع', kicker='مقالات و منابع مورد استفاده', num='24')
    refs = [('مقاله PMC — مروری بر معکوس‌سازی', 'pmc.ncbi.nlm.nih.gov/articles/PMC12133272'),
            ('سری آموزشی 3Blue1Brown — Neural Networks', 'youtube.com/playlist'),
            ('Analyzing Inverse Problems with Invertible Neural Networks', 'arxiv.org/abs/1808.04730'),
            ('Model Inversion Attacks that Exploit Confidence Information', 'doi.org/10.1145/2810103.2813677'),
            ('Image Inversion: A Survey from GANs to Diffusion and Beyond', 'arxiv.org/abs/2502.11974'),
            ('Inverse Problems in Invertible Neural Networks (Heidelberg)', 'hci.iwr.uni-heidelberg.de/vislearn'),
            ('Analyzing Inverse Problems with INNs (OpenReview)', 'openreview.net/forum'),
            ('Model Inversion Attacks (CMU)', 'cs.cmu.edu/~mfredrik/papers/fjr2015ccs.pdf')]
    cw_ = 5.80
    for k, (t, url) in enumerate(refs):
        col, row = k % 2, k // 2
        x = SW - M - cw_ - col * (cw_ + 0.30)
        y = top + 0.16 + row * 1.06
        c = card(s, x, y, cw_, 0.90, fill=WHITE, radius=0.08, line=BORDER)
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + cw_ - 0.055, y + 0.18, 0.055,
              0.54, fill=BLUE_2, radius=0.5)
        add_text(s, x + 0.26, y + 0.12, cw_ - 0.60, 0.36,
                 [dict(t=t, size=13, bold=True, color=NAVY, align='right')],
                 anchor='middle')
        add_text(s, x + 0.26, y + 0.50, cw_ - 0.60, 0.30,
                 [dict(t=url, size=11, color=BLUE_2, align='right', rtl=False)],
                 anchor='middle')
    footer(s, i)
    return s


# ============================================================== slide 25 ===
def s25(prs, i):
    s = blank(prs)
    set_bg(s, RGBColor(0xFF, 0xFF, 0xFF))
    add_pic(s, os.path.join(ART, 'cover_bg.png'), 0, 0, SW, SH, radius=None,
            border=None, shadow=False)
    # centered closing card
    card(s, 3.05, 1.85, 7.25, 3.85, fill=WHITE, radius=0.035, line=BORDER,
         shadow=True, shadow_alpha=18)
    add_pic(s, img('image25.jpg'), 3.20, 2.00, 6.95, 1.55, radius=0.03,
            border=None, shadow=False, crop=(0.10, 0.30, 0.10, 0.18))
    add_text(s, 3.30, 3.72, 6.75, 0.50,
             [dict(t='پرسش و پاسخ', size=34, bold=True, color=NAVY,
                   align='center')])
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, SW / 2 - 0.85, 4.32, 1.7, 0.045,
          fill=GOLD, radius=0.5)
    add_text(s, 3.30, 4.52, 6.75, 0.40,
             [dict(t='از توجه شما سپاسگزارم', size=17, color=MUTED,
                   align='center')])
    add_text(s, 3.30, 4.94, 6.75, 0.36,
             [dict(t='علیرضا خدادادی  ·  دانشگاه شهید باهنر شیراز', size=13,
                   color=BLUE, align='center')])
    return s

def auto_anim(slide):
    """fallback: fade in the content shapes (skip full-bleed art)"""
    ids = []
    for sh in slide.shapes:
        try:
            w = sh.width / 914400 if sh.width else 0
            h = sh.height / 914400 if sh.height else 0
            top = sh.top / 914400 if sh.top is not None else 0
        except Exception:
            continue
        if w > SW * 0.85 and h > SH * 0.85:
            continue
        if top > SH - 0.75:
            continue
        ids.append(sh.shape_id)
    return ids[:8]




def apply_theme(prs):
    """Make the deck's default theme match the new palette/typography, so any
    newly added slide or stray placeholder stays on-brand."""
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    import re
    master = prs.slide_masters[0]
    theme_part = master.part.part_related_by(RT.THEME)
    xml = theme_part.blob.decode('utf-8')

    colors = {'dk1': '16203A', 'dk2': '182C55', 'lt1': 'FFFFFF',
              'lt2': 'EEF3FB', 'accent1': '2159A8', 'accent2': '3E7BD6',
              'accent3': '108C92', 'accent4': 'D99A22', 'accent5': 'B04845',
              'accent6': '5A6A85', 'hlink': '2159A8', 'folHlink': '108C92'}
    for key, val in colors.items():
        xml = re.sub(r'(<a:%s>)(.*?)(</a:%s>)' % (key, key),
                     lambda m: '<a:%s><a:srgbClr val="%s"/></a:%s>' % (key, val, key),
                     xml, count=1, flags=re.S)
    # Persian-friendly fonts everywhere
    def fix_font(block, name):
        block = re.sub(r'<a:latin typeface="[^"]*"[^/]*/>',
                       '<a:latin typeface="%s"/>' % name, block)
        block = re.sub(r'<a:cs typeface="[^"]*"/>',
                       '<a:cs typeface="%s"/>' % name, block)
        return block
    xml = re.sub(r'<a:majorFont>.*?</a:majorFont>',
                 lambda m: fix_font(m.group(0), 'Tahoma'), xml, flags=re.S)
    xml = re.sub(r'<a:minorFont>.*?</a:minorFont>',
                 lambda m: fix_font(m.group(0), 'Tahoma'), xml, flags=re.S)
    theme_part._blob = xml.encode('utf-8')

    cp = prs.core_properties
    cp.title = 'معکوس‌سازی شبکه‌های عصبی'
    cp.subject = 'ارائه دانشگاهی — از شبکه‌های عصبی تا معکوس کردن آن‌ها'
    cp.author = 'علیرضا خدادادی'
    cp.keywords = 'neural networks, inversion, INN, privacy, presentation'
    cp.comments = ('نسخه بازطراحی‌شده: تم روشن، تایپوگرافی فارسی، کارت‌ها و سایه‌های '
                   'یکدست، انیمیشن و ترنزیشن ساده و یکسان.')


def main():
    prs = new_deck()
    builders = [s01, s02, s03, s04, s05, s06, s07, s08, s09, s10,
                s11, s12, s13, s14, s15, s16, s17, s18, s19, s20, s21,
                s22, s23, s24, s25]
    for n, b in enumerate(builders, 1):
        b(prs, n)
    for sl in prs.slides:
        add_transition(sl)
    for sid, ids in SLIDES_ANIM.items():
        add_fade_anim(ANIM_SLIDES[sid], ids)
    for sl in prs.slides:
        if id(sl) not in SLIDES_ANIM:
            add_fade_anim(sl, auto_anim(sl))
    apply_theme(prs)
    prs.save(OUT)
    print('saved', OUT, len(prs.slides._sldIdLst))


if __name__ == '__main__':
    main()
