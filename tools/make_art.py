"""Generate the decorative art used by cover + section slides (vector-like,
deterministic, drawn with PIL so it stays crisp and on-palette)."""
import math
import random
import os
from PIL import Image, ImageDraw, ImageFilter

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(HERE, 'assets')

BLUE = (33, 89, 168)
BLUE2 = (62, 123, 214)
TEAL = (16, 140, 146)
GOLD = (217, 154, 34)


def lerp(c1, c2, t):
    return tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))


def net_art(w, h, seed=7, density=1.0, spread=(0.0, 1.0)):
    """Sparse, elegant network motif: free-floating nodes, long thin edges."""
    random.seed(seed)
    S = 2
    W, H = w * S, h * S
    img = Image.new('RGBA', (W, H), (255, 255, 255, 0))
    d = ImageDraw.Draw(img)
    target = int(120 * density)
    nodes = []
    tries = 0
    x0, x1 = spread
    while len(nodes) < target and tries < target * 60:
        tries += 1
        fx = random.uniform(x0, x1)
        fy = random.random() ** 0.92
        # gentle bias toward the left/lower area for depth
        x = fx * W
        y = fy * H
        if any(abs(x - n[0]) < 0.045 * W and abs(y - n[1]) < 0.05 * H
               for n in nodes):
            continue
        nodes.append((x, y, fx, fy))
    lim = 0.20 * max(W, H)
    for i, (x1_, y1_, _, _) in enumerate(nodes):
        cands = []
        for j, (x2, y2, _, _) in enumerate(nodes):
            if j <= i:
                continue
            dist = math.hypot(x2 - x1_, y2 - y1_)
            if dist < lim:
                cands.append((dist, x2, y2))
        cands.sort()
        for dist, x2, y2 in cands[:3]:
            a = int(96 * (1 - dist / lim) ** 1.4 + 10)
            col = lerp(TEAL, BLUE2, random.random() * 0.9)
            d.line([x1_, y1_, x2, y2], fill=col + (a,), width=max(1, int(1.0 * S)))
    glow = Image.new('RGBA', (W, H), (255, 255, 255, 0))
    gd = ImageDraw.Draw(glow)
    for x, y, fx, fy in nodes:
        rad = (2.2 + 6.0 * (1 - fx)) * S
        gold = random.random() > 0.80
        c = GOLD if gold else lerp(BLUE, TEAL, random.random())
        gd.ellipse([x - rad * 2.4, y - rad * 2.4, x + rad * 2.4, y + rad * 2.4],
                   fill=c + (52,))
        gd.ellipse([x - rad, y - rad, x + rad, y + rad], fill=c + (215,))
    glow = glow.filter(ImageFilter.GaussianBlur(5 * S))
    img = Image.alpha_composite(img, glow)
    d = ImageDraw.Draw(img)
    for x, y, fx, fy in nodes:
        rad = (1.5 + 3.2 * (1 - fx)) * S
        d.ellipse([x - rad, y - rad, x + rad, y + rad], fill=(255, 255, 255, 235))
        d.ellipse([x - rad * .62, y - rad * .62, x + rad * .62, y + rad * .62],
                  fill=lerp(BLUE, BLUE2, .4) + (242,))
    return img.resize((w, h), Image.LANCZOS)


def cover_bg(w=2400, h=1350, alpha=0.55):
    """Full-bleed 16:9 background: light gradient + very faint network motif."""
    base = Image.new('RGBA', (w, h), (255, 255, 255, 255))
    d = ImageDraw.Draw(base)
    for y in range(h):
        t = y / h
        c = lerp((255, 255, 255), (229, 239, 251), t ** 1.25)
        d.line([(0, y), (w, y)], fill=c + (255,))
    for yy in range(36, h, 44):
        for xx in range(36, w, 44):
            d.ellipse([xx, yy, xx + 2, yy + 2], fill=(203, 219, 241, 120))
    motif = net_art(w, h, seed=23, density=1.15)
    fade = Image.new('L', (w, h), 0)
    fd = ImageDraw.Draw(fade)
    for x in range(w):
        fx = x / w
        fd.line([(x, 0), (x, h)], fill=int(255 * max(0.0, 1 - (fx / 0.60) ** 1.2)))
    a = Image.composite(fade, Image.new('L', (w, h), 0), motif.split()[3])
    a = a.point(lambda v: int(v * alpha))
    base.paste(motif, (0, 0), a)
    return base.convert('RGB').filter(ImageFilter.SMOOTH)


def panel_art(w=1200, h=1500):
    """Portrait panel art for the cover-of-details slide."""
    base = Image.new('RGBA', (w, h), (255, 255, 255, 255))
    d = ImageDraw.Draw(base)
    for y in range(h):
        t = y / h
        c = lerp((245, 250, 255), (222, 235, 250), t ** 1.1)
        d.line([(0, y), (w, y)], fill=c + (255,))
    for yy in range(30, h, 42):
        for xx in range(30, w, 42):
            d.ellipse([xx, yy, xx + 2, yy + 2], fill=(196, 214, 238, 120))
    motif = net_art(w, h, seed=41, density=0.9)
    base.alpha_composite(motif)
    return base.convert('RGB')


def cover_art(w=2400, h=1350):
    return cover_bg(w, h)


def band_art(w=1200, h=2000):
    base = Image.new('RGBA', (w, h), (255, 255, 255, 255))
    d = ImageDraw.Draw(base)
    for x in range(w):
        t = x / w
        c = lerp((198, 220, 246), (255, 255, 255), min(1, t ** 0.75 * 1.45))
        d.line([(x, 0), (x, h)], fill=c + (255,))
    for yy in range(30, h, 44):
        for xx in range(30, w, 44):
            d.ellipse([xx, yy, xx + 2, yy + 2], fill=(186, 208, 236, 130))
    art = net_art(w, h, seed=9, density=1.05)
    base.alpha_composite(art)
    return base.convert('RGB')


if __name__ == '__main__':
    cover_bg().save(os.path.join(OUT, 'cover_bg.png'))
    panel_art().save(os.path.join(OUT, 'panel_art.png'))
    band_art().save(os.path.join(OUT, 'section_art.png'))
    print('art written')
