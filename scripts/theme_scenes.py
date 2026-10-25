# Пять тематических иллюстраций-образцов для слайда номинации («вариант 3», одобрен оператором).
# Новые темы Claude пишет так же: функция scene(rnd) -> (defs, элементы SVG) на холсте 1920×1080,
# рисунок справа (x > ~1100), слева прозрачно — там заголовок
import math

from scripts.theme_art import GLOW, H, SOFT, W, bg_glow, sparkle, star_path

# ---------- новогодняя: гирлянды, большая снежинка, снег ----------
def newyear(rnd):
    defs, out = [GLOW.format(s=9), SOFT.format(s=3)], []
    g, r = bg_glow("#3FA9FF", 0.82, 0.5, 0.55, 0.5)
    defs.append(g); out.append(r)
    # большая снежинка
    cx, cy, R = 1620, 600, 300
    arms = []
    for k in range(6):
        a = math.radians(k * 60 - 90)
        ex, ey = cx + R * math.cos(a), cy + R * math.sin(a)
        arms.append(f'<line x1="{cx}" y1="{cy}" x2="{ex:.0f}" y2="{ey:.0f}"/>')
        for t, L in ((0.38, 95), (0.62, 80), (0.83, 55)):
            bx, by = cx + R * t * math.cos(a), cy + R * t * math.sin(a)
            for s in (-1, 1):
                b = a + s * math.radians(42)
                arms.append(f'<line x1="{bx:.0f}" y1="{by:.0f}" x2="{bx + L * math.cos(b):.0f}" y2="{by + L * math.sin(b):.0f}"/>')
        arms.append(f'<circle cx="{ex:.0f}" cy="{ey:.0f}" r="13"/>')
    out.append(f'<g stroke="#EAF6FF" stroke-width="15" stroke-linecap="round" fill="#EAF6FF" filter="url(#glow)">{"".join(arms)}</g>')
    out.append(f'<circle cx="{cx}" cy="{cy}" r="46" fill="none" stroke="#EAF6FF" stroke-width="13" filter="url(#glow)"/>')
    # гирлянды с лампочками
    colors = ["#FF4D5E", "#FFD447", "#3DFF8A", "#46B9FF", "#FF7BE0"]
    for (x0, y0), (x1, y1), sag in (((1000, -30), (1950, 60), 230), ((1180, 1140), (1960, 720), 210)):
        mx, my = (x0 + x1) / 2, (y0 + y1) / 2 + sag
        out.append(f'<path d="M{x0} {y0} Q{mx} {my} {x1} {y1}" stroke="#1d2a33" stroke-width="5" fill="none"/>')
        for i in range(1, 15):
            t = i / 15
            px = (1 - t) ** 2 * x0 + 2 * (1 - t) * t * mx + t ** 2 * x1
            py = (1 - t) ** 2 * y0 + 2 * (1 - t) * t * my + t ** 2 * y1
            c = colors[i % len(colors)]
            out.append(f'<g transform="translate({px:.0f} {py:.0f}) rotate({rnd.randint(-20, 20)})">'
                       f'<rect x="-8" y="-4" width="16" height="16" rx="3" fill="#2b3a44"/>'
                       f'<ellipse cx="0" cy="30" rx="15" ry="24" fill="{c}" filter="url(#glow)"/>'
                       f'<ellipse cx="-5" cy="22" rx="4" ry="8" fill="#fff" opacity="0.7"/></g>')
    # снег
    for _ in range(140):
        x, y, rr = rnd.randint(1050, 1920), rnd.randint(0, 1080), rnd.choice((2, 3, 3, 4, 6))
        out.append(f'<circle cx="{x}" cy="{y}" r="{rr}" fill="#fff" opacity="{rnd.uniform(0.35, 0.9):.2f}"/>')
    for _ in range(9):
        out.append(sparkle(rnd.randint(1150, 1900), rnd.randint(150, 1000), rnd.randint(14, 26), "#fff", 0.9))
    return defs, out


# ---------- готический бал: луна, замок, летучие мыши, туман ----------
BAT = ("M0 0 C-12 -10 -30 -14 -46 -6 C-38 -4 -34 2 -36 10 C-28 4 -20 6 -16 12 C-12 6 -6 4 0 8 "
       "C6 4 12 6 16 12 C20 6 28 4 36 10 C34 2 38 -4 46 -6 C30 -14 12 -10 0 0 Z")


def gothic(rnd):
    defs, out = [GLOW.format(s=30), SOFT.format(s=28)], []
    g, r = bg_glow("#8A3DFF", 0.8, 0.4, 0.6, 0.55)
    defs.append(g); out.append(r)
    defs.append('<radialGradient id="moon" cx="0.4" cy="0.35" r="0.7"><stop offset="0" stop-color="#FFFFFF"/>'
                '<stop offset="0.6" stop-color="#E9E1FF"/><stop offset="1" stop-color="#B7A2EE"/></radialGradient>')
    mx, my, mr = 1560, 400, 280
    out.append(f'<circle cx="{mx}" cy="{my}" r="{mr}" fill="#CDBBFF" opacity="0.55" filter="url(#glow)"/>')
    out.append(f'<circle cx="{mx}" cy="{my}" r="{mr}" fill="url(#moon)"/>')
    for cx, cy, cr in ((-90, -60, 48), (70, 40, 34), (-20, 110, 26), (110, -100, 22), (-130, 80, 18)):
        out.append(f'<circle cx="{mx + cx}" cy="{my + cy}" r="{cr}" fill="#A994E0" opacity="0.35"/>')
    # замок
    base = 1080
    towers = [(1180, 640, 70), (1290, 560, 90), (1420, 470, 110), (1560, 600, 80), (1680, 520, 100), (1820, 650, 80)]
    parts = [f'<rect x="1150" y="760" width="800" height="{base - 760}"/>']
    for x, top, w in towers:
        parts.append(f'<rect x="{x - w / 2}" y="{top}" width="{w}" height="{base - top}"/>')
        parts.append(f'<path d="M{x - w / 2 - 12} {top} L{x} {top - w * 1.25} L{x + w / 2 + 12} {top} Z"/>')
        parts.append(f'<rect x="{x - 2}" y="{top - w * 1.25 - 40}" width="4" height="40"/>')
    for x in range(1150, 1950, 46):
        parts.append(f'<rect x="{x}" y="740" width="26" height="24"/>')
    out.append(f'<g fill="#0A0612">{"".join(parts)}</g>')
    for x, top, w in towers:
        for k in range(2):
            wy = top + 60 + k * 110
            if wy < base - 60 and rnd.random() < 0.8:
                out.append(f'<path d="M{x - 12} {wy + 40} V{wy + 12} Q{x} {wy - 6} {x + 12} {wy + 12} V{wy + 40} Z" '
                           f'fill="#FFC857" filter="url(#glow)" opacity="0.95"/>')
    # туман
    for _ in range(7):
        out.append(f'<ellipse cx="{rnd.randint(1100, 1950)}" cy="{rnd.randint(960, 1080)}" rx="{rnd.randint(200, 360)}" '
                   f'ry="{rnd.randint(40, 70)}" fill="#B9A3FF" opacity="0.18" filter="url(#soft)"/>')
    # летучие мыши
    for x, y, s in ((1330, 240, 1.6), (1700, 210, 1.1), (1780, 330, 0.8), (1880, 300, 0.9), (1620, 150, 0.7),
                    (1450, 330, 0.6), (1880, 120, 0.9)):
        out.append(f'<path d="{BAT}" fill="#0A0612" transform="translate({x} {y}) scale({s * 1.6}) rotate({rnd.randint(-15, 15)})"/>')
    return defs, out


# ---------- Гэтсби: золотой веер ар-деко и бокал шампанского ----------
def gatsby(rnd):
    defs, out = [GLOW.format(s=6), SOFT.format(s=4)], []
    g, r = bg_glow("#E8B64C", 0.82, 0.75, 0.6, 0.4)
    defs.append(g); out.append(r)
    defs.append('<linearGradient id="gold" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#FFF1B8"/>'
                '<stop offset="0.45" stop-color="#E8B64C"/><stop offset="1" stop-color="#9C6B1F"/></linearGradient>')
    cx, cy = 1690, 1080
    rays = []
    for i in range(37):
        a = math.radians(180 + i * 5)
        rays.append(f'<line x1="{cx + 150 * math.cos(a):.0f}" y1="{cy + 150 * math.sin(a):.0f}" '
                    f'x2="{cx + 600 * math.cos(a):.0f}" y2="{cy + 600 * math.sin(a):.0f}" stroke-width="{3 if i % 2 else 6}"/>')
    out.append(f'<g stroke="url(#gold)" opacity="0.85">{"".join(rays)}</g>')
    for rr, sw in ((170, 8), (300, 4), (430, 10), (465, 3), (580, 5)):
        out.append(f'<path d="M{cx - rr} {cy} A{rr} {rr} 0 0 1 {cx + rr} {cy}" fill="none" stroke="url(#gold)" stroke-width="{sw}"/>')
    # ступенчатые «лучи» ар-деко по краям
    for k in range(5):
        y = 120 + k * 34
        out.append(f'<path d="M{1330 + k * 30} {y} H{1910 - k * 30}" stroke="url(#gold)" stroke-width="4"/>')
    out.append(f'<path d="{star_path(1690, 120, 34, 12, 4, -90)}" fill="url(#gold)" filter="url(#glow)"/>')
    # бокал-шампанка
    gx, gy = 1690, 560
    glass = (f'<path d="M{gx - 150} {gy} Q{gx} {gy + 170} {gx + 150} {gy} Z" fill="url(#gold)" opacity="0.92"/>'
             f'<path d="M{gx - 150} {gy} Q{gx} {gy + 170} {gx + 150} {gy}" fill="none" stroke="#FFF1B8" stroke-width="6"/>'
             f'<rect x="{gx - 6}" y="{gy + 84}" width="12" height="190" fill="url(#gold)"/>'
             f'<ellipse cx="{gx}" cy="{gy + 280}" rx="95" ry="18" fill="url(#gold)"/>')
    out.append(glass)
    for _ in range(26):
        bx, by = gx + rnd.randint(-110, 110), gy - rnd.randint(10, 380)
        out.append(f'<circle cx="{bx}" cy="{by}" r="{rnd.randint(4, 12)}" fill="none" stroke="#FFF1B8" stroke-width="3" '
                   f'opacity="{rnd.uniform(0.5, 1):.2f}"/>')
    # жемчужные нити с подвесками ар-деко по верху (кроме столбика пузырьков над бокалом)
    for x, length in ((1270, 190), (1335, 300), (1400, 230), (1465, 360), (1530, 260), (1850, 330), (1905, 210)):
        beads = "".join(f'<circle cx="{x}" cy="{y}" r="7"/>' for y in range(8, length, 22))
        out.append(f'<g fill="url(#gold)">{beads}</g>')
        out.append(f'<path d="M{x} {length - 4} L{x + 16} {length + 22} L{x} {length + 56} L{x - 16} {length + 22} Z" '
                   f'fill="url(#gold)" filter="url(#glow)"/>')
    out.append('<path d="M1240 6 H1940" stroke="url(#gold)" stroke-width="6"/>')
    for _ in range(10):
        out.append(sparkle(rnd.randint(1200, 1900), rnd.randint(200, 900), rnd.randint(12, 24), "#FFF1B8", 0.95))
    return defs, out


# ---------- 90-е: диско-шар, лучи, неоновый пол ----------
def disco90(rnd):
    defs, out = [GLOW.format(s=10), SOFT.format(s=14)], []
    g, r = bg_glow("#FF3DB8", 0.8, 0.45, 0.6, 0.5)
    defs.append(g); out.append(r)
    bx, by, br = 1560, 380, 250
    for k, c in enumerate(("#FF3DB8", "#3DE8FF", "#FFE14D", "#9D5CFF", "#3DE8FF", "#FF3DB8")):
        a = math.radians(40 + k * 20)
        x2, y2 = bx + 1100 * math.cos(a) * (-1 if k % 2 else 1), by + 1100 * math.sin(a)
        out.append(f'<path d="M{bx} {by} L{x2 - 90:.0f} {y2:.0f} L{x2 + 90:.0f} {y2:.0f} Z" fill="{c}" opacity="0.16" filter="url(#soft)"/>')
    out.append(f'<line x1="{bx}" y1="0" x2="{bx}" y2="{by - br}" stroke="#ccc" stroke-width="5"/>')
    defs.append(f'<clipPath id="ball"><circle cx="{bx}" cy="{by}" r="{br}"/></clipPath>')
    tiles = []
    rows, cols = 14, 16
    for i in range(rows):
        la0, la1 = -math.pi / 2 + i * math.pi / rows, -math.pi / 2 + (i + 1) * math.pi / rows
        for j in range(cols):
            lo0, lo1 = -math.pi / 2 + j * math.pi / cols, -math.pi / 2 + (j + 1) * math.pi / cols
            pts = []
            for la, lo in ((la0, lo0), (la0, lo1), (la1, lo1), (la1, lo0)):
                pts.append(f"{bx + br * math.cos(la) * math.sin(lo):.1f},{by + br * math.sin(la):.1f}")
            mid_la, mid_lo = (la0 + la1) / 2, (lo0 + lo1) / 2
            light = max(0.0, math.cos(mid_la + 0.5) * math.cos(mid_lo + 0.45))
            base = rnd.choice(("#F4F4FF", "#D8D8E8", "#A8A8C0", "#FF8AD6", "#8AF2FF", "#FFFFFF", "#D8D8E8"))
            tiles.append(f'<polygon points="{" ".join(pts)}" fill="{base}" stroke="#1d1830" stroke-width="2" '
                         f'opacity="{0.35 + 0.65 * light:.2f}"/>')
    defs.append('<radialGradient id="shade" cx="0.38" cy="0.35" r="0.7"><stop offset="0.55" stop-color="#000" stop-opacity="0"/>'
                '<stop offset="1" stop-color="#14051f" stop-opacity="0.75"/></radialGradient>')
    out.append(f'<circle cx="{bx}" cy="{by}" r="{br}" fill="#2a2440"/>')
    out.append(f'<g clip-path="url(#ball)">{"".join(tiles)}</g>')
    out.append(f'<circle cx="{bx}" cy="{by}" r="{br}" fill="url(#shade)"/>')
    out.append(f'<circle cx="{bx}" cy="{by}" r="{br}" fill="none" stroke="#FF8AD6" stroke-width="4" filter="url(#glow)"/>')
    for x, y, s in ((1410, 250, 34), (1660, 300, 24), (1500, 470, 20), (1760, 170, 30), (1290, 140, 22), (1850, 520, 26)):
        out.append(sparkle(x, y, s, "#FFFFFF"))
    # неоновый пол: ровная сетка в перспективе, обрезана одной областью и гаснет к левому краю
    hx, hy, x_left, vy = 1560, 800, 0, 430  # линии сходятся выше горизонта — ровный веер без яркого пятна
    defs.append(f'<clipPath id="floor"><rect x="{x_left}" y="{hy}" width="{W - x_left}" height="{H - hy}"/></clipPath>'
                '<linearGradient id="ffade" x1="0" y1="0" x2="1" y2="0"><stop offset="0.3" stop-color="#fff" stop-opacity="0"/>'
                '<stop offset="0.62" stop-color="#fff" stop-opacity="1"/></linearGradient>'
                f'<mask id="floorfade" maskUnits="userSpaceOnUse" x="0" y="0" width="{W}" height="{H}">'
                f'<rect width="{W}" height="{H}" fill="url(#ffade)"/></mask>')
    floor = []
    for k in range(1, 10):
        y = hy + (H - hy) * (k / 9) ** 1.8
        floor.append(f'<line x1="0" y1="{y:.0f}" x2="{W}" y2="{y:.0f}"/>')
    for k in range(-12, 13):
        floor.append(f'<line x1="{hx}" y1="{vy}" x2="{hx + k * 210}" y2="{H}"/>')
    out.append(f'<g mask="url(#floorfade)"><g clip-path="url(#floor)" stroke="#FF3DB8" stroke-width="3" filter="url(#glow)" '
               f'opacity="0.9">{"".join(floor)}</g>'
               f'<line x1="{x_left}" y1="{hy}" x2="{W}" y2="{hy}" stroke="#3DE8FF" stroke-width="5" filter="url(#glow)"/></g>')
    return defs, out


# ---------- СССР: конструктивистская звезда с лучами и ракета ----------
def ussr(rnd):
    defs, out = [GLOW.format(s=8), SOFT.format(s=6)], []
    g, r = bg_glow("#E5322D", 0.8, 0.5, 0.6, 0.45)
    defs.append(g); out.append(r)
    cx, cy = 1560, 520
    defs.append('<linearGradient id="fade" x1="0" y1="0" x2="1" y2="0"><stop offset="0.35" stop-color="#fff" stop-opacity="0"/>'
                '<stop offset="0.7" stop-color="#fff" stop-opacity="1"/></linearGradient>'
                f'<mask id="raysmask" maskUnits="userSpaceOnUse" x="0" y="0" width="{W}" height="{H}">'
                f'<rect width="{W}" height="{H}" fill="url(#fade)"/></mask>')
    rays = []
    for i in range(18):
        a0, a1 = math.radians(i * 20), math.radians(i * 20 + 10)
        rays.append(f'<path d="M{cx} {cy} L{cx + 1300 * math.cos(a0):.0f} {cy + 1300 * math.sin(a0):.0f} '
                   f'L{cx + 1300 * math.cos(a1):.0f} {cy + 1300 * math.sin(a1):.0f} Z" fill="#E5322D" opacity="0.25"/>')
    out.append(f'<g mask="url(#raysmask)">{"".join(rays)}</g>')
    # диагональные плашки
    out.append('<path d="M1100 1080 L1500 680 L1600 680 L1200 1080 Z" fill="#F2E6D0" opacity="0.9"/>')
    out.append('<path d="M1700 0 L1920 0 L1920 90 L1610 90 Z" fill="#F2E6D0" opacity="0.9"/>')
    out.append(f'<circle cx="{cx}" cy="{cy}" r="330" fill="none" stroke="#F2E6D0" stroke-width="10" stroke-dasharray="4 26"/>')
    out.append(f'<path d="{star_path(cx, cy, 300, 120)}" fill="#E5322D" filter="url(#glow)"/>')
    out.append(f'<path d="{star_path(cx, cy, 300, 120)}" fill="none" stroke="#F2E6D0" stroke-width="10" stroke-linejoin="round"/>')
    out.append(f'<path d="{star_path(cx, cy, 150, 60)}" fill="#B31F1B"/>')
    # ракета с хвостом
    rx, ry, ang = 1330, 250, 35
    rocket = ('<path d="M0 -70 C26 -40 30 10 22 50 L-22 50 C-30 10 -26 -40 0 -70 Z" fill="#F2E6D0"/>'
              '<circle cx="0" cy="-12" r="11" fill="#2E8FE5" stroke="#B31F1B" stroke-width="4"/>'
              '<path d="M-22 26 L-42 62 L-20 52 Z M22 26 L42 62 L20 52 Z" fill="#E5322D"/>'
              '<path d="M-14 52 Q0 120 14 52 Z" fill="#FFC857" filter="url(#glow)"/>')
    out.append(f'<path d="M{rx - 120} {ry + 190} Q{rx - 60} {ry + 110} {rx} {ry}" stroke="#F2E6D0" stroke-width="8" '
               f'fill="none" stroke-dasharray="18 14" opacity="0.8"/>')
    out.append(f'<g transform="translate({rx} {ry}) rotate({ang}) scale(1.6)">{rocket}</g>')
    for x, y, s in ((1820, 220, 26), (1880, 760, 20), (1260, 560, 16), (1720, 900, 22), (1400, 120, 14)):
        out.append(f'<path d="{star_path(x, y, s, s * 0.42)}" fill="#F2E6D0"/>')
    return defs, out


# ---------- Гэтсби, вариант «шампанская башня»: пирамида бокалов, жемчуг, золотой веер ----------
def gatsby_tower(rnd):
    defs, out = [GLOW.format(s=7), SOFT.format(s=10)], []
    g, r = bg_glow("#F2C46D", 0.8, 0.7, 0.6, 0.45)
    defs.append(g); out.append(r)
    defs.append('<linearGradient id="gold" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#FFF4C9"/>'
                '<stop offset="0.5" stop-color="#F2C46D"/><stop offset="1" stop-color="#A9771F"/></linearGradient>'
                '<linearGradient id="fizz" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#FFF6D6" stop-opacity="0.95"/>'
                '<stop offset="1" stop-color="#E9B44C" stop-opacity="0.85"/></linearGradient>')
    cx, base = 1580, 1040
    # веер-лучи ар-деко за башней
    rays = []
    for i in range(25):
        a = math.radians(195 + i * 6.25)
        rays.append(f'<line x1="{cx}" y1="{base}" x2="{cx + 760 * math.cos(a):.0f}" y2="{base + 760 * math.sin(a):.0f}" '
                    f'stroke-width="{2 if i % 2 else 5}"/>')
    out.append(f'<g stroke="url(#gold)" opacity="0.35">{"".join(rays)}</g>')
    for rr in (520, 640):
        out.append(f'<path d="M{cx - rr} {base} A{rr} {rr} 0 0 1 {cx + rr} {base}" fill="none" stroke="url(#gold)" '
                   f'stroke-width="3" opacity="0.5"/>')
    # пирамида бокалов-шампанок: 4 ряда
    gw, gh = 118, 150
    glasses = []
    for row, count in enumerate((4, 3, 2, 1)):
        y = base - 30 - row * gh
        for k in range(count):
            x = cx + (k - (count - 1) / 2) * (gw + 14)
            glasses.append(
                f'<path d="M{x - gw / 2} {y - gh + 40} Q{x} {y - gh + 120} {x + gw / 2} {y - gh + 40} Z" fill="url(#fizz)"/>'
                f'<path d="M{x - gw / 2} {y - gh + 40} Q{x} {y - gh + 120} {x + gw / 2} {y - gh + 40}" fill="none" '
                f'stroke="#FFF4C9" stroke-width="4"/>'
                f'<rect x="{x - 4}" y="{y - gh + 96}" width="8" height="{gh - 110}" fill="url(#gold)"/>'
                f'<ellipse cx="{x}" cy="{y - 12}" rx="38" ry="8" fill="url(#gold)"/>')
    out.append(f'<g filter="url(#glow)">{"".join(glasses)}</g>')
    # струя из бутылки в верхний бокал
    top = base - 30 - 3 * gh - gh + 40
    out.append(f'<path d="M{cx + 170} {top - 230} Q{cx + 40} {top - 170} {cx + 6} {top + 10}" stroke="#FFF1B8" '
               f'stroke-width="9" fill="none" stroke-linecap="round" filter="url(#glow)" opacity="0.9"/>')
    out.append(f'<g transform="translate({cx + 205} {top - 265}) rotate(-125)">'
               '<rect x="-34" y="-20" width="68" height="190" rx="26" fill="#1F3B2C" stroke="url(#gold)" stroke-width="4"/>'
               '<rect x="-14" y="-90" width="28" height="80" rx="8" fill="#1F3B2C" stroke="url(#gold)" stroke-width="4"/>'
               '<rect x="-34" y="40" width="68" height="60" fill="url(#gold)"/></g>')
    # жемчужные нити дугами по верху
    for x0, x1, sag in ((1180, 1660, 150), (1560, 1940, 120), (1300, 1900, 260)):
        mx = (x0 + x1) / 2
        beads = []
        for i in range(1, 30):
            t = i / 30
            px = (1 - t) ** 2 * x0 + 2 * (1 - t) * t * mx + t ** 2 * x1
            py = 2 * (1 - t) * t * (sag * 2) - 10
            beads.append(f'<circle cx="{px:.0f}" cy="{py:.0f}" r="8"/>')
        out.append(f'<g fill="#FFF7E6" opacity="0.92">{"".join(beads)}</g>')
    for _ in range(40):
        bx, by = cx + rnd.randint(-260, 260), rnd.randint(260, 1000)
        out.append(f'<circle cx="{bx}" cy="{by}" r="{rnd.randint(3, 9)}" fill="none" stroke="#FFF4C9" stroke-width="2" '
                   f'opacity="{rnd.uniform(0.4, 0.9):.2f}"/>')
    for _ in range(12):
        out.append(sparkle(rnd.randint(1200, 1900), rnd.randint(150, 950), rnd.randint(12, 26), "#FFF4C9", 0.95))
    return defs, out


# ---------- Гэтсби, вариант «изумрудный Нью-Йорк»: небоскрёбы ар-деко, салют, зелёный огонёк ----------
def gatsby_skyline(rnd):
    defs, out = [GLOW.format(s=10), SOFT.format(s=16)], []
    g, r = bg_glow("#2BD9A0", 0.8, 0.55, 0.62, 0.42)
    defs.append(g); out.append(r)
    defs.append('<linearGradient id="emer" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7DFFD2"/>'
                '<stop offset="1" stop-color="#127A5C"/></linearGradient>')
    # салют
    for fx, fy, fr, c in ((1450, 230, 150, "#7DFFD2"), (1780, 160, 120, "#F2C46D"), (1650, 360, 90, "#FFFFFF")):
        lines = []
        for k in range(28):
            a = math.radians(k * 360 / 28)
            lines.append(f'<line x1="{fx + fr * 0.25 * math.cos(a):.0f}" y1="{fy + fr * 0.25 * math.sin(a):.0f}" '
                         f'x2="{fx + fr * math.cos(a):.0f}" y2="{fy + fr * math.sin(a):.0f}"/>')
        out.append(f'<g stroke="{c}" stroke-width="4" stroke-linecap="round" filter="url(#glow)" opacity="0.9">{"".join(lines)}</g>')
    # небоскрёбы ар-деко
    horizon = 860
    towers = [(1190, 520, 90), (1290, 380, 120), (1430, 250, 110), (1560, 430, 100), (1680, 300, 130), (1830, 470, 110)]
    parts, windows = [], []
    for x, top, w in towers:
        parts.append(f'<rect x="{x - w / 2}" y="{top}" width="{w}" height="{horizon - top}"/>')
        for step in range(3):  # ступенчатая верхушка
            sw, sh = w * (0.75 - step * 0.22), 34
            parts.append(f'<rect x="{x - sw / 2}" y="{top - sh * (step + 1)}" width="{sw}" height="{sh}"/>')
        parts.append(f'<path d="M{x - 6} {top - 102} L{x} {top - 190} L{x + 6} {top - 102} Z"/>')
        for wy in range(top + 24, horizon - 20, 34):
            for wx in range(int(x - w / 2 + 14), int(x + w / 2 - 14), 22):
                if rnd.random() < 0.45:
                    windows.append(f'<rect x="{wx}" y="{wy}" width="9" height="14"/>')
    # шпиль в стиле Крайслер-билдинг
    cxs, ct = 1430, 250
    for k in range(5):
        rr = 60 - k * 11
        parts.append(f'<path d="M{cxs - rr} {ct - 102 - k * 26} A{rr} {rr} 0 0 1 {cxs + rr} {ct - 102 - k * 26}" '
                     f'stroke="#F2C46D" stroke-width="5" fill="none"/>')
    out.append(f'<g fill="#07130F">{"".join(parts)}</g>')
    out.append(f'<g fill="#F2C46D" opacity="0.85">{"".join(windows)}</g>')
    # вода, отражение и зелёный огонёк на причале
    out.append(f'<rect x="1100" y="{horizon}" width="{W - 1100}" height="{H - horizon}" fill="#062019" opacity="0.85"/>')
    for k in range(9):
        y = horizon + 18 + k * 22
        out.append(f'<line x1="{1150 + rnd.randint(0, 120)}" y1="{y}" x2="{1900 - rnd.randint(0, 160)}" y2="{y}" '
                   f'stroke="#2BD9A0" stroke-width="2" opacity="{0.35 - k * 0.03:.2f}"/>')
    dx, dy = 1760, 960
    out.append(f'<rect x="{dx - 6}" y="{dy - 70}" width="12" height="90" fill="#0C2A21"/>')
    out.append(f'<circle cx="{dx}" cy="{dy - 78}" r="60" fill="#2BD9A0" opacity="0.35" filter="url(#soft)"/>')
    out.append(f'<circle cx="{dx}" cy="{dy - 78}" r="16" fill="#AFFFE6" filter="url(#glow)"/>')
    out.append(f'<path d="M{dx - 12} {dy - 60} L{dx - 300} {H} L{dx + 300} {H} L{dx + 12} {dy - 60} Z" fill="url(#emer)" '
               f'opacity="0.12" filter="url(#soft)"/>')
    for _ in range(80):
        out.append(f'<circle cx="{rnd.randint(1150, 1910)}" cy="{rnd.randint(20, 600)}" r="{rnd.choice((1, 2, 2, 3))}" '
                   f'fill="#FFFFFF" opacity="{rnd.uniform(0.3, 0.9):.2f}"/>')
    for _ in range(7):
        out.append(sparkle(rnd.randint(1200, 1900), rnd.randint(80, 650), rnd.randint(10, 20), "#AFFFE6", 0.9))
    return defs, out


THEMES = {
    "newyear": {"name": "Новогодняя", "title": "НАРЯДНЫЕ,КАК ЁЛКИ", "glow": "#3FA9FF", "scene": newyear},
    "gothic": {"name": "Готический бал", "title": "ТЬМА ВАМК ЛИЦУ", "glow": "#8A3DFF", "scene": gothic},
    "gatsby": {"name": "Гэтсби", "title": "БЛЕСКПО‑ГЭТСБИ", "glow": "#E8B64C", "scene": gatsby},
    "disco90": {"name": "Дискотека 90-х", "title": "ПРИКИДЧТО НАДО", "glow": "#FF3DB8", "scene": disco90},
    "ussr": {"name": "Назад в СССР", "title": "ОДЕТЫПО ГОСТУ", "glow": "#E5322D", "scene": ussr},
}
