# Иллюстрация тематического слайда номинации: сцена (SVG, рисует Claude кодом) → PNG 1920×1080
# с прозрачным фоном. Снимок делает браузер Edge без окна — без нейросетей и платных сервисов
import math
import os
import random
import subprocess
import tempfile

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
W, H = 1920, 1080

# Общие помощники для сцен: свечение, размытие, подсветка фона, звезда, искра
GLOW = """<filter id="glow" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="{s}" result="b"/>
<feMerge><feMergeNode in="b"/><feMergeNode in="b"/><feMergeNode in="SourceGraphic"/></feMerge></filter>"""
SOFT = '<filter id="soft" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="{s}"/></filter>'


def bg_glow(color, cx=0.85, cy=0.55, r=0.6, a=0.45):
    return (f'<radialGradient id="bgg" cx="{cx}" cy="{cy}" r="{r}"><stop offset="0" stop-color="{color}" stop-opacity="{a}"/>'
            f'<stop offset="1" stop-color="{color}" stop-opacity="0"/></radialGradient>', f'<rect width="{W}" height="{H}" fill="url(#bgg)"/>')


def star_path(cx, cy, r_out, r_in, n=5, rot=-90):
    pts = []
    for i in range(n * 2):
        r = r_out if i % 2 == 0 else r_in
        a = math.radians(rot + i * 180 / n)
        pts.append(f"{cx + r * math.cos(a):.1f},{cy + r * math.sin(a):.1f}")
    return "M" + " L".join(pts) + " Z"


def sparkle(cx, cy, r, color, op=1.0):
    return (f'<path d="M{cx} {cy - r} Q{cx} {cy} {cx + r} {cy} Q{cx} {cy} {cx} {cy + r} Q{cx} {cy} {cx - r} {cy} '
            f'Q{cx} {cy} {cx} {cy - r} Z" fill="{color}" opacity="{op}"/>')


def render_scene(scene, out_png: str, seed: str) -> str:
    # seed — чтобы одна и та же сцена (снег, пузырьки) каждый раз выходила одинаковой
    if not os.path.exists(EDGE):
        raise FileNotFoundError(f"Не найден браузер Edge ({EDGE}) — без него иллюстрацию не нарисовать")
    defs, body = scene(random.Random(seed))
    html = (f'<!doctype html><html><head><meta charset="utf-8"><style>html,body{{margin:0;background:transparent;'
            f'overflow:hidden}}</style></head><body><svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
            f'viewBox="0 0 {W} {H}"><defs>{"".join(defs)}</defs>{"".join(body)}</svg></body></html>')
    out_png = os.path.abspath(out_png)
    with tempfile.TemporaryDirectory() as tmp:
        page = os.path.join(tmp, "scene.html")
        with open(page, "w", encoding="utf-8") as f:
            f.write(html)
        subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                        "--default-background-color=00000000", f"--window-size={W},{H}",
                        f"--screenshot={out_png}", "file:///" + page.replace("\\", "/")],
                       check=True, capture_output=True, timeout=120)
    if not os.path.exists(out_png):
        raise RuntimeError(f"Edge не сохранил картинку {out_png}")
    return out_png
