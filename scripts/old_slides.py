# Узнавание песенных слайдов в старых играх — без PowerPoint, чтением .pptx как zip
import glob
import html
import os
import posixpath
import re
import zipfile
from collections import Counter

import openpyxl

from scripts.song_names import make_raw, norm, split_raw

AUDIO_EXT = (".mp3", ".m4a", ".wav", ".wma", ".aac")
VIDEO_EXT = (".mp4", ".mov", ".wmv", ".avi", ".m4v")
IMAGE_EXT = (".png", ".jpg", ".jpeg", ".jfif", ".gif", ".bmp", ".bin")  # .bin в старых играх — PNG
# Картинка, которая стоит на половине песенных слайдов игры и больше, — оформление, а не фото
DECOR_SHARE = 0.5
PLACEHOLDER_AREA = 10**18  # «площадь» картинки в рамке макета: больше любой картинки со своим размером
CHECK_UNKNOWN = "не узнана по своду — проверьте написание"
CHECK_AMBIGUOUS = "подходят несколько песен свода — проверьте"


def load_svod_songs(folder: str) -> list[dict]:
    songs = {}
    for path in sorted(glob.glob(os.path.join(folder, "*", "*СВОД*.xlsx"))):
        if os.path.basename(path).startswith("~$"):
            continue
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets[1:]:  # первый лист — «СОДЕРЖАНИЕ»
            for row in ws.iter_rows(values_only=True):
                if not row or not isinstance(row[0], int) or not row[1]:
                    continue
                if len(row) > 3 and row[2] and row[3]:
                    artist, title = str(row[2]), str(row[3])
                else:
                    artist, title = split_raw(str(row[1]))
                a, t = norm(artist), norm(title)
                if a and t:
                    songs.setdefault((a, t), {"raw": make_raw(artist, title), "artist_key": a, "title_key": t})
        wb.close()
    return list(songs.values())


def _guess_from_texts(words: list[str]) -> str:
    # Название на старых слайдах обычно заглавными, исполнитель — обычными буквами
    upper = [w for w in words if w == w.upper() and any(c.isalpha() for c in w)]
    rest = [w for w in words if w not in upper]
    if upper and rest:
        return make_raw(" ".join(rest), " ".join(upper).capitalize())
    return make_raw(" ".join(words[:-1]), words[-1])


def identify_song(texts: list[str], svod: list[dict]) -> tuple[str | None, str]:
    words = [t.strip() for t in texts if t.strip() and not t.strip().isdigit()]
    blob = norm(" ".join(words))
    if not blob:
        return None, ""
    hits = [s for s in svod if s["artist_key"] in blob and s["title_key"] in blob]
    if not hits:
        return _guess_from_texts(words), CHECK_UNKNOWN
    length = lambda s: len(s["artist_key"]) + len(s["title_key"])  # noqa: E731
    hits.sort(key=length, reverse=True)
    rivals = [h for h in hits[1:] if length(h) == length(hits[0]) and h["raw"] != hits[0]["raw"]]
    return hits[0]["raw"], CHECK_AMBIGUOUS if rivals else ""


def _rels(z: zipfile.ZipFile, names: set, part: str) -> dict:
    base, name = posixpath.split(part)
    rels_name = f"{base}/_rels/{name}.rels"
    if rels_name not in names:
        return {}
    out = {}
    for tag in re.findall(r"<Relationship\b[^>]*>", z.read(rels_name).decode("utf-8", "replace")):
        if 'TargetMode="External"' in tag:
            continue
        rid = re.search(r'\bId="([^"]+)"', tag).group(1)
        target = re.search(r'\bTarget="([^"]+)"', tag).group(1)
        out[rid] = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
    return out


def scan_presentation(path: str, svod: list[dict]) -> list[dict]:
    found, image_use = [], Counter()
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        slides = sorted(
            (n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
            key=lambda n: int(re.search(r"(\d+)\.xml$", n).group(1)),
        )
        for part in slides:
            rels = _rels(z, names, part)
            media = [m.lower() for m in rels.values()]
            if not any(m.endswith(AUDIO_EXT) for m in media) or any(m.endswith(VIDEO_EXT) for m in media):
                continue
            xml = z.read(part).decode("utf-8", "replace")
            texts = [html.unescape(t) for t in re.findall(r"<a:t>([^<]*)</a:t>", xml)]
            raw, check = identify_song(texts, svod)
            if raw is None:
                continue
            pics = []
            for pic in re.findall(r"<p:pic>.*?</p:pic>", xml, re.S):
                if "audioFile" in pic or "videoFile" in pic:
                    continue
                emb = re.search(r'r:embed="([^"]+)"', pic)
                ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', pic)
                target = rels.get(emb.group(1), "") if emb else ""
                if target.lower().endswith(IMAGE_EXT):
                    # фото в рамке макета не имеет своего размера — оно во всю рамку, считаем его крупным
                    area = int(ext.group(1)) * int(ext.group(2)) if ext else PLACEHOLDER_AREA
                    pics.append((target, area))
            for target in {p[0] for p in pics}:
                image_use[target] += 1
            found.append({"raw": raw, "check": check, "source": path.replace("\\", "/"),
                          "slide_part": part, "pics": pics, "has_trim": "<p14:trim" in xml})
    limit = max(2, DECOR_SHARE * len(found))
    for f in found:
        photos = [p for p in f.pop("pics") if image_use[p[0]] < limit]
        f["photo_part"] = max(photos, key=lambda p: p[1])[0] if photos else None
    return found


def choose_best(entries: list[dict]) -> list[dict]:
    """Из дублей одной песни — лучший слайд: узнан уверенно, есть фото, задана обрезка."""
    groups = {}
    for e in entries:
        groups.setdefault(norm(e["raw"]), []).append(e)
    result = []
    for group in groups.values():
        group.sort(key=lambda e: (e["check"] != "", e["photo_part"] is None, not e["has_trim"]))
        best = dict(group[0])
        best["spares"] = [{"source": e["source"], "slide_part": e["slide_part"]} for e in group[1:]]
        result.append(best)
    return sorted(result, key=lambda e: e["raw"].lower())
