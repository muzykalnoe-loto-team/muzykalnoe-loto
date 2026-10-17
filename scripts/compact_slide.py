# Сжатие старого слайда без видимой и слышимой разницы: из mp3 остаётся только играемый отрывок
# с запасом 5 секунд по краям (целыми кадрами, без пересжатия), фото — только видимая в рамке часть, не шире 1920 точек
import io
import posixpath
import re
import zipfile

from PIL import Image

from scripts.slide_split import CONTENT_TYPES, _attr, _rel_tags, _rels_path, _resolve

MPEG1, MPEG2, MPEG25 = 3, 2, 0
LAYER3, LAYER2 = 1, 2
BITRATES = {  # кбит/с по индексу из заголовка кадра
    (MPEG1, LAYER3): [0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320],
    (MPEG1, LAYER2): [0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384],
    "MPEG2": [0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160],
}
SAMPLE_RATES = {MPEG1: [44100, 48000, 32000], MPEG2: [22050, 24000, 16000], MPEG25: [11025, 12000, 8000]}
VBR_TAGS = (b"Xing", b"Info", b"VBRI")  # служебный кадр с длиной всей песни — после обрезки он бы врал
JPEG_QUALITY = 90
MARGIN_MS = 5000  # запас звука до и после отрывка (просьба оператора)


def _id3_size(data: bytes) -> int:
    if data[:3] != b"ID3" or len(data) < 10:
        return 0
    size = (data[6] << 21) | (data[7] << 14) | (data[8] << 7) | data[9]
    return 10 + size + (10 if data[5] & 0x10 else 0)


def _frame(data: bytes, pos: int) -> tuple[int, int, int] | None:
    """(длина кадра в байтах, отсчётов в кадре, частота) или None, если здесь не заголовок кадра."""
    if pos + 4 > len(data) or data[pos] != 0xFF or data[pos + 1] & 0xE0 != 0xE0:
        return None
    version, layer = (data[pos + 1] >> 3) & 3, (data[pos + 1] >> 1) & 3
    bitrate_i, rate_i, padding = data[pos + 2] >> 4, (data[pos + 2] >> 2) & 3, (data[pos + 2] >> 1) & 1
    if version == 1 or layer not in (LAYER3, LAYER2) or bitrate_i in (0, 15) or rate_i == 3:
        return None
    table = BITRATES[(MPEG1, layer)] if version == MPEG1 else BITRATES["MPEG2"]
    bitrate, rate = table[bitrate_i] * 1000, SAMPLE_RATES[version][rate_i]
    samples = 1152 if version == MPEG1 or layer == LAYER2 else 576
    return samples // 8 * bitrate // rate + padding, samples, rate


def mp3_frames(data: bytes) -> list[tuple[int, int, float]]:
    """Звуковые кадры mp3: (начало в байтах, длина в байтах, длительность в мс).
    Кадр засчитывается, только если сразу за ним — следующий кадр, тег или конец файла:
    иначе случайные байты «0xFF 0xFB» внутри тегов и мусора сбивают разбор."""
    frames, pos, n = [], 0, len(data)
    while pos < n:
        if data[pos:pos + 3] == b"ID3":  # тег ID3v2 — в начале, а бывает и посреди файла
            size = _id3_size(data[pos:pos + 10])
            if size:
                pos += size
                continue
        if data[pos:pos + 3] == b"TAG" and n - pos <= 128:  # тег ID3v1 в конце файла
            break
        info = _frame(data, pos)
        nxt = pos + info[0] if info else 0
        if info is None or nxt > n or not (nxt == n or _frame(data, nxt) or data[nxt:nxt + 3] in (b"ID3", b"TAG")):
            pos += 1  # не кадр — ищем следующий заголовок
            continue
        length, samples, rate = info
        if not any(tag in data[pos + 4:pos + 40] for tag in VBR_TAGS):
            frames.append((pos, length, samples / rate * 1000))
        pos = nxt
    return frames


def fragment_bounds(total_ms: float, st: float, end_trimmed: float) -> tuple[float, float]:
    """Начало и конец играемого отрывка (мс). Ошибка, если отрывок не помещается в прочитанный звук —
    значит, mp3 разобран не полностью, и сжимать его нельзя (песня в игре замолчала бы)."""
    end_point = total_ms - end_trimmed
    if not 0 <= st < end_point:
        raise ValueError(f"отрывок {st:.0f}–{end_point:.0f} мс вне прочитанного звука ({total_ms:.0f} мс)")
    return st, end_point


def _bitrate(data: bytes, pos: int) -> int:
    version, layer = (data[pos + 1] >> 3) & 3, (data[pos + 1] >> 1) & 3
    table = BITRATES[(MPEG1, layer)] if version == MPEG1 else BITRATES["MPEG2"]
    return table[data[pos + 2] >> 4] * 1000


def mp3_clock(data: bytes) -> tuple[list[tuple[int, int, float]], list[float], float]:
    """Кадры, время начала каждого и длина всего звука — так, как их считает PowerPoint.
    Есть служебный кадр с длиной (Xing/Info/VBRI) — время по кадрам; нет — по размеру файла
    от начала звука с битрейтом первого кадра (мусор посреди файла для PowerPoint тоже «звучит»)."""
    frames = mp3_frames(data)
    if not frames:
        return frames, [], 0.0
    if any(tag in data[:frames[0][0] + 64] for tag in VBR_TAGS):
        starts, t = [], 0.0
        for _, _, duration in frames:
            starts.append(t)
            t += duration
        return frames, starts, t
    audio_start = _id3_size(data[:10])
    bytes_per_ms = _bitrate(data, frames[0][0]) / 8 / 1000
    end = len(data) - (128 if data[-128:-125] == b"TAG" else 0)
    return frames, [(pos - audio_start) / bytes_per_ms for pos, _, _ in frames], (end - audio_start) / bytes_per_ms


def cut_mp3(data: bytes, start_ms: float, end_ms: float, margin_ms: float = MARGIN_MS) -> tuple[bytes, float]:
    """Оставляет кадры отрывка [start_ms; end_ms] (по часам PowerPoint) с запасом margin_ms по краям.
    Возвращает новые байты и время (мс) в старом файле, с которого начинается новый."""
    frames, starts, _ = mp3_clock(data)
    kept = [(pos, length, t) for (pos, length, duration), t in zip(frames, starts)
            if t + duration > start_ms - margin_ms and t < end_ms + margin_ms]
    if not kept:
        raise ValueError("в отрывок не попал ни один кадр звука")
    if len({_bitrate(data, pos) for pos, _, _ in kept}) > 1:
        # без служебного кадра PowerPoint посчитал бы время такого файла по размеру — неверно
        raise ValueError("переменный битрейт — звук не обрезаем")
    return b"".join(data[pos:pos + length] for pos, length, _ in kept), kept[0][2]


def shifted_trim(st: float, end_trimmed: float, total_ms: float, cut_from_ms: float, kept_ms: float) -> tuple[float, float]:
    """Настройка отрывка для обрезанного файла: «отрезать с начала» и «отрезать с конца» (мс)."""
    end_point = total_ms - end_trimmed
    return round(st - cut_from_ms, 3), round(kept_ms - (end_point - cut_from_ms), 3)


def crop_photo(data: bytes, src_rect: tuple[int, int, int, int], max_width: int = 1920) -> bytes:
    """src_rect — обрезка из слайда (l, t, r, b) в тысячных долях процента, как в <a:srcRect>."""
    img = Image.open(io.BytesIO(data)).convert("RGB")
    w, h = img.size
    l, t, r, b = (max(0, v) / 100000 for v in src_rect)  # отрицательная обрезка — поля вокруг фото, их не добавляем
    img = img.crop((round(w * l), round(h * t), round(w * (1 - r)), round(h * (1 - b))))
    if img.width > max_width:
        img = img.resize((max_width, round(img.height * max_width / img.width)), Image.LANCZOS)
    out = io.BytesIO()
    img.save(out, "JPEG", quality=JPEG_QUALITY)
    return out.getvalue()


JPEG_DEFAULT = '<Default Extension="jpg" ContentType="image/jpeg"/>'


def _new_jpeg_name(photo_part: str, names: set[str]) -> str:
    stem = posixpath.splitext(photo_part)[0]
    name, i = f"{stem}.jpg", 1
    while name in names and name != photo_part:
        name, i = f"{stem}_{i}.jpg", i + 1
    return name


def compact_slide(src: str, out_path: str, photo_part: str | None) -> str | None:
    """Файл банка старых слайдов (один слайд) → сжатая копия. Возвращает новое имя фото внутри файла."""
    with zipfile.ZipFile(src) as zin:
        names = set(zin.namelist())
        slide = next(n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n))
        rels_name = _rels_path(slide)
        xml = zin.read(slide).decode("utf-8")
        rels = zin.read(rels_name).decode("utf-8")
        targets = {_attr(t, "Id"): _resolve("ppt/slides", _attr(t, "Target")) for t in _rel_tags(rels)
                   if _attr(t, "TargetMode") != "External"}
        replaced: dict[str, bytes] = {}
        renamed: dict[str, str] = {}

        # 1. Звук: оставляем отрывок с запасом по краям, настройку отрывка сдвигаем
        media = re.search(r'<p14:media\b[^>]*r:embed="([^"]+)"[^>]*>.*?</p14:media>', xml, re.S)
        trim = re.search(r"<p14:trim\b[^>]*/>", media.group(0)) if media else None
        if trim:
            audio_part = targets[media.group(1)]
            data = zin.read(audio_part)
            total = mp3_clock(data)[2]
            st, end_trimmed = float(_attr(trim.group(0), "st") or 0), float(_attr(trim.group(0), "end") or 0)
            cut, cut_from = cut_mp3(data, *fragment_bounds(total, st, end_trimmed))
            new_st, new_end = shifted_trim(st, end_trimmed, total, cut_from, sum(d for _, _, d in mp3_frames(cut)))
            replaced[audio_part] = cut
            xml = xml.replace(trim.group(0), f'<p14:trim st="{new_st:g}" end="{max(0, new_end):g}"/>')

        # 2. Фото: только видимая в рамке часть, в JPEG; обрезка в слайде больше не нужна
        new_photo = photo_part
        if photo_part:
            rids = [rid for rid, part in targets.items() if part == photo_part]
            src_rect = (0, 0, 0, 0)
            for pic in re.findall(r"<p:pic>.*?</p:pic>", xml, re.S):
                if any(f'r:embed="{rid}"' in pic for rid in rids):
                    rect = re.search(r"<a:srcRect\b[^>]*/>", pic)
                    if rect:
                        src_rect = tuple(int(_attr(rect.group(0), k) or 0) for k in "ltrb")
                        xml = xml.replace(pic, pic.replace(rect.group(0), ""))
            new_photo = _new_jpeg_name(photo_part, names)
            replaced[new_photo] = crop_photo(zin.read(photo_part), src_rect)
            if new_photo != photo_part:
                renamed[photo_part] = new_photo

        # 3. Ссылки на переименованное фото и типы содержимого
        texts = {slide: xml}
        for name in names:
            if name.endswith(".rels") and renamed:
                text = texts.get(name) or zin.read(name).decode("utf-8")
                owner_base = posixpath.dirname(posixpath.dirname(name))
                for tag in _rel_tags(text):
                    part = _resolve(owner_base, _attr(tag, "Target"))
                    if part in renamed and _attr(tag, "TargetMode") != "External":
                        new_target = posixpath.relpath(renamed[part], owner_base or ".")
                        text = text.replace(tag, re.sub(r'\bTarget="[^"]*"', f'Target="{new_target}"', tag))
                texts[name] = text
        ct = zin.read(CONTENT_TYPES).decode("utf-8")
        if renamed and 'Extension="jpg"' not in ct:
            ct = ct.replace("<Default ", JPEG_DEFAULT + "<Default ", 1)
        for old in renamed:
            ct = re.sub(rf'<Override\b[^>]*PartName="/{re.escape(old)}"[^>]*/>', "", ct)
        texts[CONTENT_TYPES] = ct

        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                name = info.filename
                if name in renamed:
                    zout.writestr(renamed[name], replaced[renamed[name]])
                elif name in replaced:
                    zout.writestr(info, replaced[name])
                elif name in texts:
                    zout.writestr(info, texts[name].encode("utf-8"), compress_type=zipfile.ZIP_DEFLATED)
                else:
                    with zin.open(info) as fsrc, zout.open(info, "w") as fdst:
                        while chunk := fsrc.read(1 << 20):
                            fdst.write(chunk)
    return new_photo
