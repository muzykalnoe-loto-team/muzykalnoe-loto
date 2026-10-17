import io

from PIL import Image

from scripts.compact_slide import crop_photo, cut_mp3, mp3_frames, shifted_trim

# MPEG-1 Layer III, 128 кбит/с, 44,1 кГц, без padding: кадр 417 байт, 1152 отсчёта ≈ 26,12 мс
HEADER = b"\xff\xfb\x90\x00"
FRAME = HEADER + b"\x00" * 413
FRAME_MS = 1152 / 44100 * 1000


def _mp3(frames=100, id3=b""):
    return id3 + FRAME * frames


def test_frames_are_found_after_id3_tag():
    tag = b"ID3\x03\x00\x00\x00\x00\x00\x0a" + b"x" * 10  # тег ID3 размером 10 байт
    frames = mp3_frames(_mp3(5, tag))
    assert len(frames) == 5 and frames[0][0] == 20 and frames[0][1] == 417


def test_cut_keeps_fragment_with_margin_and_drops_the_rest():
    data, cut_from_ms = cut_mp3(_mp3(100), start_ms=1000, end_ms=1500, margin_ms=200)

    kept = len(mp3_frames(data))
    assert data == FRAME * kept  # только целые кадры, без пересжатия
    assert cut_from_ms <= 800 < cut_from_ms + FRAME_MS  # запас 200 мс до начала отрывка
    assert cut_from_ms + kept * FRAME_MS >= 1700  # и после конца
    assert kept * FRAME_MS < 1700 - 800 + 3 * FRAME_MS  # остальное выброшено


def test_cut_at_song_start_does_not_go_negative():
    data, cut_from_ms = cut_mp3(_mp3(100), start_ms=100, end_ms=500, margin_ms=1000)
    assert cut_from_ms == 0


def test_shifted_trim_keeps_the_same_fragment():
    # было: песня 160 835 мс, отрывок 31 269 … 53 966 (с конца отрезано 106 869 мс)
    # стало: файл с 30 000 мс длиной 25 000 мс → тот же отрывок 1 269 … 23 966
    st, end = shifted_trim(st=31269, end_trimmed=106869, total_ms=160835, cut_from_ms=30000, kept_ms=25000)
    assert st == 1269 and end == 25000 - 23966


def _png(w, h):
    img = Image.new("RGB", (w, h), (255, 0, 0))
    img.paste((0, 0, 255), (w // 2, 0, w, h))  # правая половина синяя
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def test_crop_photo_keeps_only_visible_part_then_shrinks():
    # в слайде видна только правая (синяя) половина фото 4000×1000: srcRect l=50 %
    jpeg = crop_photo(_png(4000, 1000), (50000, 0, 0, 0), max_width=1920)

    img = Image.open(io.BytesIO(jpeg))
    assert img.format == "JPEG" and img.size == (1920, 960)
    r, g, b = img.getpixel((10, 10))
    assert b > 200 and r < 50  # красной части нет — уменьшена только видимая


def test_small_photo_is_not_enlarged():
    img = Image.open(io.BytesIO(crop_photo(_png(800, 400), (0, 0, 0, 0), max_width=1920)))
    assert img.size == (800, 400)


def _bank_entry(tmp_path):
    """Несжатый слайд «Little Big - UNO», нарезанный заново из старой игры (файл банка уже сжат)."""
    import json
    import os
    import posixpath
    import zipfile

    import pytest

    from scripts import paths
    from scripts.slide_split import extract_slide
    if not os.path.exists(paths.OLD_CATALOG):
        pytest.skip("банк старых слайдов ещё не собран")
    with open(paths.OLD_CATALOG, encoding="utf-8") as f:
        entry = next(c for c in json.load(f) if c["raw"].startswith("Little Big - UNO"))
    original = str(tmp_path / "оригинал.pptx")
    extract_slide(entry["source"], entry["slide_part"], original)
    # в каталоге — имя уже сжатого фото (.jpg); у оригинала то же имя, но своё расширение
    stem = posixpath.splitext(entry["photo_part"])[0]
    with zipfile.ZipFile(original) as z:
        photo = next(n for n in z.namelist() if posixpath.splitext(n)[0] == stem)
    return original, photo


def test_compact_slide_makes_file_smaller_and_photo_jpeg(tmp_path):
    import os
    import re
    import zipfile

    from scripts.compact_slide import compact_slide
    src, photo_part = _bank_entry(tmp_path)
    out = tmp_path / "сжатый.pptx"

    new_photo = compact_slide(src, str(out), photo_part)

    assert os.path.getsize(out) < os.path.getsize(src) / 3
    with zipfile.ZipFile(out) as z:
        assert Image.open(io.BytesIO(z.read(new_photo))).format == "JPEG"
        slide = next(n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n))
        assert "<a:srcRect" not in z.read(slide).decode()  # обрезка уже применена к самому фото


import pytest  # noqa: E402


@pytest.mark.integration
def test_compact_slide_plays_the_same_fragment(tmp_path):
    import win32com.client as win32

    from scripts.compact_slide import compact_slide
    src, photo_part = _bank_entry(tmp_path)
    out = tmp_path / "сжатый.pptx"
    compact_slide(src, str(out), photo_part)

    app = win32.Dispatch("PowerPoint.Application")
    fragments = []
    try:
        for path in (src, str(out)):
            prs = app.Presentations.Open(path, ReadOnly=True, WithWindow=False)
            try:
                media = next(s for s in prs.Slides(1).Shapes if s.Type == 16).MediaFormat
                fragments.append((media.EndPoint - media.StartPoint, media.FadeInDuration, media.FadeOutDuration,
                                  media.StartPoint, media.Length))
            finally:
                prs.Close()
    finally:
        if app.Presentations.Count == 0:
            app.Quit()
    (before_len, *before_fades, _, before_total), (after_len, *after_fades, after_start, after_total) = fragments
    assert abs(after_len - before_len) < 30  # отрывок той же длины
    assert after_fades == before_fades
    assert 4900 <= after_start <= 5100  # 5 секунд запаса до отрывка (просьба оператора)
    assert after_total < before_total / 3  # остальная песня выброшена
