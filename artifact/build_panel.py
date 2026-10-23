"""Собирает страницу «Панель мероприятий» со встроенным банком песен.

artifact/index.html хранится в git без списка песен; собранная страница
(artifact/dist/index.html) содержит песни общего банка (видео + старые слайды; если его нет —
«Исходники/Банк песен.xlsx») и в git не попадает. Каждая песня — [название, 1 если есть только старый слайд].
Запуск из корня проекта: python artifact/build_panel.py
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts import paths  # noqa: E402
from scripts.song_bank import load_song_bank  # noqa: E402
from scripts.stop_list import blocked_artist, load_stop_list  # noqa: E402

PLACEHOLDER = '<script id="songs-data" type="application/json">[]</script>'


def build(bank_path=paths.SONG_BANK, source="artifact/index.html", output="artifact/dist/index.html"):
    if os.path.exists(paths.COMMON_BANK_JSON):
        with open(paths.COMMON_BANK_JSON, encoding="utf-8") as f:
            songs = [[e["raw"], 0 if e["video_file"] else 1] for e in json.load(f)]
    else:
        songs = [[s["raw"], 0] for s in load_song_bank(bank_path)]
    stop = load_stop_list()  # исполнителей из стоп-листа в поиске «песен по желанию» нет
    songs = [s for s in songs if not blocked_artist(s[0], stop)]
    with open(source, encoding="utf-8") as f:
        page = f.read()
    if PLACEHOLDER not in page:
        raise ValueError(f"В {source} нет места для списка песен")
    # "</" внутри песни закрыл бы тег script — экранируем
    data = json.dumps(songs, ensure_ascii=False, separators=(",", ":")).replace("</", r"<\/")
    page = page.replace(PLACEHOLDER, PLACEHOLDER.replace("[]", data))
    os.makedirs(os.path.dirname(output), exist_ok=True)
    with open(output, "w", encoding="utf-8") as f:
        f.write(page)
    return len(songs)


if __name__ == "__main__":
    print(f"Встроено песен: {build()}")
