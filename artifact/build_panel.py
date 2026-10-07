"""Собирает страницу «Панель мероприятий» со встроенным банком песен.

artifact/index.html хранится в git без списка песен; собранная страница
(artifact/dist/index.html) содержит песни из «Банк песен.xlsx» и в git не попадает.
Запуск из корня проекта: python artifact/build_panel.py
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts.song_bank import load_song_bank  # noqa: E402

PLACEHOLDER = '<script id="songs-data" type="application/json">[]</script>'


def build(bank_path="Банк песен.xlsx", source="artifact/index.html", output="artifact/dist/index.html"):
    songs = [song["raw"] for song in load_song_bank(bank_path)]
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
