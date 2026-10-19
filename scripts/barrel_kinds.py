# Вид слайда у бочонка: "video" — караоке из видео-банка, "old" — старый слайд (фото + звук).
# Бочонки, кратные 3, — всегда старые слайды, остальные — видео (список из одних видео не собирается).
from scripts.song_names import norm

KIND_NAMES = {"video": "видео", "old": "старый слайд"}


def bank_index(bank: list[dict]) -> dict[str, dict]:
    # по norm: в общем банке разные написания одной песни (ё/е, регистр) слиты в одну строку
    return {norm(e["raw"]): e for e in bank}


def check_kinds(barrels: list[dict], index: dict[str, dict]) -> list[str]:
    errors = []
    for b in barrels:
        number, song, kind = b["number"], b["song"], b.get("kind", "video")
        if kind not in KIND_NAMES:
            errors.append(f"Бочонок {number}: непонятный вид «{kind}»")
            continue
        want = "old" if number % 3 == 0 else "video"
        if kind != want:
            errors.append(f"Бочонок {number} («{song}»): здесь должен быть {KIND_NAMES[want]}, а стоит {KIND_NAMES[kind]}")
            continue
        entry = index.get(norm(song), {})
        if kind == "old" and not entry.get("old_file"):
            errors.append(f"Бочонок {number}: у песни «{song}» нет старого слайда")
        if kind == "video" and not entry.get("video_file"):
            errors.append(f"Бочонок {number}: у песни «{song}» нет видео")
    return errors


def resolve_files(barrels: list[dict], index: dict[str, dict]) -> list[dict]:
    resolved = []
    for b in barrels:
        entry = index[norm(b["song"])]
        kind = b.get("kind", "video")
        item = dict(b, kind=kind)
        if kind == "old":
            item.update(file=entry["old_file"], photo_part=entry.get("old_photo_part"))
        else:
            item.update(file=entry["video_file"], photo_part=None)
        resolved.append(item)
    return resolved
