import re

_FORBIDDEN = re.compile(r'[<>:"/\|?*\x00-\x1f]')


def norm(text: str) -> str:
    """Ключ для сравнения песен: без регистра и пунктуации, ё = е."""
    return re.sub(r"[^0-9a-zа-я]", "", str(text).lower().replace("ё", "е"))


def split_raw(raw: str) -> tuple[str, str]:
    """«Исполнитель - Название» → (исполнитель, название). В базе бывает дефис без пробела."""
    if " - " in raw:
        artist, title = raw.split(" - ", 1)
    else:
        artist, _, title = raw.partition("-")
    return artist.strip(), title.strip()


def make_raw(artist: str, title: str) -> str:
    return f"{' '.join(str(artist).split())} - {' '.join(str(title).split())}"


def safe_filename(raw: str) -> str:
    """Имя файла Windows из названия песни: запрещённые символы → «_», без точек в конце."""
    return _FORBIDDEN.sub("_", raw).strip().rstrip(". ")[:150]
