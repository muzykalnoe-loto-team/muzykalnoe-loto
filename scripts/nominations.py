# Какая номинация будет в каждом раунде игры и запись о прошлых играх.
# SlideID — внутренний номер слайда в шаблоне: не меняется, когда другие слайды удаляются
import json
import os

from scripts import paths

# Порядок вариантов задаёт очередь: первый — для самой первой игры
ROUNDS = {
    1: {"theme": {"name": "Тематическая", "slide_id": 374},
        "sing": {"name": "Щас спою!", "slide_id": 375},
        "cry": {"name": "Крик души", "slide_id": 258}},
    2: {"goddess": {"name": "Синяя богиня", "slide_id": 259},
        "cherry": {"name": "Пьяная вишня", "slide_id": 376}},
    3: {"dancefloor": {"name": "Танцпол в огне", "slide_id": 377},
        "nobrakes": {"name": "Мы без тормозов", "slide_id": 378}},
}
THEME = "theme"


def _regular(round_no: int) -> list[str]:
    """Обычные варианты раунда (тематический ставится только галочкой)."""
    return [key for key in ROUNDS[round_no] if key != THEME]


def _next_in_queue(round_no: int, history: list[dict]) -> str:
    # противоположная той, что была в прошлой игре в этом раунде; для 1-го раунда — только обычные игры
    variants = _regular(round_no)
    for game in reversed(history):
        if round_no == 1 and game.get("themed"):
            continue
        last = game.get("nominations", {}).get(str(round_no))
        if last in variants:
            return variants[(variants.index(last) + 1) % len(variants)]
    return variants[0]


def choose_nominations(themed: bool, choices: dict[int, str], history: list[dict]) -> dict[int, str]:
    chosen = {}
    for round_no in ROUNDS:
        choice = choices.get(round_no, "auto")
        if round_no == 1 and themed:
            chosen[round_no] = THEME  # галочка «Тематический слайд» важнее выбора в брифе
        elif choice == "auto":
            chosen[round_no] = _next_in_queue(round_no, history)
        elif choice in _regular(round_no):
            chosen[round_no] = choice
        else:
            raise ValueError(f"Номинации «{choice}» нет среди вариантов {round_no}-го раунда")
    return chosen


def load_history(path: str = paths.NOMINATION_HISTORY) -> list[dict]:
    # нет записи или файл испорчен — считаем, что игр ещё не было: сборка не должна падать
    try:
        with open(path, encoding="utf-8") as f:
            history = json.load(f)
    except (OSError, ValueError):
        return []
    return history if isinstance(history, list) else []


def record_game(event: str, date: str, themed: bool, chosen: dict[int, str],
                path: str = paths.NOMINATION_HISTORY) -> None:
    history = load_history(path)
    history.append({"event": event, "date": date, "themed": themed,
                    "nominations": {str(r): key for r, key in chosen.items()}})
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)
