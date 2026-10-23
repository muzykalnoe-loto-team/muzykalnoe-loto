import json
import os
import tempfile

from scripts import paths
from scripts.barrel_kinds import bank_index, check_kinds, resolve_files
from scripts.build_presentation import assemble_presentation
from scripts.generate_blanks import generate_blanks
from scripts.nominations import choose_nominations, load_history, record_game
from scripts.song_bank import load_song_bank
from scripts.stop_list import load_stop_list
from scripts.validate_selection import validate_selection

BLANKS_NAME = "Бланки.xlsm"
PRESENTATION_NAME = "Презентация.pptx"
SIZE_LIMIT_MB = 995  # больше Яндекс Диск может не принять (решение оператора)


def size_warning(path: str, limit_mb: int = SIZE_LIMIT_MB) -> str:
    size_mb = os.path.getsize(path) / 2**20
    if size_mb <= limit_mb:
        return ""
    return f"Презентация весит {size_mb:.0f} МБ — больше {limit_mb} МБ: на Яндекс Диск может не загрузиться"


def _load_bank(song_bank: str) -> list[dict]:
    if os.path.exists(paths.COMMON_BANK_JSON):
        from scripts.common_bank import load_common_bank
        return load_common_bank()
    # общий банк ещё не собран — работаем только с видео-банком, как раньше
    return [{"raw": s["raw"], "video_file": f"{s['raw']}.pptx", "old_file": None, "old_photo_part": None}
            for s in load_song_bank(song_bank)]


def run_pipeline(
    selected_songs_path: str,
    n_individual: int,
    n_team: int,
    output_dir: str,
    blanks_template: str = paths.BLANKS_TEMPLATE,
    presentation_template: str = paths.PRESENTATION_TEMPLATE,
    archive_folder: str = paths.SLIDES_ARCHIVE,
    song_bank: str = paths.SONG_BANK,
    event: str = "",
    event_date: str = "",
    nominations: dict | None = None,
) -> dict:
    # nominations = {"themed": bool, "choices": {раунд: "auto" | ключ}, "theme": {"title", "glow", "art_png", "name"}};
    # None — слайды номинаций не трогаются
    with open(selected_songs_path, encoding="utf-8") as f:
        barrels = sorted(json.load(f)["barrels"], key=lambda b: b["number"])
    songs_in_order = [b["song"] for b in barrels]

    # Всё проверяем до запуска Excel и PowerPoint: сборка долгая, падать на середине обидно
    bank = _load_bank(song_bank)
    check = validate_selection(songs_in_order, bank, load_stop_list())
    if not check["ok"]:
        raise ValueError("Список песен не прошёл проверку: " + "; ".join(check["errors"]))
    if [b["number"] for b in barrels] != list(range(1, 91)):
        raise ValueError("Номера бочонков должны идти от 1 до 90 без пропусков")
    index = bank_index(bank)
    kind_errors = check_kinds(barrels, index)
    if kind_errors:
        raise ValueError("Виды слайдов не на своих местах: " + "; ".join(kind_errors))
    chosen = theme = None
    if nominations is not None:
        themed = bool(nominations.get("themed"))
        theme = nominations.get("theme")
        if themed and not (theme and theme.get("art_png")):
            raise ValueError("Тематическая номинация: нет картинки темы — сначала выберите вариант слайда")
        choices = {int(r): key for r, key in (nominations.get("choices") or {}).items()}
        chosen = choose_nominations(themed, choices, load_history(paths.NOMINATION_HISTORY))

    blanks_path = os.path.join(output_dir, BLANKS_NAME)
    generate_blanks(blanks_template, blanks_path, songs_in_order, n_individual, n_team)

    presentation_path = os.path.join(output_dir, PRESENTATION_NAME)
    finish = None
    if chosen is not None:
        from scripts.nomination_slides import apply_nominations
        def finish(prs):
            apply_nominations(prs, chosen, theme)
    assemble_presentation(presentation_template, archive_folder, resolve_files(barrels, index), presentation_path,
                          finish=finish)

    warnings = [size_warning(presentation_path)]
    if chosen is not None:
        record_game(event, event_date, bool(nominations.get("themed")), chosen, paths.NOMINATION_HISTORY)
        if chosen[1] == "theme":
            warnings.append(_archive_theme_slide(theme, event, event_date))
    warning = "; ".join(w for w in warnings if w)
    if warning:
        print(warning)
    return {"blanks_path": blanks_path, "presentation_path": presentation_path, "warning": warning,
            "nominations": chosen}


def _archive_theme_slide(theme: dict, event: str, event_date: str) -> str:
    # в архив — маленький файл с одним тематическим слайдом, а не вся игра
    from scripts.nomination_slides import append_to_archive, theme_slide_file
    with tempfile.TemporaryDirectory() as tmp:
        single = theme_slide_file(theme, os.path.join(tmp, "Тематический слайд.pptx"))
        return append_to_archive(single, event, event_date, theme.get("name", ""), paths.NOMINATION_ARCHIVE)
