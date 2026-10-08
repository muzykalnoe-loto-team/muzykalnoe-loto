import json
import os

from scripts import paths
from scripts.barrel_kinds import bank_index, check_kinds, resolve_files
from scripts.build_presentation import assemble_presentation
from scripts.generate_blanks import generate_blanks
from scripts.song_bank import load_song_bank
from scripts.validate_selection import validate_selection

BLANKS_NAME = "Бланки.xlsm"
PRESENTATION_NAME = "Презентация.pptx"


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
) -> dict:
    with open(selected_songs_path, encoding="utf-8") as f:
        barrels = sorted(json.load(f)["barrels"], key=lambda b: b["number"])
    songs_in_order = [b["song"] for b in barrels]

    # Всё проверяем до запуска Excel и PowerPoint: сборка долгая, падать на середине обидно
    bank = _load_bank(song_bank)
    check = validate_selection(songs_in_order, bank)
    if not check["ok"]:
        raise ValueError("Список песен не прошёл проверку: " + "; ".join(check["errors"]))
    if [b["number"] for b in barrels] != list(range(1, 91)):
        raise ValueError("Номера бочонков должны идти от 1 до 90 без пропусков")
    index = bank_index(bank)
    kind_errors = check_kinds(barrels, index)
    if kind_errors:
        raise ValueError("Виды слайдов не на своих местах: " + "; ".join(kind_errors))

    blanks_path = os.path.join(output_dir, BLANKS_NAME)
    generate_blanks(blanks_template, blanks_path, songs_in_order, n_individual, n_team)

    presentation_path = os.path.join(output_dir, PRESENTATION_NAME)
    assemble_presentation(presentation_template, archive_folder, resolve_files(barrels, index), presentation_path)

    return {"blanks_path": blanks_path, "presentation_path": presentation_path}
