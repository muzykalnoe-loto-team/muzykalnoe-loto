import json
import os
import shutil

import openpyxl
import pytest
from pptx import Presentation

from scripts import paths
from scripts.export_selection import export_selection
from scripts.run_pipeline import run_pipeline

SELECTION = f"{paths.EVENTS}/пример/selected_songs.json"


@pytest.mark.integration
def test_run_pipeline_produces_both_files(tmp_path):
    selected_path = tmp_path / "selected_songs.json"
    shutil.copy(SELECTION, selected_path)

    result = run_pipeline(str(selected_path), n_individual=3, n_team=2, output_dir=str(tmp_path))

    blanks = tmp_path / "Бланки.xlsm"
    presentation = tmp_path / "Презентация.pptx"
    assert result == {"blanks_path": str(blanks), "presentation_path": str(presentation)}

    wb = openpyxl.load_workbook(blanks, data_only=True)
    with open(SELECTION, encoding="utf-8") as f:
        barrels = json.load(f)["barrels"]
    # в шаблон бланков ушли все 90 песен по порядку бочонков
    assert wb["Список"]["B2"].value == barrels[0]["song"]
    assert wb["Список"]["E1"].value == 3 and wb["Список"]["E2"].value == 2

    prs = Presentation(str(presentation))
    assert len(prs.slides) == 118
    # на слоте бочонка 90 (слайд 105) — видео нужной песни
    names = [s.name for s in prs.slides[104].shapes]
    assert any(n.startswith(barrels[89]["song"].split(" - ")[0]) for n in names)


def test_run_pipeline_rejects_bad_selection_before_office(tmp_path):
    selected_path = tmp_path / "selected_songs.json"
    with open(selected_path, "w", encoding="utf-8") as f:
        json.dump({"barrels": [{"number": 1, "song": "Несуществующий Артист - Песня"}]}, f, ensure_ascii=False)

    with pytest.raises(ValueError, match="90"):
        run_pipeline(str(selected_path), n_individual=1, n_team=1, output_dir=str(tmp_path))

    assert os.listdir(tmp_path) == ["selected_songs.json"]


def _common_bank():
    if not os.path.exists(paths.COMMON_BANK_JSON):
        pytest.skip("общий банк ещё не собран (Task 4)")
    from scripts.common_bank import load_common_bank
    return load_common_bank()


def test_run_pipeline_rejects_old_slide_on_wrong_barrel(tmp_path):
    bank = _common_bank()
    old = [e["raw"] for e in bank if e["old_file"]][:30]
    video = [e["raw"] for e in bank if e["video_file"] and e["raw"] not in old][:60]
    songs, kinds = [], []
    for n in range(1, 91):
        is_old = n % 3 == 0
        songs.append(old.pop() if is_old else video.pop())
        kinds.append("old" if is_old else "video")
    kinds[3], kinds[2] = "old", "video"  # бочонок 4 — старый слайд, бочонок 3 — видео: ошибка
    selected_path = tmp_path / "selected_songs.json"
    export_selection(songs, str(selected_path), kinds=kinds)

    with pytest.raises(ValueError, match="Бочонок 3"):
        run_pipeline(str(selected_path), n_individual=1, n_team=1, output_dir=str(tmp_path))
    assert sorted(os.listdir(tmp_path)) == ["selected_songs.json"]


def test_earlier_event_list_passes_checks_with_common_bank():
    # список «Готического бала» собран до общего банка: часть песен в нём написана иначе (ё/е, регистр)
    from scripts.barrel_kinds import bank_index, check_kinds, resolve_files
    from scripts.validate_selection import validate_selection
    bank = _common_bank()
    with open("tests/data/Готический бал — список.json", encoding="utf-8") as f:
        barrels = json.load(f)["barrels"]

    assert validate_selection([b["song"] for b in barrels], bank)["errors"] == []
    index = bank_index(bank)
    assert check_kinds(barrels, index) == []
    for b in resolve_files(barrels, index):
        assert os.path.exists(os.path.join(paths.SLIDES_ARCHIVE, b["file"]))
