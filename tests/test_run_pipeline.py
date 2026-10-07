import json
import os
import shutil

import openpyxl
import pytest
from pptx import Presentation

from scripts import paths
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
