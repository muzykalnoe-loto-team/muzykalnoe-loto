import os
import shutil

import pytest
import win32com.client as win32

from scripts.ppt_paste import insert_song_slide

TEMPLATE = "Рабочий 2 вар.pptx"
ARCHIVE_SONG = "Банк всех слайдов23.01.2026/Натали - Шахерезада.pptx"


@pytest.mark.integration
def test_insert_song_slide_adds_shapes_without_removing_existing(tmp_path):
    working_copy = tmp_path / "template_copy.pptx"
    shutil.copy(TEMPLATE, working_copy)

    ppt_app = win32.Dispatch("PowerPoint.Application")
    try:
        archive_prs = ppt_app.Presentations.Open(
            os.path.abspath(ARCHIVE_SONG), ReadOnly=True, WithWindow=False
        )
        archive_positions = [(s.Left, s.Top) for s in archive_prs.Slides(1).Shapes]
        archive_prs.Close()

        template_prs = ppt_app.Presentations.Open(str(working_copy), WithWindow=False)
        try:
            target_slide = template_prs.Slides(16)  # слот бочонка №1
            names_before = [s.Name for s in target_slide.Shapes]

            insert_song_slide(ppt_app, template_prs, ARCHIVE_SONG, 16)

            shapes = list(target_slide.Shapes)
            names_after = [s.Name for s in shapes]
            assert len(names_after) == len(names_before) + len(archive_positions)
            assert template_prs.Slides.Count == 118

            # фигуры шаблона (в т.ч. кнопка навигации) на месте и лежат поверх вставленного
            assert names_after[-len(names_before):] == names_before

            # вставленное стоит там же, где на слайде архива, без сдвига
            pasted = shapes[: len(archive_positions)]
            for shape, (left, top) in zip(pasted, archive_positions):
                assert (round(shape.Left), round(shape.Top)) == (round(left), round(top))
        finally:
            template_prs.Close()
    finally:
        ppt_app.Quit()
