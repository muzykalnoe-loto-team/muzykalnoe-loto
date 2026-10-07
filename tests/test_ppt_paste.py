import os
import shutil

import pytest
import win32com.client as win32

from scripts.ppt_paste import insert_song_slide

TEMPLATE = "Рабочий 2 вар.pptx"
ARCHIVE_SONG = "Банк всех слайдов23.01.2026/Натали - Шахерезада.pptx"
# на слайде этой песни в поле номера осталась цифра "2" из прошлого мероприятия
ARCHIVE_SONG_WITH_OLD_NUMBER = "Банк всех слайдов23.01.2026/Руки Вверх! - Девочки колхозницы.pptx"


@pytest.fixture
def template(tmp_path):
    working_copy = tmp_path / "template_copy.pptx"
    shutil.copy(TEMPLATE, working_copy)
    ppt_app = win32.Dispatch("PowerPoint.Application")
    prs = ppt_app.Presentations.Open(str(working_copy), WithWindow=False)
    try:
        yield ppt_app, prs
    finally:
        prs.Close()
        if ppt_app.Presentations.Count == 0:  # не закрываем PowerPoint с чужими файлами
            ppt_app.Quit()


def _top_shape_at(slide, x, y):
    hits = [s for s in slide.Shapes
            if s.Left <= x <= s.Left + s.Width and s.Top <= y <= s.Top + s.Height]
    return max(hits, key=lambda s: s.ZOrderPosition)


def _texts(slide):
    return [s.TextFrame.TextRange.Text for s in slide.Shapes if s.HasTextFrame]


@pytest.mark.integration
def test_insert_song_slide_keeps_navigation_on_top(template):
    ppt_app, prs = template
    slide = prs.Slides(16)  # слот бочонка №1, с кнопкой навигации
    names_before = [s.Name for s in slide.Shapes]
    button = next(s for s in slide.Shapes if s.ActionSettings(1).Action != 0)

    insert_song_slide(ppt_app, prs, ARCHIVE_SONG, 16)

    names_after = [s.Name for s in slide.Shapes]
    assert len(names_after) > len(names_before)
    assert set(names_before) <= set(names_after)  # фигуры шаблона на месте
    assert prs.Slides.Count == 118

    # видео песни стоит ровно на слайде, без сдвига вставки
    video = next(s for s in slide.Shapes if s.Type == 16)
    assert (round(video.Left), round(video.Top)) == (0, 0)

    # нажатие на кнопку попадает в рабочую кнопку шаблона
    top = _top_shape_at(slide, button.Left + button.Width / 2, button.Top + button.Height / 2)
    assert top.Name == button.Name


@pytest.mark.integration
def test_insert_song_slide_shows_only_barrel_number(template):
    ppt_app, prs = template
    slide = prs.Slides(56)  # слот бочонка №41

    insert_song_slide(ppt_app, prs, ARCHIVE_SONG_WITH_OLD_NUMBER, 56)

    numbers = [t.strip() for t in _texts(slide) if t.strip().isdigit()]
    assert numbers == ["41"]


@pytest.mark.integration
def test_insert_song_slide_hides_template_placeholder_text(template):
    ppt_app, prs = template
    slide = prs.Slides(18)  # слот бочонка №3: без кнопки, с заглушкой "ТЕКСТ НА 2 СТРОЧКИ"
    placeholder = next(s for s in slide.Shapes if s.Type == 6)

    insert_song_slide(ppt_app, prs, ARCHIVE_SONG, 18)

    video = next(s for s in slide.Shapes if s.Type == 16)
    assert placeholder.ZOrderPosition < video.ZOrderPosition  # заглушка под видео
    number = next(s for s in slide.Shapes if s.HasTextFrame and s.TextFrame.TextRange.Text == "3")
    assert number.ZOrderPosition > video.ZOrderPosition  # номер бочонка поверх видео
