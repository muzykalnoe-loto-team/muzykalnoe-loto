# Номинации в собранной игре (PowerPoint COM): по одной на раунд, лишние варианты удаляются,
# кнопки 1–3 «Таблицы» и все ссылки на удалённые варианты ведут на оставшуюся номинацию раунда.
# Здесь же — превью вариантов тематического слайда и архив «Номинация» (Claude его не читает)
import os
import re
import shutil
import zipfile

from scripts import paths
from scripts.nominations import ROUNDS, THEME
from scripts.theme_art import render_scene

PP_HYPERLINK = 7
MSO_PICTURE = 13
MSO_PLACEHOLDER = 14
MSO_SEND_TO_BACK = 1
MSO_TEXT_HORIZONTAL = 1
TABLE_LAYOUT = "Таблица"
TABLE_BUTTONS = {1: 505, 2: 594, 3: 679}  # кнопки «1», «2», «3» под сеткой «Таблицы» (Left, pt)
BUTTONS_TOP = 450
NOMINATION_LINE = "НОМИНАЦИЯ «САМАЯ КОСТЮМИРОВАННАЯ КОМАНДА»"
NOMINATION_FONT = ("Unbounded", 15)


def _rgb(hex_color: str) -> int:
    # PowerPoint хранит цвет как 0xBBGGRR
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (1, 3, 5))
    return r + g * 256 + b * 65536


def _title(slide):
    return next(s for s in slide.Shapes if s.Type == MSO_PLACEHOLDER and s.HasTextFrame)


def fill_theme_slide(slide, title: str, glow: str, art_png: str) -> None:
    width, height = slide.Parent.PageSetup.SlideWidth, slide.Parent.PageSetup.SlideHeight
    # картинка «рок» на весь слайд → иллюстрация темы, под заголовком и кнопкой в углу
    for shape in [s for s in slide.Shapes if s.Type == MSO_PICTURE and s.Left < 1 and s.Top < 1
                  and s.Width >= width - 1]:
        shape.Delete()
    art = slide.Shapes.AddPicture(art_png, False, True, 0, 0, width, height)
    art.Name = "Тематический рисунок"
    art.ZOrder(MSO_SEND_TO_BACK)

    head = _title(slide)
    head.TextFrame.TextRange.Text = title
    line = slide.Shapes.AddTextbox(MSO_TEXT_HORIZONTAL, head.Left + 11, head.Top + head.Height + 4, 640, 30)
    line.Name = "Номинация"
    line.TextFrame.WordWrap = True
    text = line.TextFrame.TextRange
    text.Text = NOMINATION_LINE
    text.Font.Name, text.Font.Size = NOMINATION_FONT
    text.Font.Bold = False
    text.Font.Color.RGB = _rgb(glow)


def title_lines(slide) -> int:
    return _title(slide).TextFrame2.TextRange.Lines.Count


def _link(shape):
    try:
        action = shape.ActionSettings(1)
    except Exception:  # у групп нет действия
        return None
    if action.Action != PP_HYPERLINK or not action.Hyperlink.SubAddress:
        return None
    return action.Hyperlink


def _point(hyperlink, slide) -> None:
    hyperlink.SubAddress = f"{slide.SlideID},{slide.SlideIndex},"


def apply_nominations(prs, chosen: dict[int, str], theme: dict | None) -> None:
    keep = {r: ROUNDS[r][chosen[r]]["slide_id"] for r in ROUNDS}
    if chosen[1] == THEME:
        fill_theme_slide(prs.Slides.FindBySlideID(keep[1]), theme["title"], theme["glow"], theme["art_png"])

    # удалённый вариант → оставшаяся номинация того же раунда
    replace = {v["slide_id"]: keep[r] for r in ROUNDS for v in ROUNDS[r].values() if v["slide_id"] != keep[r]}
    for slide_id in replace:
        prs.Slides.FindBySlideID(slide_id).Delete()

    slides = [prs.Slides(i) for i in range(1, prs.Slides.Count + 1)]
    by_id = {s.SlideID: s for s in slides}
    # номера слайдов сдвинулись: обновляем все ссылки на слайды (и в макетах), заменяя удалённые варианты
    layouts = {(s.Design.Name, s.CustomLayout.Name): s.CustomLayout for s in slides}.values()
    for owner in slides + list(layouts):
        for shape in owner.Shapes:
            hyperlink = _link(shape)
            if hyperlink is None:
                continue
            head = hyperlink.SubAddress.split(",")[0]
            if head.isdigit():
                target = replace.get(int(head), int(head))
                if target in by_id:
                    _point(hyperlink, by_id[target])

    # кнопки 1–3 «Таблицы» в шаблоне ведут не туда — каждая на номинацию своего раунда
    table = next(s for s in slides if s.CustomLayout.Name == TABLE_LAYOUT)
    for shape in table.Shapes:
        hyperlink = _link(shape)
        for round_no, left in TABLE_BUTTONS.items():
            if hyperlink and abs(shape.Left - left) < 5 and shape.Top > BUTTONS_TOP:
                _point(hyperlink, by_id[keep[round_no]])


THEME_SLIDE_ID = ROUNDS[1][THEME]["slide_id"]
TITLE_LINES = 2


def _powerpoint():
    import win32com.client as win32
    return win32.Dispatch("PowerPoint.Application")


def _quit_if_idle(app) -> None:
    if app.Presentations.Count == 0:
        app.Quit()


def _keep_only(prs, slide_id: int) -> None:
    for i in range(prs.Slides.Count, 0, -1):
        if prs.Slides(i).SlideID != slide_id:
            prs.Slides(i).Delete()


def preview_variants(variants: list[dict], out_dir: str) -> dict:
    # файл с вариантами тематического слайда (по слайду на вариант) и их картинки — оператору на выбор
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    pptx = os.path.join(out_dir, "Номинация — варианты.pptx")
    shutil.copy(paths.PRESENTATION_TEMPLATE, pptx)
    arts = [render_scene(v["scene"], os.path.join(out_dir, f"Номинация — вариант {n}, рисунок.png"), v["seed"])
            for n, v in enumerate(variants, start=1)]
    pngs, warnings = [], []
    app = _powerpoint()
    prs = app.Presentations.Open(pptx, WithWindow=False)
    try:
        _keep_only(prs, THEME_SLIDE_ID)
        base = prs.Slides(1)
        slides = [base] + [base.Duplicate()(1) for _ in variants[1:]]  # первый вариант — на самом слайде 374
        for n, (slide, v, art) in enumerate(zip(slides, variants, arts), start=1):
            slide.MoveTo(n)
            fill_theme_slide(slide, v["title"], v["glow"], art)
            lines = title_lines(slide)
            if lines != TITLE_LINES:
                warnings.append(f"Вариант {n}: заголовок занимает {lines} строк(и) вместо {TITLE_LINES} — сократите")
            png = os.path.join(out_dir, f"Номинация — вариант {n}.png")
            slide.Export(png, "PNG", 1280, 720)
            pngs.append(png)
        prs.Save()
    finally:
        prs.Close()
        _quit_if_idle(app)
    return {"pptx": pptx, "pngs": pngs, "arts": arts, "warnings": warnings}


def theme_slide_file(theme: dict, out_pptx: str) -> str:
    # файл с одним готовым тематическим слайдом (SlideID 374) — для архива
    out_pptx = os.path.abspath(out_pptx)
    shutil.copy(paths.PRESENTATION_TEMPLATE, out_pptx)
    app = _powerpoint()
    prs = app.Presentations.Open(out_pptx, WithWindow=False)
    try:
        _keep_only(prs, THEME_SLIDE_ID)
        fill_theme_slide(prs.Slides(1), theme["title"], theme["glow"], os.path.abspath(theme["art_png"]))
        prs.Save()
    finally:
        prs.Close()
        _quit_if_idle(app)
    return out_pptx


def _slide_index(pptx: str, slide_id: int) -> int:
    # номер слайда по SlideID без открытия файла игры целиком (он бывает почти 1 ГБ)
    with zipfile.ZipFile(pptx) as z:
        ids = [int(x) for x in re.findall(r'<p:sldId\b[^>]*?\sid="(\d+)"', z.read("ppt/presentation.xml").decode("utf-8"))]
    return ids.index(slide_id) + 1


def _new_archive(app, archive: str) -> None:
    # пустой архив с оформлением шаблона: макет «Тематическая» и шрифты те же
    shutil.copy(paths.PRESENTATION_TEMPLATE, archive)
    prs = app.Presentations.Open(archive, WithWindow=False)
    try:
        _keep_only(prs, -1)
        prs.Save()
    finally:
        prs.Close()


def append_to_archive(game_pptx: str, event: str, date: str, theme_name: str,
                      archive: str = paths.NOMINATION_ARCHIVE) -> str:
    # дописывает тематический слайд игры в конец архива; содержимое архива не читается и не выводится
    archive, game_pptx = os.path.abspath(archive), os.path.abspath(game_pptx)
    app = _powerpoint()
    try:
        if any(app.Presentations(i).FullName.lower() == archive.lower()
               for i in range(1, app.Presentations.Count + 1)):
            return f"Архив номинаций открыт в PowerPoint — слайд не дописан. Закройте {archive} и допишите позже"
        if not os.path.exists(archive):
            _new_archive(app, archive)
        index = _slide_index(game_pptx, THEME_SLIDE_ID)
        prs = app.Presentations.Open(archive, WithWindow=False)
        try:
            prs.Slides.InsertFromFile(game_pptx, prs.Slides.Count, index, index)
            notes = prs.Slides(prs.Slides.Count).NotesPage.Shapes.Placeholders(2)
            notes.TextFrame.TextRange.Text = f"{event} — {date} — {theme_name}"
            prs.Save()
        finally:
            prs.Close()
    except Exception as error:  # архив занят другой программой и т. п. — игре это не мешает
        return f"Архив номинаций не дополнен: {error}"
    finally:
        _quit_if_idle(app)
    return ""


def seed_archive() -> None:
    # разовая закладка пяти одобренных образцов в архив
    import tempfile

    from scripts.theme_scenes import THEMES
    with tempfile.TemporaryDirectory() as tmp:
        for key, theme in THEMES.items():
            v = {"title": theme["title"], "glow": theme["glow"], "scene": theme["scene"], "seed": key}
            res = preview_variants([v], os.path.join(tmp, key))
            warning = append_to_archive(res["pptx"], "Образец", "2026-10-17", theme["name"])
            if warning:
                raise RuntimeError(warning)


if __name__ == "__main__":
    seed_archive()
