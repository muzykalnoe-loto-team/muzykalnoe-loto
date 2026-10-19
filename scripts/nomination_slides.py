# Номинации в собранной игре (PowerPoint COM): по одной на раунд, лишние варианты удаляются,
# кнопки 1–3 «Таблицы» и все ссылки на удалённые варианты ведут на оставшуюся номинацию раунда
from scripts.nominations import ROUNDS, THEME

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
