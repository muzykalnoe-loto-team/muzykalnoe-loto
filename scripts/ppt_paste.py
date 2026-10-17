import os
import time

MSO_BRING_TO_FRONT = 0
PP_MOUSE_CLICK = 1
PP_ACTION_NONE = 0
PASTE_ATTEMPTS = 3
MSO_MEDIA = 16


def insert_song_slide(ppt_app, template_prs, archive_file_path: str, slide_index_1based: int) -> None:
    # Копируем все фигуры слайда песни и вставляем поверх слота: сам слот не заменяем,
    # иначе пропадут вшитые гиперссылки навигации
    archive_prs = ppt_app.Presentations.Open(
        os.path.abspath(archive_file_path), ReadOnly=True, WithWindow=False
    )
    try:
        archive_shapes = archive_prs.Slides(1).Shapes
        # видео песни на весь слайд — по нему меряем сдвиг вставки (обычно первая фигура, но не всегда)
        video = next(s for s in archive_shapes if s.Type == MSO_MEDIA)
        anchor = (video.Left, video.Top)
        video_name = video.Name

        target_slide = template_prs.Slides(slide_index_1based)
        own_shapes = list(target_slide.Shapes)
        pasted = copy_paste_checked(archive_shapes.Range().Copy, target_slide, video.Name, video.Type)
    finally:
        archive_prs.Close()

    # PowerPoint сдвигает вставку на 12 pt — возвращаем фигуры на места из архива.
    # Пустые текстовые поля архива не вставляются, поэтому сопоставлять по порядку нельзя
    pasted_video = _find(pasted, video_name, MSO_MEDIA)
    dx = pasted_video.Left - anchor[0]
    dy = pasted_video.Top - anchor[1]
    for i in range(1, pasted.Count + 1):
        pasted(i).Left -= dx
        pasted(i).Top -= dy

    # На слайдах архива в поле номера часто остался старый номер ("1", "2") — он наложился бы
    # на номер бочонка из шаблона, поэтому стираем его; номер показывает фигура шаблона
    for i in range(1, pasted.Count + 1):
        if _is_barrel_number(pasted(i)):
            pasted(i).TextFrame.TextRange.Text = ""

    # Видео на весь слайд перекрыло бы кнопку навигации и номер бочонка шаблона — поднимаем
    # наверх только их. Заглушки шаблона (рамка под картинку, "ТЕКСТ НА 2 СТРОЧКИ") остаются
    # под видео, как при ручной вставке
    for shape in own_shapes:
        if _has_click_action(shape) or _is_barrel_number(shape):
            shape.ZOrder(MSO_BRING_TO_FRONT)

    # На слотах «Игровой слайд» (бочонки 3, 6 … 90) кнопка «Таблица» живёт в макете, а макет всегда
    # под фигурами слайда — видео закрыло бы её. Ставим на слайд поверх видео такую же невидимую
    # кнопку с тем же переходом (через буфер обмена нельзя: при вставке кнопка переименовывается)
    if not any(_has_click_action(s) for s in own_shapes):
        for button in [s for s in target_slide.CustomLayout.Shapes if _has_click_action(s)]:
            copy = target_slide.Shapes.AddShape(button.AutoShapeType, button.Left, button.Top, button.Width, button.Height)
            copy.Name = button.Name
            copy.Fill.Visible = False
            copy.Line.Visible = False
            action, src = copy.ActionSettings(PP_MOUSE_CLICK), button.ActionSettings(PP_MOUSE_CLICK)
            action.Action = src.Action
            action.Hyperlink.Address = src.Hyperlink.Address
            action.Hyperlink.SubAddress = src.Hyperlink.SubAddress


def _has_click_action(shape) -> bool:
    try:
        return shape.ActionSettings(PP_MOUSE_CLICK).Action != PP_ACTION_NONE
    except Exception:  # у групп нет настроек действия
        return False


def _is_barrel_number(shape) -> bool:
    return shape.HasTextFrame and shape.TextFrame.TextRange.Text.strip().isdigit()


def _find(shapes, name: str, type_: int):
    return next((shapes(i) for i in range(1, shapes.Count + 1)
                 if shapes(i).Name == name and shapes(i).Type == type_), None)


def copy_paste_checked(copy, target_slide, expected_name: str, expected_type: int):
    # Буфер обмена иногда не успевает обновиться, и вставляется предыдущая песня
    # (картинкой вместо видео или звука) или ещё пуст — PowerPoint тогда ругается
    # «Clipboard is empty»; сверяем вставленное и повторяем
    for attempt in range(1, PASTE_ATTEMPTS + 1):
        copy()
        time.sleep(0.2 * attempt)
        try:
            pasted = target_slide.Shapes.Paste()
        except Exception:  # буфер ещё пуст — следующая попытка
            continue
        if _find(pasted, expected_name, expected_type) is not None:
            return pasted
        pasted.Delete()
    raise RuntimeError(f"Не удалось вставить «{expected_name}» за {PASTE_ATTEMPTS} попытки")
