import os
import time

MSO_BRING_TO_FRONT = 0
PP_MOUSE_CLICK = 1
PP_ACTION_NONE = 0
PASTE_ATTEMPTS = 3


def insert_song_slide(ppt_app, template_prs, archive_file_path: str, slide_index_1based: int) -> None:
    # Копируем все фигуры слайда песни и вставляем поверх слота: сам слот не заменяем,
    # иначе пропадут вшитые гиперссылки навигации
    archive_prs = ppt_app.Presentations.Open(
        os.path.abspath(archive_file_path), ReadOnly=True, WithWindow=False
    )
    try:
        archive_shapes = archive_prs.Slides(1).Shapes
        # первая фигура — видео песни на весь слайд, по нему меряем сдвиг вставки
        video = archive_shapes(1)
        anchor = (video.Left, video.Top)

        target_slide = template_prs.Slides(slide_index_1based)
        own_shapes = list(target_slide.Shapes)
        pasted = copy_paste_checked(archive_shapes.Range().Copy, target_slide, video.Name, video.Type)
    finally:
        archive_prs.Close()

    # PowerPoint сдвигает вставку на 12 pt — возвращаем фигуры на места из архива.
    # Пустые текстовые поля архива не вставляются, поэтому сопоставлять по порядку нельзя
    dx = pasted(1).Left - anchor[0]
    dy = pasted(1).Top - anchor[1]
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


def _has_click_action(shape) -> bool:
    try:
        return shape.ActionSettings(PP_MOUSE_CLICK).Action != PP_ACTION_NONE
    except Exception:  # у групп нет настроек действия
        return False


def _is_barrel_number(shape) -> bool:
    return shape.HasTextFrame and shape.TextFrame.TextRange.Text.strip().isdigit()


def copy_paste_checked(copy, target_slide, expected_name: str, expected_type: int):
    # Буфер обмена иногда не успевает обновиться, и вставляется предыдущая песня
    # (картинкой вместо видео или звука) — сверяем вставленное и повторяем
    for attempt in range(1, PASTE_ATTEMPTS + 1):
        copy()
        time.sleep(0.2 * attempt)
        pasted = target_slide.Shapes.Paste()
        if pasted(1).Name == expected_name and pasted(1).Type == expected_type:
            return pasted
        pasted.Delete()
    raise RuntimeError(f"Не удалось вставить «{expected_name}» за {PASTE_ATTEMPTS} попытки")
