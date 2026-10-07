import os

MSO_BRING_TO_FRONT = 0


def insert_song_slide(ppt_app, template_prs, archive_file_path: str, slide_index_1based: int) -> None:
    # Копируем все фигуры слайда песни и вставляем поверх слота: сам слот не заменяем,
    # иначе пропадут вшитые гиперссылки навигации
    archive_prs = ppt_app.Presentations.Open(
        os.path.abspath(archive_file_path), ReadOnly=True, WithWindow=False
    )
    try:
        archive_shapes = archive_prs.Slides(1).Shapes
        positions = [(s.Left, s.Top) for s in archive_shapes]
        archive_shapes.Range().Copy()

        target_slide = template_prs.Slides(slide_index_1based)
        own_shapes = list(target_slide.Shapes)
        pasted = target_slide.Shapes.Paste()
    finally:
        archive_prs.Close()

    # PowerPoint сдвигает вставку на 12 pt — возвращаем фигуры на места из архива
    for i, (left, top) in enumerate(positions, start=1):
        pasted(i).Left = left
        pasted(i).Top = top

    # Видео на весь слайд перекрыло бы кнопку навигации шаблона — поднимаем фигуры слота наверх
    for shape in own_shapes:
        shape.ZOrder(MSO_BRING_TO_FRONT)
