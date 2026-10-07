import os
import shutil

import win32com.client as win32
from pptx import Presentation

from scripts.ppt_paste import insert_song_slide
from scripts.template_slots import find_slide_index_by_barrel_number


def assemble_presentation(
    template_path: str,
    archive_folder: str,
    barrels: list[dict],
    output_path: str,
) -> None:
    # 1. Быстрая проверка наличия всех файлов архива до запуска PowerPoint
    resolved = []
    for barrel in barrels:
        archive_file = os.path.join(archive_folder, f"{barrel['song']}.pptx")
        if not os.path.exists(archive_file):
            raise FileNotFoundError(
                f"Не найден файл в архиве для бочонка {barrel['number']}: {barrel['song']}"
            )
        resolved.append((barrel["number"], os.path.abspath(archive_file)))

    # 2. Номера слайдов-слотов — без PowerPoint, быстрым чтением через python-pptx
    prs_readonly = Presentation(template_path)
    slot_indices = {
        number: find_slide_index_by_barrel_number(prs_readonly, number) + 1  # 1-based для COM
        for number, _ in resolved
    }

    # 3. Работаем на копии, чтобы не трогать оригинал шаблона
    output_path = os.path.abspath(output_path)
    shutil.copy(template_path, output_path)

    # Dispatch (позднее связывание): EnsureDispatch падает на битом кэше gen_py
    ppt_app = win32.Dispatch("PowerPoint.Application")
    try:
        template_prs = ppt_app.Presentations.Open(output_path, WithWindow=False)
        try:
            for number, archive_file in resolved:
                insert_song_slide(ppt_app, template_prs, archive_file, slot_indices[number])
            template_prs.Save()
        finally:
            template_prs.Close()
    finally:
        # Оператор может держать открытыми свои презентации в том же PowerPoint —
        # закрываем его, только если других файлов не открыто
        if ppt_app.Presentations.Count == 0:
            ppt_app.Quit()
