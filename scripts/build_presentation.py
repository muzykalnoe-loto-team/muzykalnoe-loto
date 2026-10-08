import os
import shutil

import win32com.client as win32
from pptx import Presentation

from scripts import paths
from scripts.old_slot import fill_old_slot, insert_photos
from scripts.ppt_paste import insert_song_slide
from scripts.song_names import split_raw
from scripts.template_slots import find_slide_index_by_barrel_number


def assemble_presentation(
    template_path: str,
    archive_folder: str,
    barrels: list[dict],
    output_path: str,
    old_bank_folder: str = paths.OLD_BANK,
) -> None:
    # 1. Все файлы песен на месте — до запуска PowerPoint
    resolved = []
    for barrel in barrels:
        folder = old_bank_folder if barrel.get("kind") == "old" else archive_folder
        song_file = os.path.join(folder, barrel.get("file") or f"{barrel['song']}.pptx")
        if not os.path.exists(song_file):
            raise FileNotFoundError(f"Не найден файл для бочонка {barrel['number']}: {barrel['song']}")
        resolved.append((barrel, os.path.abspath(song_file)))

    # 2. Номера слайдов-слотов — без PowerPoint, быстрым чтением через python-pptx
    prs_readonly = Presentation(template_path)
    slot = {b["number"]: find_slide_index_by_barrel_number(prs_readonly, b["number"]) + 1 for b, _ in resolved}

    # 3. Работаем на копии, чтобы не трогать оригинал шаблона; фото старых песен — в рамки
    #    шаблона до PowerPoint (python-pptx)
    output_path = os.path.abspath(output_path)
    shutil.copy(template_path, output_path)
    insert_photos(output_path, {
        slot[b["number"]]: (f, b["photo_part"])
        for b, f in resolved if b.get("kind") == "old" and b.get("photo_part")
    })

    # Dispatch (позднее связывание): EnsureDispatch падает на битом кэше gen_py
    ppt_app = win32.Dispatch("PowerPoint.Application")
    try:
        template_prs = ppt_app.Presentations.Open(output_path, WithWindow=False)
        try:
            for barrel, song_file in resolved:
                if barrel.get("kind") == "old":
                    artist, title = split_raw(barrel["song"])
                    fill_old_slot(ppt_app, template_prs, slot[barrel["number"]], song_file, title, artist)
                else:
                    insert_song_slide(ppt_app, template_prs, song_file, slot[barrel["number"]])
            template_prs.Save()
        finally:
            template_prs.Close()
    finally:
        # Оператор может держать открытыми свои презентации в том же PowerPoint —
        # закрываем его, только если других файлов не открыто
        if ppt_app.Presentations.Count == 0:
            ppt_app.Quit()
