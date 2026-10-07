import os
import re
import threading
import time

import win32com.client as win32
import win32con
import win32gui

XL_MACRO_ENABLED_WORKBOOK = 52  # xlOpenXMLWorkbookMacroEnabled


def _normalize_separator(song: str) -> str:
    # Формула шаблона ищет " - ": у части песен в базе пробел с одной стороны
    # ("Macan- Пополам", "Шатунов -А лето"), формула даёт #VALUE! и макрос падает
    if " - " in song:
        return song
    return re.sub(r"(?<=\S)-\s+|\s+-(?=\S)", " - ", song, count=1)


def _dialog_text(hwnd: int) -> str:
    parts = []

    def collect(child, _):
        if win32gui.GetClassName(child) == "Static":
            text = win32gui.GetWindowText(child)
            if text:
                parts.append(text)

    win32gui.EnumChildWindows(hwnd, collect, None)
    return " ".join(parts)


def _dismiss_excel_dialogs(stop: threading.Event, seen: list[str]) -> None:
    # Макрос заканчивается MsgBox "Готово!", который ждёт нажатия ОК — закрываем его сами
    while not stop.is_set():
        hwnd = win32gui.FindWindow("#32770", "Microsoft Excel")
        if hwnd:
            seen.append(_dialog_text(hwnd))
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        time.sleep(0.3)


def generate_blanks(
    template_path: str,
    output_path: str,
    selected_songs: list[str],
    n_individual: int,
    n_team: int,
) -> None:
    if len(selected_songs) != 90:
        raise ValueError(f"Ожидается ровно 90 песен, передано {len(selected_songs)}")

    template_path = os.path.abspath(template_path)
    output_path = os.path.abspath(output_path)

    # DispatchEx (позднее связывание, отдельный процесс Excel): EnsureDispatch падает на битом
    # кэше gen_py, а общий Excel спрятал бы и закрыл таблицы, открытые у оператора
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    stop = threading.Event()
    dialogs: list[str] = []
    watcher = threading.Thread(target=_dismiss_excel_dialogs, args=(stop, dialogs), daemon=True)
    watcher.start()
    try:
        wb = excel.Workbooks.Open(template_path)
        try:
            ws_list = wb.Sheets("Список")
            for i, song in enumerate(selected_songs):
                ws_list.Cells(i + 2, 2).Value = _normalize_separator(song)  # B2:B91

            ws_list.Range("E1").Value = n_individual
            ws_list.Range("E2").Value = n_team

            excel.Run("GenerateBlanks.GenerateAllBlanks")

            errors = [d for d in dialogs if "Ошибка" in d]
            if errors:
                raise RuntimeError("Макрос сообщил об ошибке: " + "; ".join(errors))

            wb.SaveAs(output_path, FileFormat=XL_MACRO_ENABLED_WORKBOOK)
        finally:
            wb.Close(SaveChanges=False)
    finally:
        stop.set()
        excel.Quit()
