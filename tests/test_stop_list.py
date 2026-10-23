from openpyxl import Workbook

from scripts.stop_list import blocked_artist, load_stop_list
from scripts.validate_selection import validate_selection

STOP = ["LOBODA", "Моргенштерн / MORGENSHTERN"]


def test_artist_found_in_any_spelling_and_in_duets():
    assert blocked_artist("Loboda - Americano", STOP) == "LOBODA"
    assert blocked_artist("MORGENSHTERN, Элджей - Cadillac", STOP) == "Моргенштерн / MORGENSHTERN"
    assert blocked_artist("Элджей, MORGENSHTERN - Cadillac", STOP) == "Моргенштерн / MORGENSHTERN"
    assert blocked_artist("Моргенштерн feat. Кто-то - Песня", STOP) == "Моргенштерн / MORGENSHTERN"


def test_other_artists_and_song_titles_are_not_blocked():
    assert blocked_artist("Элджей - Розовое вино", STOP) is None
    assert blocked_artist("Ирина Аллегрова - Лобода", STOP) is None  # совпадение в названии песни — не в счёт
    assert blocked_artist("Lobodanov - Песня", STOP) is None  # только целое имя


def test_validate_selection_rejects_stop_list_songs():
    songs = [f"Исполнитель {i} - Песня" for i in range(89)] + ["LOBODA - SuperSTAR"]
    bank = [{"raw": s} for s in songs]
    result = validate_selection(songs, bank, stop_list=STOP)
    assert not result["ok"] and "LOBODA - SuperSTAR" in " ".join(result["errors"])
    assert validate_selection(songs, bank)["ok"]  # без стоп-листа — как раньше


def test_load_stop_list_reads_first_column_without_header(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["Исполнитель", "Причина"])
    ws.append(["LOBODA", "иноагент"])
    ws.append([None, None])
    ws.append(["  MORGENSHTERN ", ""])
    path = tmp_path / "стоп.xlsx"
    wb.save(path)
    assert load_stop_list(str(path)) == ["LOBODA", "MORGENSHTERN"]
    assert load_stop_list(str(tmp_path / "нет.xlsx")) == []
