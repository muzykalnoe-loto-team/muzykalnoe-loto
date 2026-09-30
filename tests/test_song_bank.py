import openpyxl
from scripts.song_bank import load_song_bank


def make_bank_fixture(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "2301"
    rows = [
        (1, "Юта - Белая вьюга"),
        (2, "#2Маши - Босая"),
        (3, "5sta Family - Я буду"),
    ]
    for r in rows:
        ws.append(r)
    wb.save(path)


def test_load_song_bank_parses_artist_and_title(tmp_path):
    fixture = tmp_path / "bank.xlsx"
    make_bank_fixture(fixture)

    result = load_song_bank(str(fixture))

    assert result == [
        {"n": 1, "raw": "Юта - Белая вьюга", "artist": "Юта", "title": "Белая вьюга"},
        {"n": 2, "raw": "#2Маши - Босая", "artist": "#2Маши", "title": "Босая"},
        {"n": 3, "raw": "5sta Family - Я буду", "artist": "5sta Family", "title": "Я буду"},
    ]


def test_load_song_bank_skips_empty_rows(tmp_path):
    fixture = tmp_path / "bank_with_gap.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "2301"
    ws.append((1, "Юта - Белая вьюга"))
    ws.append((2, None))
    wb.save(fixture)

    result = load_song_bank(str(fixture))

    assert result == [
        {"n": 1, "raw": "Юта - Белая вьюга", "artist": "Юта", "title": "Белая вьюга"},
    ]
