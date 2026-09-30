import openpyxl


def load_song_bank(path: str) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]

    songs = []
    for row in ws.iter_rows(min_row=1, values_only=True):
        n, raw = row[0], row[1]
        if raw is None or str(raw).strip() == "":
            continue
        raw = str(raw).strip()
        artist, _, title = raw.partition(" - ")
        songs.append({
            "n": int(n),
            "raw": raw,
            "artist": artist.strip(),
            "title": title.strip(),
        })
    return songs
