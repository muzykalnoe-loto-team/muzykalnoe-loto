from scripts.validate_selection import validate_selection

BANK = [
    {"n": i, "raw": f"Артист{i} - Песня{i}", "artist": f"Артист{i}", "title": f"Песня{i}"}
    for i in range(1, 101)
]


def test_valid_selection_of_90_passes():
    selected = [f"Артист{i} - Песня{i}" for i in range(1, 91)]

    result = validate_selection(selected, BANK)

    assert result == {"ok": True, "errors": []}


def test_wrong_count_fails():
    selected = [f"Артист{i} - Песня{i}" for i in range(1, 90)]  # 89 штук

    result = validate_selection(selected, BANK)

    assert result["ok"] is False
    assert any("90" in e for e in result["errors"])


def test_duplicate_fails():
    selected = [f"Артист{i} - Песня{i}" for i in range(1, 90)] + ["Артист1 - Песня1"]

    result = validate_selection(selected, BANK)

    assert result["ok"] is False
    assert any("дубл" in e.lower() for e in result["errors"])


def test_unknown_song_fails():
    selected = [f"Артист{i} - Песня{i}" for i in range(1, 90)] + ["Неизвестный - Трек"]

    result = validate_selection(selected, BANK)

    assert result["ok"] is False
    assert any("Неизвестный - Трек" in e for e in result["errors"])


def test_song_with_other_spelling_is_known():
    # в общем банке разные написания одной песни слиты: «ё/е», регистр, знаки не мешают
    bank = BANK + [{"raw": "Ёлка - Прованс"}]
    selected = [f"Артист{i} - Песня{i}" for i in range(1, 90)] + ["Елка - прованс"]

    assert validate_selection(selected, bank) == {"ok": True, "errors": []}
