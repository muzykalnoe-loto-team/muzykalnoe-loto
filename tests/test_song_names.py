from scripts.song_names import make_raw, norm, safe_filename, split_raw


def test_norm_ignores_case_yo_spaces_and_punctuation():
    assert norm("Наташа Королёва - Жёлтые тюльпаны") == norm("наташа  королева – желтые тюльпаны!")
    assert norm("ANNA ASTI,  Филипп Киркоров") == norm("ANNA ASTI, Филипп Киркоров")


def test_split_raw_uses_first_separator():
    assert split_raw("Ю-Питер - Девушка по городу") == ("Ю-Питер", "Девушка по городу")
    assert split_raw("Macan- Пополам") == ("Macan", "Пополам")


def test_make_raw_squeezes_spaces():
    assert make_raw(" Звери ", "Просто  такая сильная любовь") == "Звери - Просто такая сильная любовь"


def test_safe_filename_replaces_forbidden_characters():
    assert safe_filename("5sta Family - Зачем?") == "5sta Family - Зачем_"
    assert safe_filename('A/B - "C"...') == "A_B - _C_"
