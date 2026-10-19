from scripts.barrel_kinds import bank_index, check_kinds, resolve_files

BANK = bank_index([
    {"raw": "Видео - Песня", "video_file": "Видео - Песня.pptx", "old_file": None, "old_photo_part": None},
    {"raw": "Старая - Песня", "video_file": None, "old_file": "Старая - Песня.pptx",
     "old_photo_part": "ppt/media/image2.png"},
    {"raw": "Обе - Песня", "video_file": "Обе - Песня.pptx", "old_file": "Обе - Песня.pptx", "old_photo_part": None},
    {"raw": "Ёлка - Прованс", "video_file": "Ёлка - Прованс.pptx", "old_file": None, "old_photo_part": None},
])


def _game(kinds_by_number):
    return [{"number": n, "song": song, "kind": kind} for n, (song, kind) in kinds_by_number.items()]


def test_mixed_game_needs_old_slides_on_every_third_barrel():
    good = _game({1: ("Видео - Песня", "video"), 3: ("Старая - Песня", "old"), 6: ("Обе - Песня", "old")})
    assert check_kinds(good, BANK) == []

    bad = _game({3: ("Старая - Песня", "old"), 4: ("Обе - Песня", "old")})
    errors = check_kinds(bad, BANK)
    assert len(errors) == 1 and "Бочонок 4" in errors[0]


def test_old_kind_requires_old_slide_file():
    errors = check_kinds(_game({3: ("Видео - Песня", "old")}), BANK)
    assert errors and "нет старого слайда" in errors[0]


def test_every_third_barrel_must_be_old_even_in_all_video_list():
    # список только из видео (как до старых слайдов) больше не собирается: на местах 3, 6 … 90 — старые слайды
    legacy = [{"number": 1, "song": "Видео - Песня"}, {"number": 3, "song": "Обе - Песня"}]
    errors = check_kinds(legacy, BANK)
    assert len(errors) == 1 and "Бочонок 3" in errors[0] and "старый слайд" in errors[0]


def test_barrel_without_kind_is_video():
    legacy = [{"number": 1, "song": "Видео - Песня"}, {"number": 3, "song": "Обе - Песня", "kind": "old"}]
    assert check_kinds(legacy, BANK) == []
    assert [b["kind"] for b in resolve_files(legacy, BANK)] == ["video", "old"]


def test_resolve_files_takes_file_of_the_right_kind():
    resolved = resolve_files(_game({3: ("Старая - Песня", "old"), 1: ("Видео - Песня", "video")}), BANK)
    assert resolved[0]["file"] == "Старая - Песня.pptx" and resolved[0]["photo_part"] == "ppt/media/image2.png"
    assert resolved[1]["file"] == "Видео - Песня.pptx"


def test_song_found_despite_different_spelling():
    # в общем банке разные написания одной песни слиты в одну строку (ё/е, регистр)
    legacy = [{"number": 1, "song": "Елка - прованс"}]
    assert check_kinds(legacy, BANK) == []
    assert resolve_files(legacy, BANK)[0]["file"] == "Ёлка - Прованс.pptx"
    assert resolve_files(legacy, BANK)[0]["song"] == "Елка - прованс"  # в бланках — как в списке
