from scripts.common_bank import load_common_bank, merge_banks, save_common_bank


def test_same_song_spelled_differently_is_one_entry_with_both_kinds():
    video = ["Наташа Королёва - Жёлтые тюльпаны", "Жуки - Батарейка"]
    files = {"Наташа Королёва - Жёлтые тюльпаны.pptx", "Жуки - Батарейка.pptx"}
    old = [
        {"raw": "Наташа Королева - Желтые тюльпаны", "file": "Наташа Королева - Желтые тюльпаны.pptx",
         "photo_part": "ppt/media/image3.png", "source": "Т/8 марта/1.pptx", "slide_part": "ppt/slides/slide9.xml", "check": ""},
        {"raw": "Звери - Танцуй", "file": "Звери - Танцуй.pptx", "photo_part": None,
         "source": "Т/Хэллоуин/1.pptx", "slide_part": "ppt/slides/slide7.xml", "check": "не узнана по своду — проверьте написание"},
    ]

    bank = {e["raw"]: e for e in merge_banks(video, files, old)}

    assert len(bank) == 3
    korol = bank["Наташа Королёва - Жёлтые тюльпаны"]  # написание видео-банка главное
    assert korol["video_file"] == "Наташа Королёва - Жёлтые тюльпаны.pptx"
    assert korol["old_file"] == "Наташа Королева - Желтые тюльпаны.pptx"
    assert bank["Жуки - Батарейка"]["old_file"] is None
    assert bank["Звери - Танцуй"]["video_file"] is None
    assert bank["Звери - Танцуй"]["check"]


def test_old_slide_without_file_is_not_offered():
    # песня каталога без нарезанного файла (ждёт проверки) не считается доступной старым слайдом
    old = [{"raw": "Звери - Танцуй", "file": None, "photo_part": None, "source": "a", "slide_part": "b", "check": "x"}]

    assert merge_banks([], set(), old) == []


def test_save_and_load_roundtrip(tmp_path):
    entries = merge_banks(["Жуки - Батарейка"], {"Жуки - Батарейка.pptx"}, [])

    save_common_bank(entries, str(tmp_path / "банк.xlsx"), str(tmp_path / "банк.json"))

    assert load_common_bank(str(tmp_path / "банк.json")) == entries
