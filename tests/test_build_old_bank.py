from scripts.build_old_bank import assign_file_names, songs_to_cut


def test_file_names_are_safe_and_unique():
    catalog = [{"raw": "5sta Family - Зачем?"}, {"raw": "5sta Family - Зачем_"}, {"raw": "Звери - Танцуй"}]

    assign_file_names(catalog)

    files = [c["file"] for c in catalog]
    assert files[0] == "5sta Family - Зачем_.pptx"
    assert len({f.lower() for f in files}) == 3  # разные песни не затирают друг друга
    assert files[2] == "Звери - Танцуй.pptx"


def test_songs_waiting_for_check_are_not_cut():
    # решение оператора 2026-10-08: песни «проверить» режем только после его сверки
    catalog = [{"raw": "А - 1", "check": ""}, {"raw": "Б - 2", "check": "не узнана по своду — проверьте написание"}]

    assert [c["raw"] for c in songs_to_cut(catalog)] == ["А - 1"]


def test_existing_file_is_recut_unless_it_has_exactly_one_slide(tmp_path):
    import zipfile

    from scripts.build_old_bank import is_cut_well

    good, bad, broken = tmp_path / "1.pptx", tmp_path / "2.pptx", tmp_path / "3.pptx"
    with zipfile.ZipFile(good, "w") as z:
        z.writestr("ppt/slides/slide7.xml", "<x/>")
    with zipfile.ZipFile(bad, "w") as z:  # нарезка с ошибкой: в файл попала вся игра
        z.writestr("ppt/slides/slide1.xml", "<x/>")
        z.writestr("ppt/slides/slide7.xml", "<x/>")
    broken.write_bytes(b"not a zip")

    assert is_cut_well(str(good))
    assert not is_cut_well(str(bad))
    assert not is_cut_well(str(broken))
    assert not is_cut_well(str(tmp_path / "нет.pptx"))


def test_operator_fixes_rename_or_drop_slides():
    from scripts.build_old_bank import DROP, apply_fixes

    entries = [
        {"raw": " - Ассорти Красивая любовь", "check": "не узнана", "source": "a.pptx", "slide_part": "ppt/slides/slide1.xml"},
        {"raw": " - FEDUK", "check": "не узнана", "source": "a.pptx", "slide_part": "ppt/slides/slide2.xml"},
        {"raw": "Звери - Танцуй", "check": "", "source": "b.pptx", "slide_part": "ppt/slides/slide1.xml"},
    ]
    fixes = {"a.pptx|ppt/slides/slide1.xml": "Ассорти - Красивая любовь", "a.pptx|ppt/slides/slide2.xml": DROP}

    result = apply_fixes(entries, fixes)

    assert [(e["raw"], e["check"]) for e in result] == [("Ассорти - Красивая любовь", ""), ("Звери - Танцуй", "")]


def test_fix_applies_to_every_slide_with_the_same_wrong_name():
    # Одна и та же опечатка бывает на слайдах разных игр — решение оператора относится ко всем
    from scripts.build_old_bank import DROP, apply_fixes

    entries = [
        {"raw": " - FEDUK", "check": "не узнана", "source": "a.pptx", "slide_part": "ppt/slides/slide2.xml"},
        {"raw": " - FEDUK", "check": "не узнана", "source": "b.pptx", "slide_part": "ppt/slides/slide9.xml"},
        {"raw": " - Ассорти Красивая любовь", "check": "не узнана", "source": "c.pptx", "slide_part": "ppt/slides/slide4.xml"},
    ]
    fixes = {"raw| - FEDUK": DROP, "raw| - Ассорти Красивая любовь": "Ассорти - Красивая любовь"}

    assert [e["raw"] for e in apply_fixes(entries, fixes)] == ["Ассорти - Красивая любовь"]
