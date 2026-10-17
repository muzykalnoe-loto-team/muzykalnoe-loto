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


def _bank_case(tmp_path, monkeypatch):
    """Игра из двух песенных слайдов и запись каталога для первого; сжатие подменено счётчиком."""
    import os

    from scripts import build_old_bank
    from tests.test_old_slides import _make_old_game

    game = tmp_path / "игра.pptx"
    _make_old_game(str(game), [(["1", "ТАНЦУЙ", "Звери"], (255, 0, 0)), (["2", "ИСКАЛА", "Земфира"], None)])
    calls = []

    def fake_compact(src, out, photo_part):
        calls.append(photo_part)
        with open(src, "rb") as f, open(out, "wb") as g:
            g.write(f.read())
        return "ppt/media/сжатое.jpg"

    monkeypatch.setattr(build_old_bank, "compact_slide", fake_compact)
    entry = {"raw": "Звери - Танцуй", "file": "Звери - Танцуй.pptx", "source": str(game),
             "slide_part": "ppt/slides/slide1.xml", "photo_part": "ppt/media/image2.png"}
    bank = tmp_path / "банк"
    os.makedirs(bank)
    return build_old_bank, entry, str(bank), calls


def test_new_bank_file_is_cut_and_compacted(tmp_path, monkeypatch):
    build_old_bank, entry, bank, calls = _bank_case(tmp_path, monkeypatch)

    build_old_bank.prepare_bank_file(entry, bank, previous={})

    assert build_old_bank.is_cut_well(f"{bank}/Звери - Танцуй.pptx")
    assert calls == ["ppt/media/image2.png"]
    assert entry["compact"] and entry["photo_part"] == "ppt/media/сжатое.jpg"


def test_compacted_file_is_not_compacted_again(tmp_path, monkeypatch):
    build_old_bank, entry, bank, calls = _bank_case(tmp_path, monkeypatch)
    build_old_bank.prepare_bank_file(dict(entry), bank, previous={})
    previous = {entry["file"]: dict(entry, compact=True, photo_part="ppt/media/сжатое.jpg")}
    calls.clear()

    again = dict(entry)
    build_old_bank.prepare_bank_file(again, bank, previous=previous)

    assert calls == []  # повторная пересборка не делает ту же работу дважды
    assert again["compact"] and again["photo_part"] == "ppt/media/сжатое.jpg"


def test_already_cut_but_not_compacted_file_gets_compacted(tmp_path, monkeypatch):
    # банк, нарезанный до появления сжатия: файл на месте, в прошлом каталоге отметки «сжат» нет
    build_old_bank, entry, bank, calls = _bank_case(tmp_path, monkeypatch)
    from scripts.slide_split import extract_slide
    extract_slide(entry["source"], entry["slide_part"], f"{bank}/{entry['file']}")

    build_old_bank.prepare_bank_file(entry, bank, previous={entry["file"]: dict(entry)})

    assert calls == ["ppt/media/image2.png"] and entry["compact"]


def test_compaction_error_keeps_uncompressed_file(tmp_path, monkeypatch):
    build_old_bank, entry, bank, calls = _bank_case(tmp_path, monkeypatch)

    def broken(src, out, photo_part):
        raise ValueError("странный mp3")

    monkeypatch.setattr(build_old_bank, "compact_slide", broken)
    build_old_bank.prepare_bank_file(entry, bank, previous={})

    assert build_old_bank.is_cut_well(f"{bank}/Звери - Танцуй.pptx")  # песня в банке осталась, просто не сжата
    assert not entry["compact"] and "странный mp3" in entry["compact_error"]
    assert entry["photo_part"] == "ppt/media/image2.png"


def test_file_is_recut_when_best_slide_changed(tmp_path, monkeypatch):
    # после исправления .jfif/.bin «лучшим» у песни стал другой слайд, а файл остался от прежнего
    build_old_bank, entry, bank, calls = _bank_case(tmp_path, monkeypatch)
    from scripts.slide_split import extract_slide
    extract_slide(entry["source"], "ppt/slides/slide2.xml", f"{bank}/{entry['file']}")  # нарезан не тот слайд
    previous = {entry["file"]: dict(entry, compact=True, photo_part="ppt/media/сжатое.jpg")}

    build_old_bank.prepare_bank_file(entry, bank, previous=previous)

    import zipfile
    with zipfile.ZipFile(f"{bank}/{entry['file']}") as z:
        assert "ppt/slides/slide1.xml" in z.namelist()
    assert calls == ["ppt/media/image2.png"]  # нарезан заново и сжат
    assert entry["cut_from"] == f"{entry['source']}|ppt/slides/slide1.xml"


def test_file_is_recut_when_recorded_source_differs(tmp_path, monkeypatch):
    build_old_bank, entry, bank, calls = _bank_case(tmp_path, monkeypatch)
    build_old_bank.prepare_bank_file(dict(entry), bank, previous={})
    previous = {entry["file"]: dict(entry, compact=True, photo_part="ppt/media/сжатое.jpg",
                                    cut_from="другая игра.pptx|ppt/slides/slide1.xml")}
    calls.clear()

    build_old_bank.prepare_bank_file(entry, bank, previous=previous)

    assert calls == ["ppt/media/image2.png"]
