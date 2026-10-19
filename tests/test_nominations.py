import pytest

from scripts.nominations import choose_nominations, load_history, record_game


def test_auto_alternates_each_round_separately():
    history = [{"event": "А", "date": "2026-10-01", "themed": False,
                "nominations": {"1": "sing", "2": "cherry", "3": "dancefloor"}}]
    assert choose_nominations(False, {}, history) == {1: "cry", 2: "goddess", 3: "nobrakes"}


def test_first_game_gets_first_variant():
    assert choose_nominations(False, {}, []) == {1: "sing", 2: "goddess", 3: "dancefloor"}


def test_operator_choice_beats_queue():
    assert choose_nominations(False, {1: "sing", 3: "auto"}, [])[1] == "sing"


def test_themed_game_takes_theme_even_if_round1_chosen():
    assert choose_nominations(True, {1: "cry"}, [])[1] == "theme"


def test_round1_queue_ignores_themed_games():
    history = [{"event": "Обычная", "date": "1", "themed": False,
                "nominations": {"1": "cry", "2": "goddess", "3": "dancefloor"}},
               {"event": "Гэтсби", "date": "2", "themed": True,
                "nominations": {"1": "theme", "2": "cherry", "3": "nobrakes"}}]
    chosen = choose_nominations(False, {}, history)
    assert chosen[1] == "sing" and chosen[2] == "goddess" and chosen[3] == "dancefloor"


def test_unknown_choice_names_the_round():
    with pytest.raises(ValueError, match="2"):
        choose_nominations(False, {2: "sing"}, [])  # «Щас спою!» — номинация 1-го раунда, не 2-го


def test_missing_or_broken_history_is_empty(tmp_path):
    bad = tmp_path / "h.json"
    bad.write_text("{не json", encoding="utf-8")
    assert load_history(str(tmp_path / "нет.json")) == [] and load_history(str(bad)) == []


def test_record_game_appends(tmp_path):
    path = str(tmp_path / "h.json")
    record_game("Осенний квиз", "2026-10-20", False, {1: "sing", 2: "goddess", 3: "dancefloor"}, path)
    record_game("Гэтсби-бал", "2026-11-01", True, {1: "theme", 2: "cherry", 3: "nobrakes"}, path)
    history = load_history(path)
    assert len(history) == 2
    assert history[-1] == {"event": "Гэтсби-бал", "date": "2026-11-01", "themed": True,
                           "nominations": {"1": "theme", "2": "cherry", "3": "nobrakes"}}
