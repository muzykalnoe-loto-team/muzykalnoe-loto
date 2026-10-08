import json
from scripts.export_selection import export_selection


def test_export_selection_writes_numbered_json(tmp_path):
    output = tmp_path / "selected_songs.json"
    selected = ["Артист1 - Песня1", "Артист2 - Песня2", "Артист3 - Песня3"]

    export_selection(selected, str(output))

    with open(output, encoding="utf-8") as f:
        data = json.load(f)

    assert data == {
        "barrels": [
            {"number": 1, "song": "Артист1 - Песня1"},
            {"number": 2, "song": "Артист2 - Песня2"},
            {"number": 3, "song": "Артист3 - Песня3"},
        ]
    }


def test_export_selection_writes_kinds_when_given(tmp_path):
    output = tmp_path / "selected_songs.json"

    export_selection(["А - 1", "Б - 2", "В - 3"], str(output), kinds=["video", "video", "old"])

    with open(output, encoding="utf-8") as f:
        barrels = json.load(f)["barrels"]
    assert [b["kind"] for b in barrels] == ["video", "video", "old"]
