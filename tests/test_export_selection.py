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
