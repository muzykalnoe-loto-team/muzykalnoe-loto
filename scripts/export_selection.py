import json


def export_selection(selected: list[str], output_path: str) -> None:
    data = {
        "barrels": [
            {"number": i + 1, "song": song}
            for i, song in enumerate(selected)
        ]
    }
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
