import json


def export_selection(selected: list[str], output_path: str, kinds: list[str] | None = None) -> None:
    barrels = [{"number": i + 1, "song": song} for i, song in enumerate(selected)]
    if kinds is not None:
        if len(kinds) != len(selected):
            raise ValueError("Видов слайдов должно быть столько же, сколько песен")
        for barrel, kind in zip(barrels, kinds):
            barrel["kind"] = kind
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump({"barrels": barrels}, f, ensure_ascii=False, indent=2)
