from scripts.song_names import norm


def validate_selection(selected: list[str], bank: list[dict]) -> dict:
    errors = []

    if len(selected) != 90:
        errors.append(f"Ожидается ровно 90 песен, передано {len(selected)}")

    seen = set()
    duplicates = set()
    for song in selected:
        # песни сравниваем по norm: «ё/е», регистр и знаки — одна и та же песня
        if norm(song) in seen:
            duplicates.add(song)
        seen.add(norm(song))
    if duplicates:
        errors.append("Найдены дубликаты: " + ", ".join(sorted(duplicates)))

    bank_keys = {norm(item["raw"]) for item in bank}
    unknown = [song for song in selected if norm(song) not in bank_keys]
    if unknown:
        errors.append("Песни не найдены в базе: " + ", ".join(unknown))

    return {"ok": len(errors) == 0, "errors": errors}
