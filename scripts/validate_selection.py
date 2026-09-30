def validate_selection(selected: list[str], bank: list[dict]) -> dict:
    errors = []

    if len(selected) != 90:
        errors.append(f"Ожидается ровно 90 песен, передано {len(selected)}")

    seen = set()
    duplicates = set()
    for song in selected:
        if song in seen:
            duplicates.add(song)
        seen.add(song)
    if duplicates:
        errors.append("Найдены дубликаты: " + ", ".join(sorted(duplicates)))

    bank_raw = {item["raw"] for item in bank}
    unknown = [song for song in selected if song not in bank_raw]
    if unknown:
        errors.append("Песни не найдены в базе: " + ", ".join(unknown))

    return {"ok": len(errors) == 0, "errors": errors}
