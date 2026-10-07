def find_slide_index_by_barrel_number(prs, number: int) -> int:
    target = str(number)
    for index, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.name == "Текст 9" and shape.has_text_frame:
                if shape.text_frame.text.strip() == target:
                    return index
    raise LookupError(f"Слот для бочонка {number} не найден в шаблоне")
