# Музыкальное лото — MVP (стадии 1–3) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Автоматизировать три стадии пайплайна музыкального лото — список песен, бланки, презентация — так, чтобы из утверждённого списка 90 песен получались готовый файл бланков и готовая презентация, без риска сломать вшитые в шаблоны формулы, макросы и гиперссылки.

**Architecture:** Набор Python-скриптов, запускаемых вручную (мной как ассистентом) в рабочей сессии на машине оператора. Часть скриптов — чистая логика (парсинг Excel, поиск слота в шаблоне) и покрывается обычными unit-тестами. Часть — драйвит реальный Microsoft Excel/PowerPoint через COM-автоматизацию (pywin32), воспроизводя тот же ручной процесс, которым оператор уже пользуется, и проверяется интеграционными тестами + ручной проверкой открытого файла. Стадия 4 (визуал) — вне scope, отдельный план позже.

**Tech Stack:** Python 3.13, openpyxl (чтение/запись Excel), python-pptx (чтение структуры pptx без COM), pywin32 / win32com.client (управление Excel и PowerPoint через COM), pytest.

**Spec:** `Спецификация.md` (корень проекта) — общая архитектура и 4 стадии пайплайна.

## Прогресс (на 2026-10-07)

- ✅ Сделано: Task 1, 2, 3, 4, 5, 6, 7, 8, 9, 10.
- ⬜ Осталось: Task 11 (оркестратор), 12 (выгрузка в панель).
- Дополнительно, вне плана: ✅ скилл подбора 90 песен (`.claude/skills/podobrat-pesni/`); ✅ «Панель мероприятий»: песни по желанию, список мест, изменение и удаление мероприятий; ✅ шаблон «Рабочий 3 вар» (все 90 слотов на макете караоке с кнопкой).
- 2026-10-07: исходные файлы перенесены в папку `Исходники/`, пути заданы в `scripts/paths.py`; рабочий шаблон презентации — `Рабочий 3 вар.pptx`. Пути в задачах ниже записаны на момент составления плана.

## Global Constraints

- Ровно 90 песен/бочонков на мероприятие — меньше или больше не допускается.
- Реальные входные файлы (уже лежат в корне проекта и изучены):
  - `Банк песен.xlsx` — лист `2301`, колонка A = номер (1..1071), колонка B = `"Исполнитель - Название"` одной строкой.
  - `Шаблон бланков.xlsm` — лист `Список`: B2:B91 = 90 песен (формат `"Исполнитель - Название"`), E1 = кол-во индивидуальных бланков, E2 = кол-во командных бланков. Содержит VBA-макрос `GenerateBlanks.GenerateAllBlanks`, который **обязателен к запуску** — без него листы `Индивидуальные бланки`/`Командные бланки` не обновятся.
  - `Рабочий 2 вар.pptx` — 118 слайдов. Слайды 16–105 (1-индексация) — это 90 слотов-заглушек под бочонки: у каждого есть фигура с именем `Текст 9`, текст которой равен номеру бочонка (`"1"`…`"90"`). Слоты уже содержат рабочую управляющую кнопку с вшитой гиперссылкой навигации — **эту фигуру нельзя удалять или заменять слайд целиком**.
  - `Банк всех слайдов23.01.2026/` — 1071 файл `.pptx`, имя файла = `"Исполнитель - Название.pptx"`, ровно соответствует строкам `Банк песен.xlsx`. Каждый файл — один слайд с готовым видео (layout `Игровой слайд караоке`).
- Способ сборки презентации: **копировать все фигуры со слайда из архива и вставлять их поверх существующих фигур слота в шаблоне** (Shapes.Range().Copy → Slides(n).Shapes.Paste()) — именно так оператор делает это вручную, чтобы не сломать гиперссылки. Полная замена слайда запрещена.
- На машине подтверждено: Microsoft Excel 16.0 и PowerPoint 16.0 доступны через COM (`win32com.client`). Для COM-запуска VBA-макроса нужно один раз включить в Excel: Файл → Параметры → Центр управления безопасностью → Параметры центра управления безопасностью → Параметры макросов → «Доверять доступ к объектной модели проектов VBA».
- COM-автоматизация Office требует интерактивной Windows-сессии (не headless-сервера) — это ограничение принято осознанно, т.к. вся система и так работает на локальной машине оператора.
- Исходные файлы-шаблоны, база песен и архив слайдов не коммитятся в git (большие/защищённые авторским правом бинарники) — исключить через `.gitignore`. Результаты по конкретным мероприятиям (`events/`) тоже не коммитятся.

---

## Task 1: Инициализация проекта ✅ сделано

**Files:**
- Create: `requirements.txt`
- Create: `scripts/__init__.py`
- Create: `tests/__init__.py`
- Modify: `.gitignore` (создать, если ещё нет)

**Interfaces:**
- Produces: структура каталогов `scripts/` и `tests/`, зависимости для всех последующих задач.

- [ ] **Step 1: Создать `requirements.txt`**

```
openpyxl>=3.1
python-pptx>=1.0
pywin32>=306
pytest>=8.0
```

- [ ] **Step 2: Создать пустые пакеты**

Создать файлы `scripts/__init__.py` и `tests/__init__.py` (пустые).

- [ ] **Step 3: Создать `.gitignore`**

```
# Исходные материалы мероприятий — большие/защищённые файлы, не хранятся в git
Банк песен.xlsx
Шаблон бланков.xlsm
Рабочий 2 вар.pptx
Банк всех слайдов*/
events/

# Служебное
__pycache__/
*.pyc
.pytest_cache/
```

- [ ] **Step 4: Установить зависимости**

Run: `pip install -r requirements.txt`
Expected: все пакеты устанавливаются без ошибок (openpyxl и python-pptx уже стояли; pywin32 и pytest — новые).

- [ ] **Step 5: Commit**

```bash
git add requirements.txt scripts/__init__.py tests/__init__.py .gitignore
git commit -m "chore: инициализировать структуру проекта для MVP-скриптов"
```

---

## Task 2: Загрузка базы песен ✅ сделано

**Files:**
- Create: `scripts/song_bank.py`
- Test: `tests/test_song_bank.py`

**Interfaces:**
- Produces: `load_song_bank(path: str) -> list[dict]`, где каждый элемент — `{"n": int, "raw": str, "artist": str, "title": str}`. `raw` — исходная строка `"Исполнитель - Название"` как она записана в базе (используется как ключ поиска в архиве слайдов).

- [ ] **Step 1: Написать падающий тест**

```python
# tests/test_song_bank.py
import openpyxl
from scripts.song_bank import load_song_bank


def make_bank_fixture(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "2301"
    rows = [
        (1, "Юта - Белая вьюга"),
        (2, "#2Маши - Босая"),
        (3, "5sta Family - Я буду"),
    ]
    for r in rows:
        ws.append(r)
    wb.save(path)


def test_load_song_bank_parses_artist_and_title(tmp_path):
    fixture = tmp_path / "bank.xlsx"
    make_bank_fixture(fixture)

    result = load_song_bank(str(fixture))

    assert result == [
        {"n": 1, "raw": "Юта - Белая вьюга", "artist": "Юта", "title": "Белая вьюга"},
        {"n": 2, "raw": "#2Маши - Босая", "artist": "#2Маши", "title": "Босая"},
        {"n": 3, "raw": "5sta Family - Я буду", "artist": "5sta Family", "title": "Я буду"},
    ]


def test_load_song_bank_skips_empty_rows(tmp_path):
    fixture = tmp_path / "bank_with_gap.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "2301"
    ws.append((1, "Юта - Белая вьюга"))
    ws.append((2, None))
    wb.save(fixture)

    result = load_song_bank(str(fixture))

    assert result == [
        {"n": 1, "raw": "Юта - Белая вьюга", "artist": "Юта", "title": "Белая вьюга"},
    ]
```

- [ ] **Step 2: Убедиться, что тест падает**

Run: `pytest tests/test_song_bank.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'scripts.song_bank'`

- [ ] **Step 3: Реализовать минимальный код**

```python
# scripts/song_bank.py
import openpyxl


def load_song_bank(path: str) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]

    songs = []
    for row in ws.iter_rows(min_row=1, values_only=True):
        n, raw = row[0], row[1]
        if raw is None or str(raw).strip() == "":
            continue
        raw = str(raw).strip()
        artist, _, title = raw.partition(" - ")
        songs.append({
            "n": int(n),
            "raw": raw,
            "artist": artist.strip(),
            "title": title.strip(),
        })
    return songs
```

- [ ] **Step 4: Убедиться, что тест проходит**

Run: `pytest tests/test_song_bank.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
git add scripts/song_bank.py tests/test_song_bank.py
git commit -m "feat: загрузка и парсинг базы песен из Банк песен.xlsx"
```

---

## Task 3: Валидация утверждённого списка из 90 песен ✅ сделано

**Files:**
- Create: `scripts/validate_selection.py`
- Test: `tests/test_validate_selection.py`

**Interfaces:**
- Consumes: список песен из Task 2 (`list[dict]` с полем `raw`).
- Produces: `validate_selection(selected: list[str], bank: list[dict]) -> dict`, возвращает `{"ok": bool, "errors": list[str]}`.

- [ ] **Step 1: Написать падающий тест**

```python
# tests/test_validate_selection.py
from scripts.validate_selection import validate_selection

BANK = [
    {"n": i, "raw": f"Артист{i} - Песня{i}", "artist": f"Артист{i}", "title": f"Песня{i}"}
    for i in range(1, 101)
]


def test_valid_selection_of_90_passes():
    selected = [f"Артист{i} - Песня{i}" for i in range(1, 91)]

    result = validate_selection(selected, BANK)

    assert result == {"ok": True, "errors": []}


def test_wrong_count_fails():
    selected = [f"Артист{i} - Песня{i}" for i in range(1, 90)]  # 89 штук

    result = validate_selection(selected, BANK)

    assert result["ok"] is False
    assert any("90" in e for e in result["errors"])


def test_duplicate_fails():
    selected = [f"Артист{i} - Песня{i}" for i in range(1, 90)] + ["Артист1 - Песня1"]

    result = validate_selection(selected, BANK)

    assert result["ok"] is False
    assert any("дубл" in e.lower() for e in result["errors"])


def test_unknown_song_fails():
    selected = [f"Артист{i} - Песня{i}" for i in range(1, 90)] + ["Неизвестный - Трек"]

    result = validate_selection(selected, BANK)

    assert result["ok"] is False
    assert any("Неизвестный - Трек" in e for e in result["errors"])
```

- [ ] **Step 2: Убедиться, что тест падает**

Run: `pytest tests/test_validate_selection.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'scripts.validate_selection'`

- [ ] **Step 3: Реализовать минимальный код**

```python
# scripts/validate_selection.py

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
```

- [ ] **Step 4: Убедиться, что тест проходит**

Run: `pytest tests/test_validate_selection.py -v`
Expected: PASS (4 passed)

- [ ] **Step 5: Commit**

```bash
git add scripts/validate_selection.py tests/test_validate_selection.py
git commit -m "feat: валидация утверждённого списка из 90 песен"
```

---

## Task 4: Экспорт утверждённого списка с номерами бочонков ✅ сделано

**Files:**
- Create: `scripts/export_selection.py`
- Test: `tests/test_export_selection.py`

**Interfaces:**
- Consumes: список из 90 строк в порядке бочонков (индекс 0 = бочонок 1).
- Produces: `export_selection(selected: list[str], output_path: str) -> None`, пишет JSON вида `{"barrels": [{"number": 1, "song": "..."}, ...]}`. Это файл читают Task 5 (бланки) и Task 8 (презентация).

- [ ] **Step 1: Написать падающий тест**

```python
# tests/test_export_selection.py
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
```

- [ ] **Step 2: Убедиться, что тест падает**

Run: `pytest tests/test_export_selection.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'scripts.export_selection'`

- [ ] **Step 3: Реализовать минимальный код**

```python
# scripts/export_selection.py
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
```

- [ ] **Step 4: Убедиться, что тест проходит**

Run: `pytest tests/test_export_selection.py -v`
Expected: PASS (1 passed)

- [ ] **Step 5: Commit**

```bash
git add scripts/export_selection.py tests/test_export_selection.py
git commit -m "feat: экспорт утверждённого списка песен с номерами бочонков"
```

---

## Task 5: Генерация бланков через Excel COM ✅ сделано

**Files:**
- Create: `scripts/generate_blanks.py`
- Test: `tests/test_generate_blanks.py` (интеграционный, требует реальный Excel и реальный `Шаблон бланков.xlsm`)

**Interfaces:**
- Consumes: `barrels` из Task 4 (список из 90 строк `"Исполнитель - Название"` в порядке бочонков), `n_individual: int`, `n_team: int`.
- Produces: `generate_blanks(template_path: str, output_path: str, selected_songs: list[str], n_individual: int, n_team: int) -> None`. Пишет готовый `.xlsm` с заполненными листами `Индивидуальные бланки` и `Командные бланки`.

- [ ] **Step 1: Написать интеграционный тест**

Этот тест реально открывает Excel и запускает существующий макрос — он медленный и требует Windows + Excel с включённым доверием к объектной модели VBA (см. Global Constraints). Помечаем его `@pytest.mark.integration`, чтобы не гонять при каждом обычном запуске.

```python
# tests/test_generate_blanks.py
import shutil
import pytest
import openpyxl
from scripts.generate_blanks import generate_blanks
from scripts.song_bank import load_song_bank

TEMPLATE = "Шаблон бланков.xlsm"
BANK = "Банк песен.xlsx"


@pytest.mark.integration
def test_generate_blanks_fills_individual_sheet(tmp_path):
    bank = load_song_bank(BANK)
    selected = [item["raw"] for item in bank[:90]]

    output = tmp_path / "blanks_output.xlsm"

    generate_blanks(TEMPLATE, str(output), selected, n_individual=3, n_team=2)

    assert output.exists()

    wb = openpyxl.load_workbook(output, data_only=True)
    ws = wb["Индивидуальные бланки"]
    # Первый бланк начинается в A1 с номером бланка в I1
    assert ws["A1"].value == "Музыкальное лото\nvk.com/dvizh_nv"
    assert ws["I1"].value == 1
    # Хотя бы одна ячейка сетки заполнена реальной песней, а не пустая
    assert ws["A3"].value not in (None, "")


@pytest.mark.integration
def test_generate_blanks_rejects_wrong_song_count(tmp_path):
    output = tmp_path / "blanks_output.xlsm"

    with pytest.raises(ValueError, match="90"):
        generate_blanks(TEMPLATE, str(output), ["Артист - Песня"], n_individual=1, n_team=1)
```

- [ ] **Step 2: Убедиться, что тест падает**

Run: `pytest tests/test_generate_blanks.py -v -m integration`
Expected: FAIL — `ModuleNotFoundError: No module named 'scripts.generate_blanks'`

- [ ] **Step 3: Реализовать минимальный код**

```python
# scripts/generate_blanks.py
import os
import win32com.client as win32

XL_MACRO_ENABLED_WORKBOOK = 52  # xlOpenXMLWorkbookMacroEnabled


def generate_blanks(
    template_path: str,
    output_path: str,
    selected_songs: list[str],
    n_individual: int,
    n_team: int,
) -> None:
    if len(selected_songs) != 90:
        raise ValueError(f"Ожидается ровно 90 песен, передано {len(selected_songs)}")

    template_path = os.path.abspath(template_path)
    output_path = os.path.abspath(output_path)

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(template_path)
        try:
            ws_list = wb.Sheets("Список")
            for i, song in enumerate(selected_songs):
                ws_list.Cells(i + 2, 2).Value = song  # B2:B91

            ws_list.Range("E1").Value = n_individual
            ws_list.Range("E2").Value = n_team

            excel.Run("GenerateBlanks.GenerateAllBlanks")

            wb.SaveAs(output_path, FileFormat=XL_MACRO_ENABLED_WORKBOOK)
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
```

- [ ] **Step 4: Убедиться, что тест проходит**

Run: `pytest tests/test_generate_blanks.py -v -m integration`
Expected: PASS (2 passed). Если падает на `excel.Run` с ошибкой доступа — проверить настройку «Доверять доступ к объектной модели проектов VBA» в Excel (см. Global Constraints).

- [x] **Step 5: Ручная проверка** (оператор проверил 2026-10-07: видео играет, кнопка переводит дальше)

Открыть `output.xlsm` в Excel руками, убедиться, что бланки визуально выглядят правильно (номера бочонков, песни на местах, нет обрезанного текста).

- [ ] **Step 6: Commit**

```bash
git add scripts/generate_blanks.py tests/test_generate_blanks.py pytest.ini
git commit -m "feat: автоматическая генерация бланков через Excel COM и штатный макрос"
```

(Примечание: `pytest.ini` создаётся в Task 9, если ещё не создан — см. ниже про маркер `integration`.)

---

## Task 6: Поиск слайда-слота по номеру бочонка ✅ сделано

**Files:**
- Create: `scripts/template_slots.py`
- Test: `tests/test_template_slots.py`

**Interfaces:**
- Consumes: путь к файлу презентации-шаблона.
- Produces: `find_slide_index_by_barrel_number(prs, number: int) -> int` — возвращает 0-based индекс слайда в `prs.slides` (python-pptx `Presentation` объект), у которого есть фигура с именем `"Текст 9"` и текстом, равным `str(number)`. Бросает `LookupError`, если не найдено.

- [x] **Step 1: Написать падающий тест**

Тест использует реальный `Рабочий 2 вар.pptx` как фикстуру для чтения (открывается только на чтение через python-pptx, без COM — быстро и безопасно).

```python
# tests/test_template_slots.py
import pytest
from pptx import Presentation
from scripts.template_slots import find_slide_index_by_barrel_number

TEMPLATE = "Рабочий 2 вар.pptx"


def test_finds_slot_for_barrel_1():
    prs = Presentation(TEMPLATE)

    index = find_slide_index_by_barrel_number(prs, 1)

    assert index == 15  # слайд №16 в 1-индексации PowerPoint


def test_finds_slot_for_barrel_90():
    prs = Presentation(TEMPLATE)

    index = find_slide_index_by_barrel_number(prs, 90)

    assert index == 104  # слайд №105 в 1-индексации PowerPoint


def test_raises_for_unknown_barrel():
    prs = Presentation(TEMPLATE)

    with pytest.raises(LookupError):
        find_slide_index_by_barrel_number(prs, 91)
```

- [x] **Step 2: Убедиться, что тест падает**

Run: `pytest tests/test_template_slots.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'scripts.template_slots'`

- [x] **Step 3: Реализовать минимальный код**

```python
# scripts/template_slots.py

def find_slide_index_by_barrel_number(prs, number: int) -> int:
    target = str(number)
    for index, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.name == "Текст 9" and shape.has_text_frame:
                if shape.text_frame.text.strip() == target:
                    return index
    raise LookupError(f"Слот для бочонка {number} не найден в шаблоне")
```

- [x] **Step 4: Убедиться, что тест проходит**

Run: `pytest tests/test_template_slots.py -v`
Expected: PASS (3 passed)

- [x] **Step 5: Commit**

```bash
git add scripts/template_slots.py tests/test_template_slots.py
git commit -m "feat: поиск слайда-слота в шаблоне по номеру бочонка"
```

---

## Task 7: Вставка слайда песни из архива через PowerPoint COM ✅ сделано

**Files:**
- Create: `scripts/ppt_paste.py`
- Test: `tests/test_ppt_paste.py` (интеграционный, требует реальный PowerPoint)

**Interfaces:**
- Consumes: живой COM-объект PowerPoint (`ppt_app`), живой COM-объект презентации-шаблона (`template_prs`), путь к файлу архива, 1-based номер слайда-слота (из Task 6, `+1` к 0-based индексу).
- Produces: `insert_song_slide(ppt_app, template_prs, archive_file_path: str, slide_index_1based: int) -> None`. После вызова у целевого слайда фигур становится больше (добавлены фигуры из архивного слайда), существующие фигуры шаблона не тронуты.

- [x] **Step 1: Написать интеграционный тест**

```python
# tests/test_ppt_paste.py
import shutil
import pytest
import win32com.client as win32
from scripts.ppt_paste import insert_song_slide

TEMPLATE = "Рабочий 2 вар.pptx"
ARCHIVE_SONG = "Банк всех слайдов23.01.2026/5sta Family - Я буду.pptx"


@pytest.mark.integration
def test_insert_song_slide_adds_shapes_without_removing_existing(tmp_path):
    working_copy = tmp_path / "template_copy.pptx"
    shutil.copy(TEMPLATE, working_copy)

    ppt_app = win32.gencache.EnsureDispatch("PowerPoint.Application")
    ppt_app.Visible = True
    try:
        template_prs = ppt_app.Presentations.Open(str(working_copy))
        try:
            target_slide = template_prs.Slides(16)  # слот бочонка №1
            shapes_before = target_slide.Shapes.Count

            insert_song_slide(ppt_app, template_prs, ARCHIVE_SONG, 16)

            shapes_after = target_slide.Shapes.Count
            assert shapes_after > shapes_before

            template_prs.Save()
        finally:
            template_prs.Close()
    finally:
        ppt_app.Quit()
```

- [x] **Step 2: Убедиться, что тест падает**

Run: `pytest tests/test_ppt_paste.py -v -m integration`
Expected: FAIL — `ModuleNotFoundError: No module named 'scripts.ppt_paste'`

- [x] **Step 3: Реализовать минимальный код**

```python
# scripts/ppt_paste.py

def insert_song_slide(ppt_app, template_prs, archive_file_path: str, slide_index_1based: int) -> None:
    archive_prs = ppt_app.Presentations.Open(archive_file_path, WithWindow=True)
    try:
        archive_slide = archive_prs.Slides(1)
        shape_range = archive_slide.Shapes.Range()
        shape_range.Copy()

        target_slide = template_prs.Slides(slide_index_1based)
        target_slide.Shapes.Paste()
    finally:
        archive_prs.Close()
```

- [x] **Step 4: Убедиться, что тест проходит**

Run: `pytest tests/test_ppt_paste.py -v -m integration`
Expected: PASS (1 passed)

- [x] **Step 5: Ручная проверка** (оператор проверил 2026-10-07: видео играет, кнопка переводит дальше)

Открыть результирующий файл из `tmp_path` (тест печатает путь при `--capture=no`, либо временно сохранить копию в постоянное место), вручную посмотреть слайд слота №1: видео должно проигрываться, кнопка навигации — вести на следующий слайд, номер бочонка отображаться один раз визуально (даже если в XML фигур две — они наложены друг на друга).

> **Отличия от плана при выполнении:** (1) PowerPoint вставляет фигуры со сдвигом на 12 pt — после вставки фигуры возвращаются на позиции из архива; (2) видео на весь слайд и неработающая кнопка из архива ложились поверх рабочей кнопки навигации шаблона — после вставки фигуры слота поднимаются наверх. Тест проверяет оба пункта. Подключение через `Dispatch` (как в Task 5), тестовая песня — «Натали - Шахерезада» (бочонок 1 списка «8 марта»).

- [x] **Step 6: Commit**

```bash
git add scripts/ppt_paste.py tests/test_ppt_paste.py
git commit -m "feat: вставка слайда песни из архива поверх слота в шаблоне через PowerPoint COM"
```

---

## Task 8: Сборка полной презентации мероприятия ✅ сделано (ждёт ручной проверки оператором)

**Files:**
- Create: `scripts/build_presentation.py`
- Test: `tests/test_build_presentation.py` (интеграционный)

**Interfaces:**
- Consumes: `barrels` — список `[{"number": int, "song": str}, ...]` из Task 4 (файл `selected_songs.json`), `template_path`, `archive_folder`.
- Produces: `assemble_presentation(template_path: str, archive_folder: str, barrels: list[dict], output_path: str) -> None`. Итоговый `.pptx`, полностью соответствующий утверждённому списку песен.

- [x] **Step 1: Написать интеграционный тест**

Полный прогон на всех 90 песнях — долгий (реально открывает 90 файлов в PowerPoint). Для теста берём укороченный набор из 3 барабанов, чтобы проверить логику склейки без ожидания полного цикла.

```python
# tests/test_build_presentation.py
import shutil
import pytest
from pptx import Presentation
from scripts.build_presentation import assemble_presentation
from scripts.song_bank import load_song_bank

TEMPLATE = "Рабочий 2 вар.pptx"
ARCHIVE = "Банк всех слайдов23.01.2026"
BANK = "Банк песен.xlsx"


@pytest.mark.integration
def test_assemble_presentation_fills_requested_slots(tmp_path):
    bank = load_song_bank(BANK)
    barrels = [
        {"number": 1, "song": bank[0]["raw"]},
        {"number": 2, "song": bank[1]["raw"]},
        {"number": 90, "song": bank[2]["raw"]},
    ]

    output = tmp_path / "presentation_output.pptx"

    assemble_presentation(TEMPLATE, ARCHIVE, barrels, str(output))

    assert output.exists()

    prs = Presentation(str(output))
    # Общее число слайдов не изменилось — мы не удаляли и не добавляли слайды
    assert len(prs.slides) == 118


@pytest.mark.integration
def test_assemble_presentation_fails_fast_on_missing_archive_file(tmp_path):
    barrels = [{"number": 1, "song": "Несуществующий Артист - Несуществующая Песня"}]
    output = tmp_path / "presentation_output.pptx"

    with pytest.raises(FileNotFoundError, match="Несуществующий Артист"):
        assemble_presentation(TEMPLATE, ARCHIVE, barrels, str(output))
```

- [x] **Step 2: Убедиться, что тест падает**

Run: `pytest tests/test_build_presentation.py -v -m integration`
Expected: FAIL — `ModuleNotFoundError: No module named 'scripts.build_presentation'`

- [x] **Step 3: Реализовать минимальный код**

```python
# scripts/build_presentation.py
import os
import shutil
from pptx import Presentation
import win32com.client as win32

from scripts.template_slots import find_slide_index_by_barrel_number
from scripts.ppt_paste import insert_song_slide


def assemble_presentation(
    template_path: str,
    archive_folder: str,
    barrels: list[dict],
    output_path: str,
) -> None:
    # 1. Быстрая проверка наличия всех файлов архива до запуска COM
    resolved = []
    for barrel in barrels:
        archive_file = os.path.join(archive_folder, f"{barrel['song']}.pptx")
        if not os.path.exists(archive_file):
            raise FileNotFoundError(
                f"Не найден файл в архиве для бочонка {barrel['number']}: {barrel['song']}"
            )
        resolved.append((barrel["number"], archive_file))

    # 2. Определяем индексы слотов без COM (быстрое чтение через python-pptx)
    prs_readonly = Presentation(template_path)
    slot_indices = {
        number: find_slide_index_by_barrel_number(prs_readonly, number) + 1  # -> 1-based для COM
        for number, _ in resolved
    }

    # 3. Рабочая копия, чтобы не трогать оригинал шаблона
    working_copy = os.path.abspath(output_path)
    shutil.copy(template_path, working_copy)

    ppt_app = win32.gencache.EnsureDispatch("PowerPoint.Application")
    ppt_app.Visible = True
    try:
        template_prs = ppt_app.Presentations.Open(working_copy)
        try:
            for number, archive_file in resolved:
                insert_song_slide(
                    ppt_app,
                    template_prs,
                    os.path.abspath(archive_file),
                    slot_indices[number],
                )
            template_prs.Save()
        finally:
            template_prs.Close()
    finally:
        ppt_app.Quit()
```

- [x] **Step 4: Убедиться, что тест проходит**

Run: `pytest tests/test_build_presentation.py -v -m integration`
Expected: PASS (2 passed)

- [ ] **Step 5: Полный прогон на 90 песнях (ручной шаг, не часть автотестов)**

Собрать `barrels` из реального утверждённого списка (Task 4), запустить `assemble_presentation` на всех 90, открыть результат в PowerPoint, пролистать несколько случайных слайдов, убедиться что видео/навигация работают.

> **Отличия от плана при выполнении:** тест берёт бочонки 1, 2, 90 из списка «8 марта» и проверяет, что на их слотах появилось видео песни, а соседний слот не тронут; подключение через `Dispatch`. Полный прогон на 90 песнях «8 марта» — 43 с, файл ~1 ГБ, автоматическая проверка всех 90 слайдов: 90 из 90 в порядке. По ходу исправлены ошибки в `ppt_paste.py` (Task 7):
> 1. пустые текстовые поля архива не вставляются — сдвиг вставки теперь считается по видео, а не по порядку фигур;
> 2. на слайдах архива остался старый номер («1», «2»), он накладывался на номер бочонка («4²1») — стирается;
> 3. каждый 3-й слот шаблона (бочонки 3, 6 … 90) — без кнопки навигации, с заглушкой «ТЕКСТ НА 2 СТРОЧКИ»; наверх поднимаются только номер и кнопки с переходом, заглушки остаются под видео;
> 4. буфер обмена иногда не успевал обновиться и вставлялась предыдущая песня картинкой — вставка сверяется с видео песни и повторяется до 3 раз (тест `tests/test_ppt_paste_retry.py`).

- [x] **Step 6: Commit**

```bash
git add scripts/build_presentation.py tests/test_build_presentation.py
git commit -m "feat: сборка полной презентации мероприятия из архива по списку бочонков"
```

---

## Task 9: pytest-конфигурация для маркера integration ✅ сделано

**Files:**
- Create: `pytest.ini`

**Interfaces:**
- Produces: регистрация маркера `integration`, чтобы `pytest -m "not integration"` быстро гонял только чистую логику (Task 2–4, 6), а `pytest -m integration` — реальные прогоны через Excel/PowerPoint (Task 5, 7, 8).

- [ ] **Step 1: Создать `pytest.ini`**

```ini
[pytest]
markers =
    integration: медленные тесты, реально открывающие Excel/PowerPoint через COM
```

- [ ] **Step 2: Проверить разделение тестов**

Run: `pytest -m "not integration" -v`
Expected: прогоняются только тесты из Task 2, 3, 4, 6 (чистая логика), все PASS, быстро (без COM).

Run: `pytest -m integration -v`
Expected: прогоняются тесты из Task 5, 7, 8 (COM), дольше, но тоже PASS при исправно настроенном Excel/PowerPoint.

- [ ] **Step 3: Commit**

```bash
git add pytest.ini
git commit -m "chore: разделить быстрые unit-тесты и медленные COM-интеграционные тесты"
```

---

## Task 10: Артефакт «Панель мероприятий» — форма брифа и статусы ✅ сделано

**Files:**
- Create: `artifact/index.html`

**Interfaces:**
- Produces: веб-страница (Artifact), которая сохраняет бриф мероприятия (тема/дата/время/место/кол-во индивидуальных и командных бланков) в общую базу артефакта (capability `db`, коллекция `events`) со статусом `draft`. Дальнейшие стадии (список песен, бланки, презентация) обрабатываются мной вручную в рабочей сессии — страница их не запускает сама.

- [x] **Step 1: Собрать HTML-форму**

Создать `artifact/index.html` с полями: тема (text), дата (date), время (time), место (text), кол-во индивидуальных бланков (number, default 90), кол-во командных бланков (number, default 35), кнопка «Отправить бриф». При отправке — писать документ в коллекцию `events` через capability `db` (поля: `theme`, `date`, `time`, `place`, `n_individual`, `n_team`, `status: "draft"`, `created_at`) и показывать список уже созданных мероприятий с их статусами ниже формы.

(Перед реализацией — загрузить skill `artifact-capabilities`, чтобы корректно объявить capability `db` и вызовы `read_db`/`write_db` через `ArtifactData`, как требует правило работы с артефактами.)

- [x] **Step 2: Опубликовать и проверить в браузере**

Опубликовать артефакт (Artifact tool, `action: publish`), открыть его, вручную создать тестовый бриф, убедиться что он появляется в списке со статусом `draft`, обновить страницу — убедиться что запись сохранилась (проверка через `db`, не через localStorage).

> **Отличия от плана при выполнении:** опубликовано: https://claude.ai/artifact/QNGUxwrX1AHnyPGuPrzMWN (capability `db`, коллекция `events`). Кроме общего `status` у мероприятия есть `stages` — статус каждой стадии (`songs`, `blanks`, `presentation`: `draft` / `in_progress` / `review` / `approved` / `ready`), как в спецификации. В базу добавлено мероприятие «8 марта» (id `8-marta-2027-03-08`). Шаг 2 выполнен оператором (пробный бриф «Готический бал» сохранился в базе). Доработка по просьбе оператора: поле «Песни по желанию» — поиск по банку песен (файл `artifact/songs.json`, выгружается из `Банк песен.xlsx`, в git не хранится), выбранные песни сохраняются в `wanted_songs`; поле «Место» — выбор из мест прошлых игр или свой вариант.

- [x] **Step 3: Commit**

```bash
git add artifact/index.html
git commit -m "feat: артефакт «Панель мероприятий» — форма брифа со статусами"
```

---

## Task 11: Оркестратор — один прогон от списка до готовых файлов ⬜ не сделано

**Files:**
- Create: `scripts/run_pipeline.py`
- Test: `tests/test_run_pipeline.py`

**Interfaces:**
- Consumes: путь к `selected_songs.json` (формат Task 4), пути к `Банк песен.xlsx` (не используется напрямую, но остаётся источником для сверки), `Шаблон бланков.xlsm`, `Рабочий 2 вар.pptx`, папке архива, `n_individual`, `n_team`, папке для вывода.
- Produces: `run_pipeline(selected_songs_path: str, blanks_template: str, presentation_template: str, archive_folder: str, n_individual: int, n_team: int, output_dir: str) -> dict`. Возвращает `{"blanks_path": str, "presentation_path": str}`. Внутри переводит формат `barrels` (Task 4) в вид, ожидаемый `generate_blanks` (Task 5, список строк по порядку) и `assemble_presentation` (Task 8, список словарей) — это единственное место, где расхождение форматов между Task 5 и Task 8 разрешается.

- [ ] **Step 1: Написать падающий тест**

```python
# tests/test_run_pipeline.py
import json
import pytest
from scripts.run_pipeline import run_pipeline
from scripts.song_bank import load_song_bank

BANK = "Банк песен.xlsx"
BLANKS_TEMPLATE = "Шаблон бланков.xlsm"
PRESENTATION_TEMPLATE = "Рабочий 2 вар.pptx"
ARCHIVE = "Банк всех слайдов23.01.2026"


@pytest.mark.integration
def test_run_pipeline_produces_both_files(tmp_path):
    bank = load_song_bank(BANK)
    selected_path = tmp_path / "selected_songs.json"
    data = {
        "barrels": [
            {"number": i + 1, "song": bank[i]["raw"]}
            for i in range(90)
        ]
    }
    with open(selected_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)

    result = run_pipeline(
        str(selected_path),
        BLANKS_TEMPLATE,
        PRESENTATION_TEMPLATE,
        ARCHIVE,
        n_individual=5,
        n_team=2,
        output_dir=str(tmp_path),
    )

    assert (tmp_path / "blanks.xlsm").exists()
    assert (tmp_path / "presentation.pptx").exists()
    assert result == {
        "blanks_path": str(tmp_path / "blanks.xlsm"),
        "presentation_path": str(tmp_path / "presentation.pptx"),
    }
```

- [ ] **Step 2: Убедиться, что тест падает**

Run: `pytest tests/test_run_pipeline.py -v -m integration`
Expected: FAIL — `ModuleNotFoundError: No module named 'scripts.run_pipeline'`

- [ ] **Step 3: Реализовать минимальный код**

```python
# scripts/run_pipeline.py
import json
import os

from scripts.generate_blanks import generate_blanks
from scripts.build_presentation import assemble_presentation


def run_pipeline(
    selected_songs_path: str,
    blanks_template: str,
    presentation_template: str,
    archive_folder: str,
    n_individual: int,
    n_team: int,
    output_dir: str,
) -> dict:
    with open(selected_songs_path, encoding="utf-8") as f:
        data = json.load(f)

    barrels = sorted(data["barrels"], key=lambda b: b["number"])
    songs_in_order = [b["song"] for b in barrels]

    blanks_path = os.path.join(output_dir, "blanks.xlsm")
    generate_blanks(blanks_template, blanks_path, songs_in_order, n_individual, n_team)

    presentation_path = os.path.join(output_dir, "presentation.pptx")
    assemble_presentation(presentation_template, archive_folder, barrels, presentation_path)

    return {"blanks_path": blanks_path, "presentation_path": presentation_path}
```

- [ ] **Step 4: Убедиться, что тест проходит**

Run: `pytest tests/test_run_pipeline.py -v -m integration`
Expected: PASS (1 passed). Это самый долгий тест в наборе (реально открывает Excel и 90 файлов PowerPoint) — не запускать его в быстром цикле разработки, только перед сдачей.

- [ ] **Step 5: Commit**

```bash
git add scripts/run_pipeline.py tests/test_run_pipeline.py
git commit -m "feat: оркестратор полного прогона от списка песен до готовых файлов"
```

---

## Task 12: Выгрузка готовых файлов в артефакт для скачивания ⬜ не сделано

**Files:**
- Modify: `artifact/index.html` (из Task 10)

**Interfaces:**
- Consumes: `blanks_path` и `presentation_path` из Task 11.
- Produces: после того как я (в рабочей сессии) прогоняю `run_pipeline` для конкретного мероприятия, я загружаю оба файла как assets артефакта (Artifact tool, `action: "publish"`, `asset: true`, `url: <ссылка на артефакт>`) и записываю их URL в документ события в `db` (поля `blanks_url`, `presentation_url`, `status: "ready"`). Страница показывает кнопки-ссылки на скачивание для каждого события со статусом `ready`.

- [ ] **Step 1: Расширить артефакт блоком статусов и скачивания**

В `artifact/index.html`, в списке мероприятий (из Task 10), для каждого события со статусом `ready` вывести две ссылки: «Скачать бланки» (`blanks_url`) и «Скачать презентацию» (`presentation_url`). Для статусов `draft`/`in_progress` — просто показывать текстовый статус без ссылок.

(Объявление capability `assets` в артефакте — по инструкциям skill `artifact-capabilities`, загруженной в Task 10.)

- [ ] **Step 2: Проверить вручную**

Создать тестовое событие через форму (Task 10), вручную (через `ArtifactData`) проставить ему `status: "ready"` и тестовые `blanks_url`/`presentation_url` (можно временную ссылку), обновить страницу артефакта — убедиться что кнопки скачивания появились и статус `draft` не показывает лишних ссылок.

- [ ] **Step 3: Commit**

```bash
git add artifact/index.html
git commit -m "feat: показ статуса и ссылок на скачивание готовых файлов мероприятия"
```

---

## Итог MVP

После выполнения задач 1–12: оператор создаёт бриф в артефакте (Task 10) → я подбираю 90 песен из `Банк песен.xlsx` под тему и присылаю на утверждение (Task 2–3) → после утверждения `export_selection` фиксирует список (Task 4) → `run_pipeline` (Task 11) одним прогоном выдаёт готовый файл бланков (Task 5) и готовую презентацию (Task 6–8) → оба файла загружаются в артефакт, где оператор может их скачать (Task 12). Визуал (афиши/сторис, Figma-плагин, Higgsfield) — отдельный план фазы 2.
