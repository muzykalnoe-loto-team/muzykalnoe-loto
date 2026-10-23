# Где лежат исходные файлы оператора (в git не хранятся). Пути — от корня проекта.
SOURCES = "Исходники"

SONG_BANK = f"{SOURCES}/Банк песен.xlsx"
BLANKS_TEMPLATE = f"{SOURCES}/Шаблон бланков.xlsm"
# Шаблон игры: каждый 3-й слот (бочонки 3, 6 … 90) — макет «Игровой слайд» под старый слайд
PRESENTATION_TEMPLATE = f"{SOURCES}/Рабочий 2 вар.pptx"
SLIDES_ARCHIVE = f"{SOURCES}/Банк всех слайдов23.01.2026"

# Готовые игры и пакеты игр (в git не хранится)
EVENTS = "Выгрузка ИИ"

# Номинации: какие были в прошлых играх (для очереди) и архив тематических слайдов.
# Архив Claude не читает — программа только дописывает в конец
NOMINATION_HISTORY = f"{EVENTS}/номинации — история.json"
NOMINATION_ARCHIVE = f"{SOURCES}/Номинация.pptx"

# Ключи и пароли (в git не хранится, значения не выводить на экран)
SECRETS_FILE = ".env"

# Старые игры и всё, что из них собирается (в git не хранится)
OLD_DESIGN = f"{SOURCES}/Старый дизайн"
OLD_BANK = f"{SOURCES}/Банк старых слайдов"
OLD_CATALOG = f"{OLD_BANK}/каталог.json"
COMMON_BANK_XLSX = f"{SOURCES}/Общий банк песен.xlsx"
COMMON_BANK_JSON = f"{SOURCES}/Общий банк песен.json"
