"""Конфигурация ExcelCostCalculator (ИСПРАВЛЕНО ПО ТЗ)"""

# Индексы столбцов (начиная с 0)
TARGET_COLUMNS = {
    'NOMENCLATURE': 18,  # Столбец S - "Номенклатура закупки"
    'DOCUMENT': 21        # Столбец V - "Документ приобретения"
}

SOURCE_COLUMNS = {
    # ⚠️ ИСПРАВЛЕНО ПО ТЗ!
    'PERIOD_QUARTER': 1,    # Столбец B - "Период месяц" (было 45!)
    'NOMENCLATURE': 8,      # Столбец I - "Номенклатура завода" (было 41!)
    'QUANTITY': 13,         # Столбец N - "Количество"
    'COST_Q': 16,          # Столбец Q - "Прямая СС на ед"
    'COST_R': 17,          # Столбец R - "Стоимость закупки НЧТ"
    'COST_X': 23,          # Столбец X - "НР на ед"
}

SHEET_NAMES = {
    'TARGET': 'СК ТПХ_1 пг',
    'SOURCE': 'ВП 2024-2025 НЧТЗ'
}

# Форматы
DATE_FORMAT_OUTPUT = '%d.%m.%Y'
PSTR_START = 44
PSTR_LENGTH = 10

# Коды ошибок
ERROR_CODES = {
    'DATE_NOT_FOUND': '#НД',
    'MANUAL_CHECK': '#РП',
}

ERROR_DESCRIPTIONS = {
    '#НД': 'Дата не найдена в документе',
    '#РП': 'Требует ручной проверки - нет совпадений в справочнике',
}

# Логирование
LOG_FORMAT = '[%(asctime)s] %(levelname)s: %(message)s'
LOG_DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

# Обработка
DEFAULT_CHUNK_SIZE = 500
GC_INTERVAL = 5
