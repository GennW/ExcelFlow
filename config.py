"""Конфигурация ExcelCostCalculator (ИСПРАВЛЕННАЯ ВЕРСИЯ)"""

# Индексы столбцов (начиная с 0)
TARGET_COLUMNS = {
    'NOMENCLATURE': 18,  # Столбец S
    'DOCUMENT': 21        # Столбец V
}

SOURCE_COLUMNS = {
    'QUANTITY': 13,         # Столбец N
    'COST_Q': 16,          # Столбец Q
    'COST_R': 17,          # Столбец R
    'COST_X': 23,          # Столбец X
    'NOMENCLATURE': 41,    # Столбец AP
    'PERIOD_QUARTER': 45   # Столбец AT
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
