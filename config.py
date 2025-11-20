"""
Конфигурация приложения ExcelCostCalculator
"""

# Структура таблицы Excel
HEADER_ROW = 11  # Строка с заголовками столбцов
DATA_START_ROW = 12  # Первая строка с данными

# Целевая таблица "СК ТПХ_1 пг"
TARGET_COLUMNS = {
    'NOMENCLATURE': 18,
    'DOCUMENT': 21,
}

# Справочная таблица "ВП 2024-2025 НЧТЗ"
SOURCE_COLUMNS = {
    'QUANTITY': 13,
    'COST_Q': 16,
    'COST_R': 17,
    'COST_X': 23,
    'NOMENCLATURE': 41,
    'PERIOD_QUARTER': 45,
}

SHEET_NAMES = {
    'TARGET': 'СК ТПХ_1 пг',
    'SOURCE': 'ВП 2024-2025 НЧТЗ',
}

DATE_FORMAT_OUTPUT = '%d.%m.%Y'
DATE_FORMAT_INPUT = '%d.%m.%Y'
PSTR_START = 44
PSTR_LENGTH = 10

LOG_FORMAT = '[%(asctime)s] %(levelname)s: %(message)s'
LOG_DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

DEFAULT_CHUNK_SIZE = 500
GC_INTERVAL = 5
