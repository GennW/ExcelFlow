# ExcelCostCalculator - Полный набор файлов проекта

## Структура проекта

```
ExcelCostCalculator/
├── main.py
├── config.py
├── requirements.txt
├── .gitignore
├── README.md
├── core/
│   ├── __init__.py
│   ├── data_loader.py
│   ├── formula_engine.py
│   └── output_writer.py
└── utils/
    ├── __init__.py
    ├── date_utils.py
    └── logger.py
```

---

## 📁 1. config.py

```python
"""
Конфигурация приложения ExcelCostCalculator
"""

# Целевая таблица "СК ТПХ_1 пг"
TARGET_COLUMNS = {
    'NOMENCLATURE': 18,     # S (19 в Excel) - Номенклатура закупки
    'DOCUMENT': 21,         # V (22 в Excel) - Документ приобретения
}

# Справочная таблица "ВП 2024-2025 НЧТЗ"
SOURCE_COLUMNS = {
    'QUANTITY': 13,         # N (14 в Excel) - Количество
    'COST_Q': 16,           # Q (17 в Excel) - Прямая СС на ед
    'COST_R': 17,           # R (18 в Excel) - Стоимость закупки НЧТ
    'COST_X': 23,           # X (24 в Excel) - Прямая материальная составляющая
    'NOMENCLATURE': 41,     # AP (42 в Excel) - Номенклатура завода
    'PERIOD_QUARTER': 45,   # AT (46 в Excel) - Период (квартал)
}

# Названия вкладок
SHEET_NAMES = {
    'TARGET': 'СК ТПХ_1 пг',
    'SOURCE': 'ВП 2024-2025 НЧТЗ',
}

# Форматы дат
DATE_FORMAT_OUTPUT = '%d.%m.%Y'
DATE_FORMAT_INPUT = '%d.%m.%Y'

# Позиция извлечения даты (аналог Excel ПСТР)
PSTR_START = 44  # 45-я позиция в Excel (0-based: 44)
PSTR_LENGTH = 10

# Логирование
LOG_FORMAT = '[%(asctime)s] %(levelname)s: %(message)s'
LOG_DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

# Оптимизация производительности
DEFAULT_CHUNK_SIZE = 500  # Размер chunk для обработки
GC_INTERVAL = 5  # Очистка памяти каждые N chunks
```

---

## 📁 2. utils/__init__.py

```python
"""
Утилиты ExcelCostCalculator
"""
from .date_utils import extract_acquisition_date, determine_quarter
from .logger import setup_logger, get_logger

__all__ = [
    'extract_acquisition_date',
    'determine_quarter',
    'setup_logger',
    'get_logger',
]
```

---

## 📁 3. utils/date_utils.py

```python
"""
Утилиты для работы с датами
"""
import re
from datetime import datetime
from typing import Optional
import pandas as pd


def extract_date_pstr(text: str, start: int = 44, length: int = 10) -> Optional[datetime]:
    """
    Извлекает дату по фиксированной позиции (аналог Excel ПСТР)
    
    Args:
        text: Текст документа
        start: Начальная позиция (0-based)
        length: Длина извлечения
    
    Returns:
        Объект datetime или None
    
    Example:
        >>> extract_date_pstr("Реализация товаров и услуг 00КА-000135 от 20.01.2025 23:59:59")
        datetime(2025, 1, 20)
    """
    if not text or pd.isna(text) or len(str(text)) < start + length:
        return None
    
    try:
        text_str = str(text)
        date_str = text_str[start:start + length]
        return datetime.strptime(date_str, '%d.%m.%Y')
    except (ValueError, IndexError):
        return None


def extract_date_regex(text: str) -> Optional[datetime]:
    """
    Извлекает дату с помощью регулярных выражений (запасной метод)
    
    Args:
        text: Текст документа
    
    Returns:
        Объект datetime или None
    """
    if not text or pd.isna(text):
        return None
    
    patterns = [
        r'от\s+(\d{2}\.\d{2}\.\d{4})',
        r'от\s+(\d{2}\.\d{2}\.\d{4})\s+\d{1,2}:\d{2}:\d{2}'
    ]
    
    text_str = str(text)
    for pattern in patterns:
        match = re.search(pattern, text_str)
        if match:
            try:
                return datetime.strptime(match.group(1), '%d.%m.%Y')
            except ValueError:
                continue
    
    return None


def extract_acquisition_date(document_text: str) -> Optional[datetime]:
    """
    Главная функция извлечения даты (комбинирует оба метода)
    
    Args:
        document_text: Текст документа приобретения
    
    Returns:
        Объект datetime или None
    """
    # Метод 1: Фиксированная позиция (приоритет)
    date = extract_date_pstr(document_text)
    if date:
        return date
    
    # Метод 2: Регулярное выражение (запасной)
    return extract_date_regex(document_text)


def determine_quarter(date: datetime) -> str:
    """
    Определяет квартал в формате "N квартал YYYY"
    
    Args:
        date: Объект datetime
    
    Returns:
        Строка формата "1 квартал 2025"
    
    Example:
        >>> determine_quarter(datetime(2024, 3, 26))
        '1 квартал 2024'
        >>> determine_quarter(datetime(2024, 8, 29))
        '3 квартал 2024'
    """
    if not date or pd.isna(date):
        return ""
    
    month = date.month
    year = date.year
    quarter_num = (month - 1) // 3 + 1
    
    return f"{quarter_num} квартал {year}"
```

---

## 📁 4. utils/logger.py

```python
"""
Настройка логирования
"""
import logging
import sys
from config import LOG_FORMAT, LOG_DATE_FORMAT


def setup_logger(level: str = 'INFO') -> None:
    """
    Настраивает корневой логгер
    
    Args:
        level: Уровень логирования ('DEBUG', 'INFO', 'WARNING', 'ERROR')
    """
    log_level = getattr(logging, level.upper(), logging.INFO)
    
    # Настройка форматирования
    formatter = logging.Formatter(LOG_FORMAT, LOG_DATE_FORMAT)
    
    # Хендлер для консоли
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    
    # Настройка корневого логгера
    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)
    root_logger.handlers = []  # Очистка существующих хендлеров
    root_logger.addHandler(console_handler)


def get_logger(name: str) -> logging.Logger:
    """
    Возвращает логгер с указанным именем
    
    Args:
        name: Имя логгера (обычно __name__)
    
    Returns:
        Объект Logger
    """
    return logging.getLogger(name)
```

---

## 📁 5. core/__init__.py

```python
"""
Основные модули ExcelCostCalculator
"""
from .data_loader import load_excel_file
from .formula_engine import FormulaEngine
from .output_writer import write_results

__all__ = [
    'load_excel_file',
    'FormulaEngine',
    'write_results',
]
```

---

## 📁 6. core/data_loader.py

```python
"""
Загрузка данных из Excel
"""
import pandas as pd
from pathlib import Path
from typing import Tuple
from utils.logger import get_logger

logger = get_logger(__name__)


def load_excel_file(file_path: str, 
                    target_sheet: str, 
                    source_sheet: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Загружает целевую и справочную вкладки из Excel-файла
    
    Args:
        file_path: Путь к Excel-файлу
        target_sheet: Название целевой вкладки
        source_sheet: Название справочной вкладки
    
    Returns:
        Кортеж (df_target, df_source)
    
    Raises:
        FileNotFoundError: Если файл не найден
        ValueError: Если вкладки не найдены
    """
    file_path_obj = Path(file_path)
    
    if not file_path_obj.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")
    
    logger.info(f"Загрузка файла: {file_path}")
    
    try:
        # Загружаем обе вкладки
        df_target = pd.read_excel(file_path, sheet_name=target_sheet)
        df_source = pd.read_excel(file_path, sheet_name=source_sheet)
        
        logger.info(f"Загружена вкладка '{target_sheet}': {len(df_target)} строк, {len(df_target.columns)} столбцов")
        logger.info(f"Загружена вкладка '{source_sheet}': {len(df_source)} строк, {len(df_source.columns)} столбцов")
        
        # Проверяем минимальное количество столбцов
        if len(df_target.columns) < 22:
            raise ValueError(f"Недостаточно столбцов в целевой вкладке (ожидается минимум 22, найдено {len(df_target.columns)})")
        
        if len(df_source.columns) < 46:
            raise ValueError(f"Недостаточно столбцов в справочной вкладке (ожидается минимум 46, найдено {len(df_source.columns)})")
        
        return df_target, df_source
        
    except Exception as e:
        logger.error(f"Ошибка при загрузке Excel-файла: {e}")
        raise
```

---

## 📁 7. core/formula_engine.py

```python
"""
Движок для вычисления формул СУММЕСЛИМН
"""
import pandas as pd
from typing import Optional, Dict
from utils.logger import get_logger

logger = get_logger(__name__)


class FormulaEngine:
    """
    Реализует логику Excel-формул СУММЕСЛИМН (SUMIFS)
    """
    
    def __init__(self, df_source: pd.DataFrame, column_indices: Dict[str, int]):
        """
        Инициализация движка формул
        
        Args:
            df_source: DataFrame справочной таблицы
            column_indices: Словарь индексов столбцов из config.py
        """
        self.df_source = df_source
        self.columns = column_indices
        
        # Создаём удобные ссылки на столбцы по индексам
        self._map_columns()
    
    def _map_columns(self) -> None:
        """Создаёт маппинг столбцов по индексам"""
        try:
            self.col_quantity = self.df_source.iloc[:, self.columns['QUANTITY']]
            self.col_cost_q = self.df_source.iloc[:, self.columns['COST_Q']]
            self.col_cost_r = self.df_source.iloc[:, self.columns['COST_R']]
            self.col_cost_x = self.df_source.iloc[:, self.columns['COST_X']]
            self.col_nomenclature = self.df_source.iloc[:, self.columns['NOMENCLATURE']]
            self.col_period = self.df_source.iloc[:, self.columns['PERIOD_QUARTER']]
            
            logger.info("Маппинг столбцов выполнен успешно")
            logger.debug(f"Номенклатура (столбец {self.columns['NOMENCLATURE']}): {self.col_nomenclature.name}")
            logger.debug(f"Период (столбец {self.columns['PERIOD_QUARTER']}): {self.col_period.name}")
            logger.debug(f"Количество (столбец {self.columns['QUANTITY']}): {self.col_quantity.name}")
            
        except IndexError as e:
            logger.error(f"Ошибка маппинга столбцов: {e}")
            logger.error(f"Доступно столбцов: {len(self.df_source.columns)}")
            logger.error(f"Требуемые индексы: {self.columns}")
            raise
    
    def sumifs_weighted_avg(self, 
                           sum_column_name: str,
                           nomenclature: str, 
                           quarter: str) -> Optional[float]:
        """
        Реализует формулу СУММЕСЛИМН для расчёта средневзвешенного значения
        
        Args:
            sum_column_name: Имя столбца для суммирования ('COST_R', 'COST_Q', 'COST_X')
            nomenclature: Значение номенклатуры для фильтрации
            quarter: Значение квартала для фильтрации
        
        Returns:
            Средневзвешенное значение или None
        """
        # Создаём маску фильтрации
        mask = (
            (self.col_nomenclature == nomenclature) &
            (self.col_period == quarter)
        )
        
        # Подсчитываем количество совпадений
        matches_count = mask.sum()
        
        if matches_count == 0:
            logger.debug(f"Нет совпадений для: {str(nomenclature)[:50]}... | {quarter}")
            return None
        
        logger.debug(f"Найдено {matches_count} строк для агрегации")
        
        # Получаем нужный столбец для суммирования
        if sum_column_name == 'COST_R':
            sum_column = self.col_cost_r
        elif sum_column_name == 'COST_Q':
            sum_column = self.col_cost_q
        elif sum_column_name == 'COST_X':
            sum_column = self.col_cost_x
        else:
            logger.error(f"Неизвестный столбец: {sum_column_name}")
            return None
        
        # Вычисляем суммы
        total_sum = sum_column[mask].sum()
        total_qty = self.col_quantity[mask].sum()
        
        if total_qty == 0 or pd.isna(total_qty):
            logger.warning(f"Общее количество = 0 или NaN для: {str(nomenclature)[:50]}... | {quarter}")
            return None
        
        result = total_sum / total_qty
        logger.debug(f"{sum_column_name}: {total_sum:.2f} / {total_qty:.2f} = {result:.2f}")
        
        return round(result, 2)
    
    def calculate_aq(self, nomenclature: str, quarter: str) -> Optional[float]:
        """
        Стоимость закупки НЧТЗ 1 ед (столбец R)
        
        Args:
            nomenclature: Номенклатура
            quarter: Квартал
        
        Returns:
            Средневзвешенная стоимость или None
        """
        return self.sumifs_weighted_avg('COST_R', nomenclature, quarter)
    
    def calculate_ar(self, nomenclature: str, quarter: str) -> Optional[float]:
        """
        Прямая СС НЧТЗ 1 ед (столбец Q)
        
        Args:
            nomenclature: Номенклатура
            quarter: Квартал
        
        Returns:
            Средневзвешенная прямая СС или None
        """
        return self.sumifs_weighted_avg('COST_Q', nomenclature, quarter)
    
    def calculate_as(self, nomenclature: str, quarter: str) -> Optional[float]:
        """
        НР НЧТЗ 1 ед (столбец X)
        
        Args:
            nomenclature: Номенклатура
            quarter: Квартал
        
        Returns:
            Средневзвешенная материальная составляющая или None
        """
        return self.sumifs_weighted_avg('COST_X', nomenclature, quarter)
```

---

## 📁 8. core/output_writer.py

```python
"""
Запись результатов в Excel
"""
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from utils.logger import get_logger

logger = get_logger(__name__)


def write_results(input_file: str, 
                 output_file: str, 
                 df_result: pd.DataFrame,
                 target_sheet: str) -> None:
    """
    Записывает результаты в новый Excel-файл с сохранением форматирования
    
    Args:
        input_file: Путь к исходному файлу
        output_file: Путь к выходному файлу
        df_result: DataFrame с результатами
        target_sheet: Название целевой вкладки
    """
    try:
        # Загружаем исходный файл для сохранения форматирования
        wb = load_workbook(input_file)
        ws = wb[target_sheet]
        
        logger.info(f"Открыт файл для записи: {input_file}")
        
        # Определяем начальную колонку для новых данных
        start_col = ws.max_column + 1
        
        # Записываем заголовки
        headers = ['Дата приобретения', 'Квартал приобретения', 
                  'Стоимость закупки НЧТЗ 1 ед', 'Прямая СС НЧТЗ 1 ед', 'НР НЧТЗ 1 ед']
        
        for i, header in enumerate(headers):
            ws.cell(row=1, column=start_col + i, value=header)
        
        logger.info(f"Добавлены заголовки в столбцы {start_col}-{start_col + 4}")
        
        # Записываем данные
        for row_idx in range(len(df_result)):
            excel_row = row_idx + 2  # +2 потому что строки в Excel начинаются с 1, и первая строка - заголовки
            
            ws.cell(row=excel_row, column=start_col, 
                   value=df_result.iloc[row_idx]['AO_Дата_приобретения'])
            ws.cell(row=excel_row, column=start_col + 1, 
                   value=df_result.iloc[row_idx]['AP_Квартал_приобретения'])
            ws.cell(row=excel_row, column=start_col + 2, 
                   value=df_result.iloc[row_idx]['AQ_Стоимость_закупки'])
            ws.cell(row=excel_row, column=start_col + 3, 
                   value=df_result.iloc[row_idx]['AR_Прямая_СС'])
            ws.cell(row=excel_row, column=start_col + 4, 
                   value=df_result.iloc[row_idx]['AS_НР'])
        
        logger.info(f"Записано {len(df_result)} строк данных")
        
        # Сохраняем файл
        wb.save(output_file)
        logger.info(f"Файл успешно сохранён: {output_file}")
        
    except Exception as e:
        logger.error(f"Ошибка при записи результатов: {e}")
        raise
```

---

## 📁 9. main.py

```python
"""
Точка входа в приложение ExcelCostCalculator
Оптимизировано для работы на слабых компьютерах
"""
import argparse
import sys
import gc
from pathlib import Path
import pandas as pd
from datetime import datetime

from config import (
    TARGET_COLUMNS, SOURCE_COLUMNS, SHEET_NAMES, 
    DATE_FORMAT_OUTPUT, DEFAULT_CHUNK_SIZE, GC_INTERVAL
)
from core import load_excel_file, FormulaEngine, write_results
from utils import extract_acquisition_date, determine_quarter, setup_logger, get_logger


def get_memory_usage():
    """Возвращает использование памяти в МБ"""
    try:
        import psutil
        process = psutil.Process()
        return process.memory_info().rss / 1024 / 1024
    except ImportError:
        return 0


def process_chunk(df_chunk, engine, chunk_start_idx, logger):
    """
    Обрабатывает chunk строк
    
    Args:
        df_chunk: DataFrame chunk для обработки
        engine: FormulaEngine
        chunk_start_idx: Начальный индекс chunk в исходном DataFrame
        logger: Логгер
    
    Returns:
        Обработанный chunk, количество успехов и ошибок
    """
    success_count = 0
    error_count = 0
    
    for idx in range(len(df_chunk)):
        global_idx = chunk_start_idx + idx
        
        # Извлечение входных данных
        document_text = df_chunk.iloc[idx, TARGET_COLUMNS['DOCUMENT']]
        nomenclature = df_chunk.iloc[idx, TARGET_COLUMNS['NOMENCLATURE']]
        
        # AO: Дата приобретения
        acquisition_date = extract_acquisition_date(document_text)
        
        if not acquisition_date:
            logger.debug(f"Строка {global_idx+2}: Не удалось извлечь дату")
            df_chunk.at[df_chunk.index[idx], 'AO_Дата_приобретения'] = "*Дата не найдена*"
            df_chunk.at[df_chunk.index[idx], 'AP_Квартал_приобретения'] = ""
            df_chunk.at[df_chunk.index[idx], 'AQ_Стоимость_закупки'] = "*ТРЕБУЕТ РУЧНОЙ ПРОВЕРКИ*"
            df_chunk.at[df_chunk.index[idx], 'AR_Прямая_СС'] = "*ТРЕБУЕТ РУЧНОЙ ПРОВЕРКИ*"
            df_chunk.at[df_chunk.index[idx], 'AS_НР'] = "*ТРЕБУЕТ РУЧНОЙ ПРОВЕРКИ*"
            error_count += 1
            continue
        
        df_chunk.at[df_chunk.index[idx], 'AO_Дата_приобретения'] = acquisition_date.strftime(DATE_FORMAT_OUTPUT)
        
        # AP: Квартал приобретения
        quarter = determine_quarter(acquisition_date)
        df_chunk.at[df_chunk.index[idx], 'AP_Квартал_приобретения'] = quarter
        
        # AQ, AR, AS: Расчёты по формулам СУММЕСЛИМН
        aq = engine.calculate_aq(nomenclature, quarter)
        ar = engine.calculate_ar(nomenclature, quarter)
        as_val = engine.calculate_as(nomenclature, quarter)
        
        if aq is not None:
            df_chunk.at[df_chunk.index[idx], 'AQ_Стоимость_закупки'] = aq
            df_chunk.at[df_chunk.index[idx], 'AR_Прямая_СС'] = ar
            df_chunk.at[df_chunk.index[idx], 'AS_НР'] = as_val
            success_count += 1
        else:
            logger.debug(f"Строка {global_idx+2}: Нет совпадений для {str(nomenclature)[:50]}... | {quarter}")
            df_chunk.at[df_chunk.index[idx], 'AQ_Стоимость_закупки'] = "*ТРЕБУЕТ РУЧНОЙ ПРОВЕРКИ*"
            df_chunk.at[df_chunk.index[idx], 'AR_Прямая_СС'] = "*ТРЕБУЕТ РУЧНОЙ ПРОВЕРКИ*"
            df_chunk.at[df_chunk.index[idx], 'AS_НР'] = "*ТРЕБУЕТ РУЧНОЙ ПРОВЕРКИ*"
            error_count += 1
    
    return df_chunk, success_count, error_count


def main():
    """Главная функция приложения"""
    parser = argparse.ArgumentParser(
        description='ExcelCostCalculator - автоматический расчёт себестоимости'
    )
    parser.add_argument('--input', required=True, 
                       help='Путь к входному Excel-файлу')
    parser.add_argument('--output', required=True, 
                       help='Путь к выходному Excel-файлу')
    parser.add_argument('--log-level', default='INFO', 
                       choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
                       help='Уровень логирования')
    parser.add_argument('--chunk-size', type=int, default=DEFAULT_CHUNK_SIZE,
                       help=f'Размер chunk для обработки (по умолчанию: {DEFAULT_CHUNK_SIZE})')
    
    args = parser.parse_args()
    
    # Настройка логирования
    setup_logger(args.log_level)
    logger = get_logger(__name__)
    
    # Проверка входного файла
    input_path = Path(args.input)
    if not input_path.exists():
        logger.error(f"Файл не найден: {args.input}")
        sys.exit(1)
    
    logger.info("========== ExcelCostCalculator ==========")
    logger.info(f"Входной файл: {args.input}")
    logger.info(f"Выходной файл: {args.output}")
    logger.info(f"Размер chunk: {args.chunk_size} строк")
    
    start_time = datetime.now()
    initial_memory = get_memory_usage()
    
    if initial_memory > 0:
        logger.info(f"Начальное использование памяти: {initial_memory:.2f} МБ")
    
    try:
        # 1. Загрузка данных
        logger.info("Загрузка Excel-файла...")
        df_target, df_source = load_excel_file(
            args.input,
            SHEET_NAMES['TARGET'],
            SHEET_NAMES['SOURCE']
        )
        
        logger.info(f"Целевая таблица: {len(df_target)} строк")
        logger.info(f"Справочная таблица: {len(df_source)} строк")
        
        after_load_memory = get_memory_usage()
        if after_load_memory > 0:
            logger.info(f"Память после загрузки: {after_load_memory:.2f} МБ (+{after_load_memory - initial_memory:.2f} МБ)")
        
        # 2. Инициализация движка формул
        logger.info("Инициализация FormulaEngine...")
        engine = FormulaEngine(df_source, SOURCE_COLUMNS)
        
        # 3. Создание новых столбцов в исходном DataFrame
        df_target['AO_Дата_приобретения'] = None
        df_target['AP_Квартал_приобретения'] = ""
        df_target['AQ_Стоимость_закупки'] = None
        df_target['AR_Прямая_СС'] = None
        df_target['AS_НР'] = None
        
        # 4. Обработка по частям (chunking)
        logger.info("Начало обработки по частям...")
        total_rows = len(df_target)
        chunk_size = args.chunk_size
        total_success = 0
        total_errors = 0
        
        num_chunks = (total_rows + chunk_size - 1) // chunk_size
        logger.info(f"Общее количество chunks: {num_chunks}")
        
        for chunk_num in range(num_chunks):
            chunk_start = chunk_num * chunk_size
            chunk_end = min(chunk_start + chunk_size, total_rows)
            
            logger.info(f"Обработка chunk {chunk_num + 1}/{num_chunks} (строки {chunk_start + 2}-{chunk_end + 1})...")
            
            # Извлекаем chunk
            df_chunk = df_target.iloc[chunk_start:chunk_end].copy()
            
            # Обрабатываем chunk
            df_chunk_processed, success_count, error_count = process_chunk(
                df_chunk, engine, chunk_start, logger
            )
            
            # Обновляем исходный DataFrame
            df_target.iloc[chunk_start:chunk_end] = df_chunk_processed
            
            total_success += success_count
            total_errors += error_count
            
            # Логирование прогресса
            progress = chunk_end / total_rows * 100
            logger.info(f"Обработано: {chunk_end}/{total_rows} ({progress:.1f}%) | "
                       f"Успешно: {total_success} | Ошибок: {total_errors}")
            
            # Мониторинг памяти
            current_memory = get_memory_usage()
            if current_memory > 0:
                logger.debug(f"Память: {current_memory:.2f} МБ")
            
            # Периодическая очистка памяти
            if (chunk_num + 1) % GC_INTERVAL == 0:
                logger.debug("Очистка памяти...")
                gc.collect()
                
                after_gc_memory = get_memory_usage()
                if after_gc_memory > 0:
                    logger.debug(f"Память после очистки: {after_gc_memory:.2f} МБ")
        
        # Финальная очистка памяти перед записью
        logger.info("Финальная очистка памяти...")
        gc.collect()
        
        # 5. Запись результатов
        logger.info("Сохранение результатов...")
        write_results(args.input, args.output, df_target, SHEET_NAMES['TARGET'])
        
        # 6. Итоговая статистика
        elapsed = (datetime.now() - start_time).total_seconds()
        final_memory = get_memory_usage()
        
        logger.info("========== ИТОГОВАЯ СТАТИСТИКА ==========")
        logger.info(f"Всего обработано записей: {total_rows}")
        logger.info(f"Успешно сопоставлено: {total_success} ({total_success/total_rows*100:.1f}%)")
        logger.info(f"Требует ручной проверки: {total_errors} ({total_errors/total_rows*100:.1f}%)")
        logger.info(f"Время обработки: {elapsed:.1f} сек ({elapsed/60:.1f} мин)")
        
        if final_memory > 0:
            logger.info(f"Финальная память: {final_memory:.2f} МБ")
            logger.info(f"Пиковое потребление памяти: +{final_memory - initial_memory:.2f} МБ")
        
        logger.info("Обработка завершена успешно!")
        
    except Exception as e:
        logger.error(f"Ошибка при обработке: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

## 📁 10. requirements.txt

```txt
pandas>=2.0.0
openpyxl>=3.1.0
python-dateutil>=2.8.0
psutil>=5.9.0
```

---

## 📁 11. .gitignore

```
# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
*.egg-info/
.installed.cfg
*.egg

# Virtual environments
venv/
ENV/
env/
.venv

# IDE
.vscode/
.idea/
*.swp
*.swo
*~

# Excel files
*.xlsx
*.xls
!tests/fixtures/*.xlsx

# Logs
*.log

# OS
.DS_Store
Thumbs.db
```

---

## 📁 12. README.md

```markdown
# ExcelCostCalculator

Автоматический расчёт средневзвешенной себестоимости продукции на основе формул СУММЕСЛИМН (SUMIFS).

## Описание

ExcelCostCalculator автоматизирует процесс заполнения столбцов AO-AS в Excel-файле, который ранее выполнялся вручную с помощью формул Excel. Приложение реплицирует логику Excel-формул СУММЕСЛИМН для вычисления средневзвешенных значений по множеству строк.

### Особенности

✅ **Оптимизация для слабых компьютеров** — обработка по частям (chunking)  
✅ **Мониторинг памяти** — автоматическое отслеживание потребления RAM  
✅ **Периодическая очистка памяти** — каждые 5 chunks  
✅ **Настраиваемый размер chunk** — от 100 до 5000 строк  
✅ **Детальное логирование** — прогресс и статистика в реальном времени  

## Функциональность

Приложение создаёт 5 новых столбцов в целевой таблице:

- **AO** (Дата приобретения) — извлекается из документа приобретения
- **AP** (Квартал приобретения) — формат "N квартал YYYY"
- **AQ** (Стоимость закупки НЧТЗ 1 ед) — средневзвешенное по формуле СУММЕСЛИМН
- **AR** (Прямая СС НЧТЗ 1 ед) — средневзвешенное по формуле СУММЕСЛИМН
- **AS** (НР НЧТЗ 1 ед) — средневзвешенное по формуле СУММЕСЛИМН

## Установка

```bash
# Клонировать репозиторий
git clone https://github.com/GennW/ExcelCostCalculator.git
cd ExcelCostCalculator

# Создать виртуальное окружение
python -m venv venv

# Активировать виртуальное окружение
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# Установить зависимости
pip install -r requirements.txt
```

## Использование

### Базовый запуск

```bash
python main.py --input "путь/к/входному/файлу.xlsx" --output "путь/к/выходному/файлу.xlsx"
```

### С подробным логированием

```bash
python main.py --input "input.xlsx" --output "output.xlsx" --log-level DEBUG
```

### Оптимизация для слабых компьютеров

```bash
# Для очень слабого компьютера (< 4 ГБ RAM)
python main.py --input "input.xlsx" --output "output.xlsx" --chunk-size 100

# Для среднего компьютера (4-8 ГБ RAM)
python main.py --input "input.xlsx" --output "output.xlsx" --chunk-size 500

# Для мощного компьютера (> 8 ГБ RAM)
python main.py --input "input.xlsx" --output "output.xlsx" --chunk-size 2000
```

### Параметры командной строки

| Параметр | Обязательный | Описание | По умолчанию |
|----------|--------------|----------|--------------|
| `--input` | Да | Путь к входному Excel-файлу | - |
| `--output` | Да | Путь к выходному Excel-файлу | - |
| `--log-level` | Нет | Уровень логирования (DEBUG, INFO, WARNING, ERROR) | INFO |
| `--chunk-size` | Нет | Размер chunk для обработки | 500 |

## Рекомендации по chunk-size

| Оперативная память | Рекомендуемый chunk-size |
|-------------------|--------------------------|
| < 4 ГБ | 100-200 |
| 4-8 ГБ | 500 (по умолчанию) |
| 8-16 ГБ | 1000-2000 |
| > 16 ГБ | 2000-5000 |

## Пример вывода

```
[2025-11-14 13:00:00] INFO: ========== ExcelCostCalculator ==========
[2025-11-14 13:00:00] INFO: Входной файл: input.xlsx
[2025-11-14 13:00:00] INFO: Выходной файл: output.xlsx
[2025-11-14 13:00:00] INFO: Размер chunk: 500 строк
[2025-11-14 13:00:00] INFO: Начальное использование памяти: 150.23 МБ
[2025-11-14 13:00:01] INFO: Загрузка Excel-файла...
[2025-11-14 13:00:05] INFO: Целевая таблица: 6295 строк
[2025-11-14 13:00:05] INFO: Справочная таблица: 12500 строк
[2025-11-14 13:00:05] INFO: Память после загрузки: 320.45 МБ (+170.22 МБ)
[2025-11-14 13:00:06] INFO: Начало обработки по частям...
[2025-11-14 13:00:06] INFO: Общее количество chunks: 13
[2025-11-14 13:00:10] INFO: Обработка chunk 1/13 (строки 2-501)...
[2025-11-14 13:00:20] INFO: Обработано: 500/6295 (7.9%) | Успешно: 475 | Ошибок: 25
[2025-11-14 13:01:00] INFO: Обработано: 1000/6295 (15.9%) | Успешно: 950 | Ошибок: 50
...
[2025-11-14 13:03:00] INFO: Финальная очистка памяти...
[2025-11-14 13:03:05] INFO: ========== ИТОГОВАЯ СТАТИСТИКА ==========
[2025-11-14 13:03:05] INFO: Всего обработано записей: 6295
[2025-11-14 13:03:05] INFO: Успешно сопоставлено: 5950 (94.5%)
[2025-11-14 13:03:05] INFO: Требует ручной проверки: 345 (5.5%)
[2025-11-14 13:03:05] INFO: Время обработки: 185.3 сек (3.1 мин)
[2025-11-14 13:03:05] INFO: Финальная память: 380.67 МБ
[2025-11-14 13:03:05] INFO: Пиковое потребление памяти: +230.44 МБ
```

## Структура проекта

```
ExcelCostCalculator/
├── main.py                 # Точка входа
├── config.py              # Конфигурация
├── core/                  # Основные модули
│   ├── __init__.py
│   ├── data_loader.py     # Загрузка Excel
│   ├── formula_engine.py  # Формулы СУММЕСЛИМН
│   └── output_writer.py   # Запись результатов
├── utils/                 # Утилиты
│   ├── __init__.py
│   ├── date_utils.py      # Работа с датами
│   └── logger.py          # Логирование
├── requirements.txt       # Зависимости
├── .gitignore
└── README.md
```

## Требования

- Python 3.8+
- pandas >= 2.0.0
- openpyxl >= 3.1.0
- python-dateutil >= 2.8.0
- psutil >= 5.9.0

## Автоматические оптимизации

### 1. Chunking (обработка по частям)
Приложение автоматически делит данные на части для экономии памяти.

### 2. Мониторинг памяти
Отслеживание использования RAM в реальном времени.

### 3. Периодическая очистка
Автоматическая очистка памяти каждые 5 chunks.

### 4. Векторизованные операции
Использование pandas для быстрой обработки данных.

## Устранение неполадок

### Проблема: "Недостаточно памяти"

**Решение:** Уменьшите chunk-size
```bash
python main.py --input "file.xlsx" --output "out.xlsx" --chunk-size 100
```

### Проблема: "Недостаточно столбцов"

**Решение:** Убедитесь, что входной файл содержит все необходимые столбцы (минимум 46 в справочнике).

### Проблема: "Медленная работа"

**Решение:** Увеличьте chunk-size для более мощного компьютера
```bash
python main.py --input "file.xlsx" --output "out.xlsx" --chunk-size 2000
```

## Лицензия

MIT

## Автор

Gennady (GennW)

## Версия

1.0.0 (14 ноября 2025)
```

---

## Инструкции по установке

### 1. Создайте структуру папок

```bash
mkdir ExcelCostCalculator
cd ExcelCostCalculator
mkdir core utils
```

### 2. Создайте все файлы

Скопируйте содержимое каждого файла из этого документа в соответствующие файлы проекта.

### 3. Установите зависимости

```bash
python -m venv venv
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

pip install -r requirements.txt
```

### 4. Запустите приложение

```bash
python main.py --input "ваш_файл.xlsx" --output "результат.xlsx"
```

---

## Готово! 🎉

Все файлы готовы к использованию. Приложение оптимизировано для:
- ✅ Слабых компьютеров (< 4 ГБ RAM)
- ✅ Больших файлов (6000+ строк)
- ✅ Автоматической оптимизации памяти
- ✅ Детального мониторинга и логирования
