"""Загрузка данных из Excel (ИСПРАВЛЕННАЯ ВЕРСИЯ)"""
import pandas as pd
from pathlib import Path
from typing import Tuple
from utils.logger import get_logger

logger = get_logger(__name__)


def find_header_row(file_path: str, sheet_name: str, search_column: int = 40) -> int:
    """
    Находит строку с заголовками в Excel-таблице.

    Args:
        file_path: Путь к Excel-файлу
        sheet_name: Имя листа
        search_column: Индекс столбца для поиска (по умолчанию 40)

    Returns:
        Номер строки с заголовками (0-индекс)
    """
    logger.info(f"Поиск строки с заголовками в '{sheet_name}'...")

    # Читаем первые 20 строк без заголовков
    df_preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=20)

    # Ищем строку, содержащую "Дата приобретения"
    for idx in range(len(df_preview)):
        cell_value = df_preview.iloc[idx, search_column]
        if pd.notna(cell_value) and "Дата приобретения" in str(cell_value):
            logger.info(f"  ✅ Найдена строка {idx} (Excel строка {idx + 1})")
            return idx

    logger.warning("  ⚠️  Строка не найдена, используем 0")
    return 0


def load_excel_file(file_path: str, target_sheet: str, source_sheet: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Загружает целевую и справочную таблицы из Excel.

    Args:
        file_path: Путь к Excel-файлу
        target_sheet: Имя целевого листа
        source_sheet: Имя справочного листа

    Returns:
        Tuple[pd.DataFrame, pd.DataFrame]: Целевая и справочная таблицы
    """
    file_path_obj = Path(file_path)
    if not file_path_obj.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")

    logger.info("="*60)
    logger.info(f"Загрузка файла: {file_path}")

    # Поиск заголовков в целевой таблице
    header_row_target = find_header_row(file_path, target_sheet, search_column=40)
    df_target = pd.read_excel(file_path, sheet_name=target_sheet, header=header_row_target)
    logger.info(f"  ✅ Целевая таблица: {len(df_target)} строк, {len(df_target.columns)} столбцов")

    # Поиск заголовков в справочной таблице
    header_row_source = find_header_row(file_path, source_sheet, search_column=40)
    df_source = pd.read_excel(file_path, sheet_name=source_sheet, header=header_row_source)
    logger.info(f"  ✅ Справочная таблица: {len(df_source)} строк, {len(df_source.columns)} столбцов")

    logger.info("="*60)

    return df_target, df_source
