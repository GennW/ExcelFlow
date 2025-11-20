"""Загрузка данных из Excel (оптимизация производительности)"""
import pandas as pd
from pathlib import Path
from typing import Tuple
from utils.logger import get_logger

logger = get_logger(__name__)

def load_excel_file(file_path: str, target_sheet: str, source_sheet: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Загружает целевую и справочную таблицы через pandas с оптимизацией.
    
    Целевая таблица: заголовки на строке 11 (header=10), данные начинаются со строки 12.
    Справочная таблица: стандартная загрузка с автоопределением типов для числовых операций.
    
    Args:
        file_path: Путь к Excel-файлу
        target_sheet: Название целевого листа
        source_sheet: Название справочного листа
    
    Returns:
        Tuple[pd.DataFrame, pd.DataFrame]: Целевая и справочная таблицы
    """
    file_path_obj = Path(file_path)
    if not file_path_obj.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")

    logger.info(f"Загрузка файла: {file_path}")
    
    try:
        # Загружаем целевую таблицу: заголовки в строке 11 (header=10 в 0-based индексации)
        df_target = pd.read_excel(
            file_path,
            sheet_name=target_sheet,
            header=10  # Строка 11 в Excel = индекс 10 в pandas
        )
        
        # Загружаем справочную таблицу с автоопределением типов
        # НЕ используем dtype=str, чтобы числовые колонки остались числовыми
        df_source = pd.read_excel(
            file_path,
            sheet_name=source_sheet,
            header=0
        )
        
        logger.info(f"Загружена вкладка '{target_sheet}': {len(df_target)} строк, {len(df_target.columns)} столбцов")
        logger.info(f"Загружена вкладка '{source_sheet}': {len(df_source)} строк, {len(df_source.columns)} столбцов")
        
        # Проверка количества столбцов
        if len(df_target.columns) < 22:
            raise ValueError(f"Недостаточно столбцов в целевой вкладке (требуется минимум 22, найдено {len(df_target.columns)})")
        if len(df_source.columns) < 46:
            raise ValueError(f"Недостаточно столбцов в справочной вкладке (требуется минимум 46, найдено {len(df_source.columns)})")

        return df_target, df_source

    except Exception as e:
        logger.error(f"Ошибка при загрузке: {e}")
        raise
