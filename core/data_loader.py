"""Загрузка данных из Excel"""
import pandas as pd
from pathlib import Path
from typing import Tuple
from utils.logger import get_logger

logger = get_logger(__name__)

def load_excel_file(file_path: str, target_sheet: str, source_sheet: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    file_path_obj = Path(file_path)
    if not file_path_obj.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")

    logger.info(f"Загрузка файла: {file_path}")
    try:
        df_target = pd.read_excel(file_path, sheet_name=target_sheet)
        df_source = pd.read_excel(file_path, sheet_name=source_sheet)
        logger.info(f"Загружена вкладка '{target_sheet}': {len(df_target)} строк, {len(df_target.columns)} столбцов")
        logger.info(f"Загружена вкладка '{source_sheet}': {len(df_source)} строк, {len(df_source.columns)} столбцов")

        if len(df_target.columns) < 22:
            raise ValueError(f"Недостаточно столбцов в целевой вкладке")
        if len(df_source.columns) < 46:
            raise ValueError(f"Недостаточно столбцов в справочной вкладке")

        return df_target, df_source
    except Exception as e:
        logger.error(f"Ошибка при загрузке: {e}")
        raise
