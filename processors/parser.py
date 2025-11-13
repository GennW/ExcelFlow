import pandas as pd
from typing import Dict, Tuple
from utils.date_extractor import extract_date_from_document, determine_quarter, extract_date_from_realization
from utils.fuzzy_matcher import normalize_nomenclature, extract_nomenclature_code
from utils.logger_config import setup_logger

class DataParser:
    """
    Класс для парсинга и нормализации данных
    """
    
    def __init__(self, logger=None):
        self.logger = logger or setup_logger()
    
    def parse_target_sheet(self, df_target: pd.DataFrame) -> pd.DataFrame:
        """
        Парсит и нормализует данные во вкладке "СК ТПХ_1 пг".
        
        :param df_target: DataFrame вкладки "СК ТПХ_1 пг"
        :return: Обновленный DataFrame с извлеченными датами и нормализованными номенклатурами
        """
        # Создаем копию DataFrame для безопасной обработки
        df = df_target.copy()
        
        self.logger.info(f"Начало парсинга целевой вкладки: {len(df)} строк")
        
        # Извлекаем дату приобретения из столбца V (документ приобретения)
        df['Дата приобретения'] = df.iloc[:, 21].apply(extract_date_from_document)  # Столбец V - индекс 21
        
        # Определяем квартал на основе извлеченной даты
        df['Квартал приобретения'] = df['Дата приобретения'].apply(determine_quarter)
        
        # Нормализуем номенклатуру из столбца S с использованием векторизованной операции
        df['Номенклатура_норм'] = df.iloc[:, 18].apply(lambda x: normalize_nomenclature(str(x)) if pd.notna(x) else "")
        # Обновляем индексацию столбцов для корректного доступа к данным
        df['Номенклатура_норм'] = df['Номенклатура_норм'].replace('', pd.NA)
        
        # Извлекаем код номенклатуры из нормализованной номенклатуры
        df['Код номенклатуры'] = df['Номенклатура_норм'].apply(extract_nomenclature_code)
        
        # Извлекаем дату реализации из столбца D
        df['Дата реализации'] = df.iloc[:, 3].apply(extract_date_from_realization) # Столбец D - индекс 3
        
        # Обновляем индексацию столбцов для корректного доступа к данным
        df['Дата приобретения'] = df['Дата приобретения'].replace('', pd.NA)
        df['Дата реализации'] = df['Дата реализации'].replace('', pd.NA)
        
        # Определяем квартал на основе даты реализации, если дата приобретения не найдена
        # Векторизованное обновление квартала приобретения
        mask = (df['Квартал приобретения'].isna()) & (df['Дата реализации'].notna())
        df.loc[mask, 'Квартал приобретения'] = df.loc[mask, 'Дата реализации'].apply(determine_quarter)
        
        self.logger.info("Парсинг целевой вкладки завершен")
        
        return df
    
    def parse_source_sheet(self, df_source: pd.DataFrame) -> pd.DataFrame:
        """
        Парсит и нормализует данные во вкладке "ВП 2024-2025 НЧТЗ".
        
        :param df_source: DataFrame вкладки "ВП 2024-2025 НЧТЗ"
        :return: Обновленный DataFrame с нормализованными номенклатурами и преобразованными периодами
        """
        # Создаем копию DataFrame для безопасной обработки
        df = df_source.copy()
        
        self.logger.info(f"Начало парсинга справочной вкладки: {len(df)} строк")
        
        # Нормализуем номенклатуру из столбца I (номенклатура завода) с использованием векторизованной операции
        df['Номенклатура_норм'] = df.iloc[:, 8].apply(lambda x: normalize_nomenclature(str(x)) if pd.notna(x) else "")
        # Обновляем индексацию столбцов для корректного доступа к данным
        df['Номенклатура_норм'] = df['Номенклатура_норм'].replace('', pd.NA)
        
        # Извлекаем код номенклатуры из нормализованной номенклатуры
        df['Код номенклатуры'] = df['Номенклатура_норм'].apply(extract_nomenclature_code)
        
        # Преобразуем период месяца из столбца B в диапазон дат
        from utils.date_extractor import parse_period_to_date_range
        df['Период_начало'], df['Период_конец'] = zip(*df.iloc[:, 1].apply(parse_period_to_date_range))  # Столбец B - индекс 1
        
        # Нормализуем контрагента из столбца F с использованием векторизованной операции
        df['Контрагент_норм'] = df.iloc[:, 5].apply(lambda x: normalize_nomenclature(str(x)) if pd.notna(x) else "")
        df['Контрагент_норм'] = df['Контрагент_норм'].replace('', pd.NA)
        
        self.logger.info("Парсинг справочной вкладки завершен")
        
        return df