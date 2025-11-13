import pandas as pd
import sys
from typing import Dict, Optional
from utils.logger_config import setup_logger

class DataLoader:
    """
    Класс для загрузки и валидации Excel-файла
    """
    
    def __init__(self, logger: Optional = None):
        self.logger = logger or setup_logger()
        self.required_sheets = ["СК ТПХ_1 пг", "ВП 2024-2025 НЧТЗ"]
        self.required_columns = {
            "СК ТПХ_1 пг": {
                "nomination": "S",  # Номенклатура закупки
                "document": "V",    # Документ приобретения
                "realization": "D"  # Дата реализации
            },
            "ВП 2024-2025 НЧТЗ": {
                "nomination_factory": "I",  # Номенклатура завода
                "period": "B",              # Период месяц
                "cost_direct": "AR",        # Прямая СС на ед
                "cost_purchase": "AQ",      # Стоимость закупки НЧТ
                "material_cost": "AS",      # Прямая материальная составляющая
                "counterparty": "F"         # Контрагент или Регистратор
            }
        }
    
    def load_excel_file(self, file_path: str) -> Dict[str, pd.DataFrame]:
        """
        Загружает Excel-файл и возвращает словарь с DataFrame для каждой вкладки.
        
        :param file_path: Путь к Excel-файлу
        :return: Словарь {имя_вкладки: DataFrame}
        """
        try:
            self.logger.info(f"Загрузка Excel-файла: {file_path}")
            
            # Проверяем существование файла
            try:
                with open(file_path, 'rb') as f:
                    pass
            except FileNotFoundError:
                self.logger.error(f"Файл не найден: {file_path}")
                sys.exit(1)
            
            # Загружаем все вкладки
            df_dict = pd.read_excel(file_path, sheet_name=None)
            
            # Проверяем наличие обязательных вкладок
            for sheet_name in self.required_sheets:
                if sheet_name not in df_dict:
                    self.logger.error(f"Вкладка '{sheet_name}' не найдена в файле")
                    sys.exit(2)
            
            # Проверяем наличие обязательных столбцов
            for sheet_name, columns in self.required_columns.items():
                df = df_dict[sheet_name]
                
                # Проверяем наличие столбцов по их буквенному обозначению
                for col_name, col_letter in columns.items():
                    # Преобразуем букву столбца в индекс
                    col_idx = self._column_letter_to_index(col_letter)
                    
                    if col_idx >= len(df.columns):
                        self.logger.error(f"Столбец '{col_letter}' ({col_name}) не найден на вкладке '{sheet_name}'")
                        sys.exit(3)
            
            self.logger.info(f"Файл успешно загружен. Найдено вкладок: {len(df_dict)}")
            
            return df_dict
            
        except Exception as e:
            self.logger.error(f"Ошибка при загрузке файла: {str(e)}")
            sys.exit(4)
    
    def _column_letter_to_index(self, letter: str) -> int:
        """
        Преобразует буквенное обозначение столбца в индекс.
        
        :param letter: Буквенное обозначение столбца (например, 'S')
        :return: Индекс столбца (0-базированный)
        """
        result = 0
        for char in letter.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1
    
    def validate_data(self, df_dict: Dict[str, pd.DataFrame]) -> bool:
        """
        Проверяет валидность данных в загруженных DataFrame.
        
        :param df_dict: Словарь с DataFrame
        :return: True, если данные валидны
        """
        for sheet_name, df in df_dict.items():
            if df.empty:
                self.logger.warning(f"Вкладка '{sheet_name}' пуста")
                continue
            
            self.logger.info(f"Вкладка '{sheet_name}': {len(df)} строк, {len(df.columns)} столбцов")
        
        return True