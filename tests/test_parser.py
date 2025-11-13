import unittest
import pandas as pd
from processors.parser import DataParser
from utils.date_extractor import extract_date_from_document, determine_quarter
from utils.fuzzy_matcher import normalize_nomenclature, extract_nomenclature_code


class TestParser(unittest.TestCase):
    
    def setUp(self):
        self.parser = DataParser()
    
    def test_nomenclature_normalization(self):
        """Тест нормализации названия номенклатуры"""
        test_cases = [
            ("  Подвески штангового глубинного насоса ПШГН Х-70-5  ", "подвески штангового глубинного насоса пшгн х-70-5"),
            ("Башмак   колонный   вращающийся", "башмак колонный вращающийся"),
            ("   Муфта резьбовая ОНГ.216.00.000-01-032   ", "муфта резьбовая онг.216.00.000-01-032"),
            ("", ""),
            (None, "")
        ]
        
        for input_val, expected in test_cases:
            with self.subTest(input_val=input_val):
                result = normalize_nomenclature(input_val)
                self.assertEqual(result, expected)
    
    def test_quarter_determination(self):
        """Тест корректного формата квартала"""
        from datetime import datetime
        test_cases = [
            (datetime(2024, 10, 11), "4 квартал 2024"),
            (datetime(2024, 8, 29), "3 квартал 2024"),
            (datetime(2025, 2, 15), "1 квартал 2025"),
            (datetime(2025, 5, 15), "2 квартал 2025"),
            (datetime(2025, 8, 15), "3 квартал 2025"),
            (datetime(2025, 11, 15), "4 квартал 2025"),
            (None, None)
        ]
        
        for input_val, expected in test_cases:
            with self.subTest(input_val=input_val):
                result = determine_quarter(input_val)
                self.assertEqual(result, expected)
    
    def test_nomenclature_code_extraction(self):
        """Тест извлечения кода номенклатуры"""
        test_cases = [
            ("Муфта резьбовая ОНГ.216.00.000-01-032 (подойди 3/4)", "ОНГ.216.00.000-01-032"),
            ("Цемент. муфта ЦМВ.32.00.000-03.4.2023 (2 шт.)", "ЦМВ.32.00.000-03.4.2023"),
            ("Подвески штангового глубинного насоса ПШГН Х-70-5", None),
            ("МШГРП.114.015-032-60,00", "МШГРП.114.015-032-60,00"),
            ("БК-Вр.114", "БК-Вр.114"),
            ("", None),
            (None, None)
        ]
        
        for input_val, expected in test_cases:
            with self.subTest(input_val=input_val):
                result = extract_nomenclature_code(input_val)
                self.assertEqual(result, expected)
    
    def test_parse_target_sheet(self):
        """Тест парсинга целевой вкладки"""
        # Создаем тестовый DataFrame
        df = pd.DataFrame({
            'A': range(5),
            'S': ['Номенклатура1', 'Номенклатура2', 'ОНГ.216.00.000-01-032', 'Номенклатура4', 'Номенклатура5'],
            'D': ['01.03.2025', '15.04.2025', '20.05.2025', 'Некорректная дата', None],
            'V': [
                'Реализация товаров и услуг 00КА-000135 от 20.01.2025 23:59:59',
                'Реализация товаров и услуг 00КА-000207 от 15.02.2025 14:30:00',
                'Запрос без даты',
                'Некорректные данные',
                'Документ от 10.06.2025'
            ]
        })
        
        result_df = self.parser.parse_target_sheet(df)
        
        # Проверяем, что новые столбцы добавлены
        self.assertIn('Дата приобретения', result_df.columns)
        self.assertIn('Квартал приобретения', result_df.columns)
        self.assertIn('Номенклатура_норм', result_df.columns)
        self.assertIn('Код номенклатуры', result_df.columns)
        self.assertIn('Дата реализации', result_df.columns)
        
        # Проверяем извлечение даты приобретения
        expected_dates = [
            pd.Timestamp(2025, 1, 20),
            pd.Timestamp(2025, 2, 15),
            None,
            None,
            pd.Timestamp(2025, 6, 10)
        ]
        
        for i, expected in enumerate(expected_dates):
            if pd.isna(expected):
                self.assertTrue(pd.isna(result_df.iloc[i]['Дата приобретения']))
            else:
                self.assertEqual(result_df.iloc[i]['Дата приобретения'], expected)
        
        # Проверяем определение квартала
        expected_quarters = ['1 квартал 2025', '1 квартал 2025', '2 квартал 2025', None, '2 квартал 2025']  # Q2 для даты реализации
        for i, expected in enumerate(expected_quarters):
            self.assertEqual(result_df.iloc[i]['Квартал приобретения'], expected)
        
        # Проверяем нормализацию номенклатуры
        self.assertEqual(result_df.iloc[0]['Номенклатура_норм'], 'номенклатура1')
        self.assertEqual(result_df.iloc[2]['Номенклатура_норм'], 'онг.216.00.000-01-032')
        
        # Проверяем извлечение кода номенклатуры
        self.assertEqual(result_df.iloc[2]['Код номенклатуры'], 'ОНГ.216.00.000-01-032')
    
    def test_parse_source_sheet(self):
        """Тест парсинга справочной вкладки"""
        # Создаем тестовый DataFrame
        df = pd.DataFrame({
            'A': range(3),
            'B': ['Январь 2025 г.', 'Февраль 2025 г.', 'Март 2025 г.'],
            'F': ['СК ТПХ ООО', 'ОО СК Татпром-холдинг', 'Другой контрагент'],
            'I': ['Номенклатура завода1', 'ОНГ.216.00.000-01-032', 'Номенклатура3']
        })
        
        result_df = self.parser.parse_source_sheet(df)
        
        # Проверяем, что новые столбцы добавлены
        self.assertIn('Номенклатура_норм', result_df.columns)
        self.assertIn('Код номенклатуры', result_df.columns)
        self.assertIn('Период_начало', result_df.columns)
        self.assertIn('Период_конец', result_df.columns)
        self.assertIn('Контрагент_норм', result_df.columns)
        
        # Проверяем нормализацию номенклатуры
        self.assertEqual(result_df.iloc[0]['Номенклатура_норм'], 'номенклатура завода1')
        self.assertEqual(result_df.iloc[1]['Номенклатура_норм'], 'онг.216.00.000-01-032')
        
        # Проверяем извлечение кода номенклатуры
        self.assertEqual(result_df.iloc[1]['Код номенклатуры'], 'ОНГ.216.00.000-01-032')
        
        # Проверяем нормализацию контрагента
        self.assertEqual(result_df.iloc[0]['Контрагент_норм'], 'ск тпх ооо')
        self.assertEqual(result_df.iloc[1]['Контрагент_норм'], 'ооо ск татпром-холдинг')


if __name__ == '__main__':
    unittest.main()