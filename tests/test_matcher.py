import unittest
import pandas as pd
from processors.matcher import DataMatcher
from processors.parser import DataParser


class TestMatcher(unittest.TestCase):
    
    def setUp(self):
        self.matcher = DataMatcher()
        self.parser = DataParser()
    
    def test_exact_match_with_date(self):
        """Тест точного совпадения с датой приобретения"""
        # Создаем тестовые DataFrame
        df_target = pd.DataFrame({
            'S': ['Башмак колонный вращающийся БК-Вр.114'],
            'V': ['Реализация товаров и услуг 00КА-00135 от 20.01.2025'],
            'D': [None]
        })
        
        df_source = pd.DataFrame({
            'I': ['Башмак колонный вращающийся БК-Вр.114'],
            'B': ['Январь 2025 г.'],
            'AR': [15000.50],
            'AQ': [14000.00],
            'AS': [14500.00],
            'F': ['СК ТПХ']
        })
        
        # Парсим данные
        df_target_parsed = self.parser.parse_target_sheet(df_target)
        df_source_parsed = self.parser.parse_source_sheet(df_source)
        
        # Выполняем сопоставление
        result_df, stats = self.matcher.match_records(df_target_parsed, df_source_parsed)
        
        # Проверяем результат
        self.assertEqual(result_df.iloc[0]['Рассчитанная себестоимость'], 15000.50)
        self.assertEqual(stats['level1_matches'], 1)
    
    def test_match_by_code_without_date(self):
        """Тест сопоставления по коду без даты приобретения"""
        # Создаем тестовые DataFrame
        df_target = pd.DataFrame({
            'S': ['Муфта резьбовая ОНГ.216.00.000-01-032 (подойди 3/4)'],
            'V': ['Без даты'],
            'D': ['15.04.2025']
        })
        
        df_source = pd.DataFrame({
            'I': ['Муфта резьбовая ОНГ.216.00.000-01-032'],
            'B': ['Апрель 2025 г.'],
            'AR': [8250.75],
            'AQ': [8000.00],
            'AS': [8100.00],
            'F': ['СК ТПХ']
        })
        
        # Парсим данные
        df_target_parsed = self.parser.parse_target_sheet(df_target)
        df_source_parsed = self.parser.parse_source_sheet(df_source)
        
        # Выполняем сопоставление
        result_df, stats = self.matcher.match_records(df_target_parsed, df_source_parsed)
        
        # Проверяем результат
        self.assertEqual(result_df.iloc[0]['Рассчитанная себестоимость'], 8250.75)
        self.assertEqual(stats['level2_matches'], 1)
    
    def test_no_match_manual_review(self):
        """Тест отсутствия совпадения и отправки на ручную проверку"""
        # Создаем тестовые DataFrame
        df_target = pd.DataFrame({
            'S': ['Неизвестная деталь XXX-YYY-ZZZ'],
            'V': ['Дата некорректна'],
            'D': [None]
        })
        
        df_source = pd.DataFrame({
            'I': ['Другая номенклатура'],
            'B': ['Январь 2025 г.'],
            'AR': [1000.00],
            'AQ': [900.00],
            'AS': [950.00],
            'F': ['Другой контрагент']
        })
        
        # Парсим данные
        df_target_parsed = self.parser.parse_target_sheet(df_target)
        df_source_parsed = self.parser.parse_source_sheet(df_source)
        
        # Выполняем сопоставление
        result_df, stats = self.matcher.match_records(df_target_parsed, df_source_parsed)
        
        # Проверяем результат
        self.assertEqual(result_df.iloc[0]['Рассчитанная себестоимость'], '*ТРЕБУЕТ РУЧНОЙ ПРОВЕРКИ*')
        self.assertEqual(stats['manual_checks'], 1)


if __name__ == '__main__':
    unittest.main()