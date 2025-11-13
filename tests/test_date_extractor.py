import unittest
from datetime import datetime
from utils.date_extractor import extract_date_from_document, determine_quarter, parse_period_to_date_range, extract_date_from_realization


class TestDateExtractor(unittest.TestCase):
    
    def test_extract_date_from_document_standard_format(self):
        """Тест парсинга стандартного формата даты"""
        document = "Реализация товаров и услуг 00КА-000135 от 20.01.2025 23:59:59"
        expected = datetime(2025, 1, 20)
        result = extract_date_from_document(document)
        self.assertEqual(result, expected)
    
    def test_extract_date_from_document_purchase_format(self):
        """Тест извлечения даты из формата 'Приобретение товаров и услуг'"""
        document = "Приобретение товаров и услуг 00КА-001861 от 28.06.2024 0:00:00"
        expected = datetime(2024, 6, 28)
        result = extract_date_from_document(document)
        self.assertEqual(result, expected)
    
    def test_extract_date_from_document_internal_invoice_format(self):
        """Тест извлечения даты из формата 'Внутренняя накладная'"""
        document = "Внутренняя накладная 0КА-000054 от 31.12.2023 23:59:59"
        expected = datetime(2023, 12, 31)
        result = extract_date_from_document(document)
        self.assertEqual(result, expected)
    
    def test_extract_date_from_document_sale_format(self):
        """Тест извлечения даты из формата 'Реализация товаров и услуг'"""
        document = "Реализация товаров и услуг 00КА-000135 от 20.01.2025 23:59:59"
        expected = datetime(2025, 1, 20)
        result = extract_date_from_document(document)
        self.assertEqual(result, expected)
    
    def test_extract_date_from_document_without_time(self):
        """Тест парсинга формата даты без времени"""
        document = "Реализация товаров и услуг 00КА-000135 от 15.02.2025"
        expected = datetime(2025, 2, 15)
        result = extract_date_from_document(document)
        self.assertEqual(result, expected)
    
    def test_extract_date_from_document_invalid_format(self):
        """Тест обработки некорректного формата даты"""
        document = "Некорректные данные"
        result = extract_date_from_document(document)
        self.assertIsNone(result)
    
    def test_extract_date_from_document_none_input(self):
        """Тест обработки None в качестве входных данных"""
        result = extract_date_from_document(None)
        self.assertIsNone(result)
    
    def test_determine_quarter_q1(self):
        """Тест определения первого квартала"""
        date = datetime(2025, 2, 15)
        result = determine_quarter(date)
        self.assertEqual(result, '1 квартал 2025')
    
    def test_determine_quarter_q2(self):
        """Тест определения второго квартала"""
        date = datetime(2025, 5, 15)
        result = determine_quarter(date)
        self.assertEqual(result, '2 квартал 2025')
    
    def test_determine_quarter_q3(self):
        """Тест определения третьего квартала"""
        date = datetime(2025, 8, 15)
        result = determine_quarter(date)
        self.assertEqual(result, '3 квартал 2025')
    
    def test_determine_quarter_q4(self):
        """Тест определения четвертого квартала"""
        date = datetime(2025, 11, 15)
        result = determine_quarter(date)
        self.assertEqual(result, '4 квартал 2025')
    
    def test_determine_quarter_none_input(self):
        """Тест определения квартала для None"""
        result = determine_quarter(None)
        self.assertIsNone(result)
    
    def test_parse_period_to_date_range_january(self):
        """Тест преобразования января в диапазон дат"""
        period = "Январь 2025 г."
        start, end = parse_period_to_date_range(period)
        expected_start = datetime(2025, 1, 1)
        expected_end = datetime(2025, 1, 31)
        self.assertEqual(start, expected_start)
        self.assertEqual(end, expected_end)
    
    def test_parse_period_to_date_range_february_leap_year(self):
        """Тест преобразования февраля високосного года в диапазон дат"""
        period = "Февраль 2024 г."
        start, end = parse_period_to_date_range(period)
        expected_start = datetime(2024, 2, 1)
        expected_end = datetime(2024, 2, 29)
        self.assertEqual(start, expected_start)
        self.assertEqual(end, expected_end)
    
    def test_parse_period_to_date_range_february_non_leap_year(self):
        """Тест преобразования февраля невисокосного года в диапазон дат"""
        period = "Февраль 2025 г."
        start, end = parse_period_to_date_range(period)
        expected_start = datetime(2025, 2, 1)
        expected_end = datetime(2025, 2, 28)
        self.assertEqual(start, expected_start)
        self.assertEqual(end, expected_end)
    
    def test_parse_period_to_date_range_december(self):
        """Тест преобразования декабря в диапазон дат"""
        period = "Декабрь 2024 г."
        start, end = parse_period_to_date_range(period)
        expected_start = datetime(2024, 12, 1)
        expected_end = datetime(2024, 12, 31)
        self.assertEqual(start, expected_start)
        self.assertEqual(end, expected_end)
    
    def test_parse_period_to_date_range_invalid_format(self):
        """Тест обработки некорректного формата периода"""
        period = "Некорректный формат"
        start, end = parse_period_to_date_range(period)
        self.assertIsNone(start)
        self.assertIsNone(end)
    
    def test_extract_date_from_realization_standard_format(self):
        """Тест извлечения даты из строки даты реализации"""
        realization = "01.03.2025"
        expected = datetime(2025, 3, 1)
        result = extract_date_from_realization(realization)
        self.assertEqual(result, expected)
    
    def test_extract_date_from_realization_none_input(self):
        """Тест обработки None в качестве входных данных для даты реализации"""
        result = extract_date_from_realization(None)
        self.assertIsNone(result)


if __name__ == '__main__':
    unittest.main()