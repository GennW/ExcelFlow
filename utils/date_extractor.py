import re
from datetime import datetime, timedelta
from typing import Optional, Union
import pandas as pd

def extract_date_from_document(document_text: str) -> Optional[datetime]:
    """
    Извлекает дату из различных форматов документов.
    
    :param document_text: Строка документа, например: "Реализация товаров и услуг 00КА-000135 от 20.01.2025 23:59:59"
    :return: Объект datetime или None, если дата не найдена
    """
    if not document_text or pd.isna(document_text) or document_text == "":
        return None
    
    patterns = [
        r'Приобретение товаров и услуг.*?от\s+(\d{2}\.\d{2}\.\d{4})',
        r'Внутренняя накладная.*?от\s+(\d{2}\.\d{2}\.\d{4})',
        r'Реализация товаров и услуг.*?от\s+(\d{2}\.\d{2}\.\d{4})',
        r'Приобретение товаров и услуг.*?от\s+(\d{2}\.\d{2}\.\d{4})\s+\d{1,2}:\d{2}:\d{2}',
        r'Внутренняя накладная.*?от\s+(\d{2}\.\d{2}\.\d{4})\s+\d{1,2}:\d{2}:\d{2}',
        r'Реализация товаров и услуг.*?от\s+(\d{2}\.\d{2}\.\d{4})\s+\d{1,2}:\d{2}:\d{2}'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, str(document_text))
        if match:
            date_str = match.group(1)
            try:
                return datetime.strptime(date_str, '%d.%m.%Y')
            except ValueError:
                continue
    
    return None


def determine_quarter(date: Optional[datetime]) -> Optional[str]:
    """
    Определяет квартал на основе даты в формате 'N квартал YYYY'.
    
    :param date: Объект datetime
    :return: Строка в формате 'N квартал YYYY' или None
    """
    if date is None:
        return None
    
    month = date.month
    year = date.year
    quarter_num = (month - 1) // 3 + 1
    
    return f"{quarter_num} квартал {year}"


def parse_period_to_date_range(period_str: str) -> tuple:
    """
    Преобразует строку периода в диапазон дат.
    
    :param period_str: Строка периода, например: "Декабрь 2024 г."
    :return: Кортеж из двух дат (начало, конец) или None
    """
    if not period_str or pd.isna(period_str) or period_str == "":
        return None, None
    
    # Регулярное выражение для извлечения месяца и года
    pattern = r'(\w+)\s+(\d{4})\s*г\.?'
    match = re.search(pattern, str(period_str))
    
    if not match:
        return None, None
    
    month_name, year = match.groups()
    year = int(year)
    
    # Словарь для сопоставления названий месяцев с номерами
    month_map = {
        'Январь': 1, 'Февраль': 2, 'Март': 3, 'Апрель': 4, 'Май': 5, 'Июнь': 6,
        'Июль': 7, 'Август': 8, 'Сентябрь': 9, 'Октябрь': 10, 'Ноябрь': 11, 'Декабрь': 12,
        'Января': 1, 'Февраля': 2, 'Марта': 3, 'Апреля': 4, 'Мая': 5, 'Июня': 6,
        'Июля': 7, 'Августа': 8, 'Сентября': 9, 'Октября': 10, 'Ноября': 11, 'Декабря': 12
    }
    
    if month_name not in month_map:
        return None, None
    
    month = month_map[month_name]
    
    # Создание начальной и конечной даты месяца
    start_date = datetime(year, month, 1)
    
    # Определение последнего дня месяца
    if month == 12:
        end_date = datetime(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = datetime(year, month + 1, 1) - timedelta(days=1)
    
    return start_date, end_date


def extract_date_from_realization(realization_date: str) -> Optional[datetime]:
    """
    Извлекает дату из строки даты реализации.
    
    :param realization_date: Строка даты реализации, например: "01.03.2025"
    :return: Объект datetime или None, если дата не найдена
    """
    if not realization_date or pd.isna(realization_date) or realization_date == "":
        return None
    
    # Попробуем разные форматы даты
    formats = ['%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y']
    
    for fmt in formats:
        try:
            return datetime.strptime(str(realization_date), fmt)
        except ValueError:
            continue
    
    return None