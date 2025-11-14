"""Утилиты для работы с датами"""
import re
from datetime import datetime
from typing import Optional
import pandas as pd

def extract_date_pstr(text: str, start: int = 44, length: int = 10) -> Optional[datetime]:
    if not text or pd.isna(text) or len(str(text)) < start + length:
        return None
    try:
        text_str = str(text)
        date_str = text_str[start:start + length]
        return datetime.strptime(date_str, '%d.%m.%Y')
    except (ValueError, IndexError):
        return None

def extract_date_regex(text: str) -> Optional[datetime]:
    if not text or pd.isna(text):
        return None
    patterns = [r'от\s+(\d{2}\.\d{2}\.\d{4})', r'от\s+(\d{2}\.\d{2}\.\d{4})\s+\d{1,2}:\d{2}:\d{2}']
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
    date = extract_date_pstr(document_text)
    if date:
        return date
    return extract_date_regex(document_text)

def determine_quarter(date: datetime) -> str:
    if not date or pd.isna(date):
        return ""
    month = date.month
    year = date.year
    quarter_num = (month - 1) // 3 + 1
    return f"{quarter_num} квартал {year}"
