"""Работа с датами"""
import re
from datetime import datetime
from typing import Optional
import pandas as pd
import logging
from config import PSTR_START, PSTR_LENGTH, DATE_FORMAT_OUTPUT

logger = logging.getLogger(__name__)

def extract_date_pstr(text: str) -> Optional[datetime]:
    if not text or pd.isna(text):
        return None
    text_str = str(text)
    if len(text_str) < PSTR_START + PSTR_LENGTH:
        return None
    try:
        date_str = text_str[PSTR_START:PSTR_START + PSTR_LENGTH]
        date_obj = datetime.strptime(date_str, '%d.%m.%Y')
        return date_obj
    except (ValueError, IndexError):
        return None

def extract_date_regex(text: str) -> Optional[datetime]:
    if not text or pd.isna(text):
        return None
    patterns = [r'от\s+(\d{2}\.\d{2}\.\d{4})']
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
    if not document_text or pd.isna(document_text):
        return None
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
