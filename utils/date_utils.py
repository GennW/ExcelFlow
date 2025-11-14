"""Работа с датами (КРИТИЧЕСКИЕ ИСПРАВЛЕНИЯ)"""
import re
from datetime import datetime
from typing import Optional, List, Tuple
import pandas as pd
import logging
from config import PSTR_START, PSTR_LENGTH

logger = logging.getLogger(__name__)


def extract_date_pstr(text: str) -> Optional[datetime]:
    """
    ИСПРАВЛЕНО: Извлечение даты из фиксированной позиции в строке.

    Дата находится в позициях PSTR_START:PSTR_START+PSTR_LENGTH (44:54).
    Формат: DD.MM.YYYY

    Args:
        text: Текст документа

    Returns:
        Объект datetime или None
    """
    if not text or pd.isna(text):
        return None

    text_str = str(text).strip()

    # Проверка минимальной длины
    if len(text_str) < PSTR_START + PSTR_LENGTH:
        logger.debug(f"Текст слишком короткий: {len(text_str)} < {PSTR_START + PSTR_LENGTH}")
        return None

    try:
        # Извлекаем подстроку с датой
        date_str = text_str[PSTR_START:PSTR_START + PSTR_LENGTH]
        logger.debug(f"Извлечена подстрока: '{date_str}' из позиции {PSTR_START}:{PSTR_START + PSTR_LENGTH}")

        # Парсим дату
        date_obj = datetime.strptime(date_str, '%d.%m.%Y')
        logger.debug(f"✅ Дата успешно извлечена (PSTR): {date_obj.strftime('%d.%m.%Y')}")
        return date_obj

    except (ValueError, IndexError) as e:
        logger.debug(f"Ошибка парсинга PSTR: {e}")
        return None


def extract_date_regex(text: str) -> Optional[datetime]:
    """
    КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Извлечение даты с помощью регулярных выражений.

    Поддерживаемые форматы:
    - от DD.MM.YYYY
    - DD.MM.YYYY г
    - DD.MM.YYYY
    - DD-MM-YYYY
    - DD/MM/YYYY

    Args:
        text: Текст документа

    Returns:
        Объект datetime или None
    """
    if not text or pd.isna(text):
        return None

    text_str = str(text).strip()

    # КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Расширенный список паттернов
    patterns = [
        # Формат: от 01.01.2024
        (r'от\s+(\d{2}\.\d{2}\.\d{4})', '%d.%m.%Y'),

        # Формат: 01.01.2024 г или 01.01.2024г
        (r'(\d{2}\.\d{2}\.\d{4})\s*г', '%d.%m.%Y'),

        # Формат: 01.01.2024 (без дополнительных символов)
        (r'\b(\d{2}\.\d{2}\.\d{4})\b', '%d.%m.%Y'),

        # Формат: 01-01-2024
        (r'\b(\d{2})-(\d{2})-(\d{4})\b', '%d-%m-%Y'),

        # Формат: 01/01/2024
        (r'\b(\d{2})/(\d{2})/(\d{4})\b', '%d/%m/%Y'),

        # Формат: дата в начале строки
        (r'^(\d{2}\.\d{2}\.\d{4})', '%d.%m.%Y'),

        # Формат: дата после "№" или "N"
        (r'[№N]\s*\d+\s+от\s+(\d{2}\.\d{2}\.\d{4})', '%d.%m.%Y'),
    ]

    # УЛУЧШЕНИЕ: Пробуем все паттерны по порядку
    for pattern, date_format in patterns:
        try:
            match = re.search(pattern, text_str, re.IGNORECASE)
            if match:
                date_str = match.group(1)
                logger.debug(f"Найдено совпадение с паттерном '{pattern}': '{date_str}'")

                # Парсим дату
                date_obj = datetime.strptime(date_str, date_format)

                # УЛУЧШЕНИЕ: Проверка валидности года (2000-2030)
                if 2000 <= date_obj.year <= 2030:
                    logger.debug(f"✅ Дата успешно извлечена (REGEX): {date_obj.strftime('%d.%m.%Y')}")
                    return date_obj
                else:
                    logger.debug(f"Год вне допустимого диапазона: {date_obj.year}")

        except (ValueError, IndexError) as e:
            logger.debug(f"Ошибка парсинга паттерна '{pattern}': {e}")
            continue

    logger.debug(f"Не удалось извлечь дату из текста: '{text_str[:100]}'")
    return None


def extract_acquisition_date(document_text: str) -> Optional[datetime]:
    """
    КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Основная функция извлечения даты приобретения.

    Стратегия:
    1. Сначала пробуем PSTR (фиксированная позиция)
    2. Если не получилось, пробуем REGEX (поиск по паттернам)

    Args:
        document_text: Текст документа

    Returns:
        Объект datetime или None
    """
    if not document_text or pd.isna(document_text):
        logger.debug("Пустой текст документа")
        return None

    # Метод 1: Фиксированная позиция (PSTR)
    date = extract_date_pstr(document_text)
    if date:
        return date

    # Метод 2: Регулярные выражения (REGEX)
    date = extract_date_regex(document_text)
    if date:
        return date

    # УЛУЧШЕНИЕ: Дополнительное логирование для отладки
    logger.debug(f"❌ Дата не найдена в документе: '{str(document_text)[:100]}'")

    return None


def determine_quarter(date: datetime) -> str:
    """
    ИСПРАВЛЕНО: Определение квартала по дате.

    Args:
        date: Объект datetime

    Returns:
        Строка в формате "X квартал YYYY" или пустая строка
    """
    if not date or pd.isna(date):
        return ""

    try:
        month = date.month
        year = date.year

        # Определяем квартал (1-4)
        quarter_num = (month - 1) // 3 + 1

        result = f"{quarter_num} квартал {year}"
        logger.debug(f"Квартал определен: {result} для даты {date.strftime('%d.%m.%Y')}")

        return result

    except Exception as e:
        logger.error(f"Ошибка определения квартала: {e}")
        return ""


def parse_date_flexible(date_str: str) -> Optional[datetime]:
    """
    УЛУЧШЕНИЕ: Гибкий парсинг даты с поддержкой разных форматов.

    Args:
        date_str: Строка с датой

    Returns:
        Объект datetime или None
    """
    if not date_str or pd.isna(date_str):
        return None

    # Список возможных форматов
    formats = [
        '%d.%m.%Y',
        '%d-%m-%Y',
        '%d/%m/%Y',
        '%Y-%m-%d',
        '%Y.%m.%d',
        '%d.%m.%y',
        '%d-%m-%y',
    ]

    date_str_clean = str(date_str).strip()

    for fmt in formats:
        try:
            return datetime.strptime(date_str_clean, fmt)
        except ValueError:
            continue

    return None
