import re
from difflib import SequenceMatcher
from typing import List, Tuple, Optional
import pandas as pd

def normalize_nomenclature(nomenclature: str) -> str:
    """
    Нормализует строку номенклатуры: удаляет лишние пробелы, приводит к нижнему регистру.

    :param nomenclature: Исходная строка номенклатуры
    :return: Нормализованная строка
    """
    if pd.isna(nomenclature):
        return ""

    nomenclature = str(nomenclature)
    # Удалить ведущие/хвостовые пробелы и привести к нижнему регистру
    nomenclature = nomenclature.strip().lower()
    # Заменить множественные пробелы одним пробелом
    nomenclature = re.sub(r'\s+', ' ', nomenclature)

    return nomenclature


def extract_nomenclature_code(nomenclature_text: str) -> Optional[str]:
    """
    Извлекает код номенклатуры из названия.

    :param nomenclature_text: Строка номенклатуры
    :return: Код номенклатуры или None
    """
    if pd.isna(nomenclature_text):
        return None
    if not str(nomenclature_text).strip():
        return None

    patterns = [
        r'[А-ЯЁ]+\.\d+\.\d+\.\d+(?:-\d+)*',
        r'[А-ЯЁ]+\.\d+\.\d+(?:-[\d\.,]+)*',
        r'\d+\.\d+\.\d+(?:-.*)?',  # Ключевое изменение — поддержка сложных кодов
        r'[А-ЯЁ]{2,4}-[А-ЯЁ]{2,4}\.\d+',
        r'\d+\.\d+\.\d+(?:-\d+)*'  # Добавлено для поддержки формата 0.0.0-0-0.0-0-32.23.0-28
    ]

    for pattern in patterns:
        match = re.search(pattern, nomenclature_text)
        if match:
            return match.group(0)

    return None


def fuzzy_match_similarity(s1: str, s2: str) -> float:
    """
    Вычисляет коэффициент схожести двух строк.

    :param s1: Первая строка
    :param s2: Вторая строка
    :return: Коэффициент схожести от 0 до 1
    """
    if pd.isna(s1) or pd.isna(s2):
        return 0.0

    s1, s2 = str(s1).lower(), str(s2).lower()
    return SequenceMatcher(None, s1, s2).ratio()


def fuzzy_match_candidates(nomenclature: str, candidates_df: pd.DataFrame,
                          nomenclature_col: str, threshold: float = 0.75) -> List[Tuple[int, float]]:
    """
    Находит потенциальные совпадения для номенклатуры в справочнике.

    :param nomenclature: Номенклатура для поиска
    :param candidates_df: DataFrame с кандидатами
    :param nomenclature_col: Название столбца с номенклатурой
    :param threshold: Порог схожести (от 0 до 1)
    :return: Список кортежей (индекс, коэффициент схожести)
    """
    if pd.isna(nomenclature):
        return []

    nomenclature = normalize_nomenclature(str(nomenclature))
    matches = []

    for idx, row in candidates_df.iterrows():
        candidate = normalize_nomenclature(str(row[nomenclature_col]))
        similarity = fuzzy_match_similarity(nomenclature, candidate)

        if similarity >= threshold:
            matches.append((idx, similarity))

    # Сортируем по коэффициенту схожести в порядке убывания
    matches.sort(key=lambda x: x[1], reverse=True)

    return matches


def partial_keyword_match(nomenclature: str, candidates_df: pd.DataFrame,
                         nomenclature_col: str) -> List[int]:
    """
    Находит частичное совпадение по ключевым словам.

    :param nomenclature: Номенклатура для поиска
    :param candidates_df: DataFrame с кандидатами
    :param nomenclature_col: Название столбца с номенклатурой
    :return: Список индексов совпадений
    """
    if pd.isna(nomenclature):
        return []

    nomenclature = normalize_nomenclature(str(nomenclature))
    # Выделяем ключевые слова (без кодов)
    keywords = [word for word in nomenclature.split() if not re.search(r'[\d\.-]', word) and len(word) > 2]

    if not keywords:
        return []

    matches = []
    for idx, row in candidates_df.iterrows():
        candidate = normalize_nomenclature(str(row[nomenclature_col]))
        candidate_words = set(candidate.split())

        # Проверяем, содержатся ли все ключевые слова в кандидате
        if all(keyword in candidate_words for keyword in keywords):
            matches.append(idx)

    return matches


def match_by_code(nomenclature: str, candidates_df: pd.DataFrame,
                 nomenclature_col: str) -> Optional[int]:
    """
    Сопоставляет по коду номенклатуры.

    :param nomenclature: Номенклатура для поиска
    :param candidates_df: DataFrame с кандидатами
    :param nomenclature_col: Название столбца с номенклатурой
    :return: Индекс совпадения или None
    """
    code = extract_nomenclature_code(nomenclature)
    if not code:
        return None

    for idx, row in candidates_df.iterrows():
        candidate_code = extract_nomenclature_code(str(row[nomenclature_col]))
        if candidate_code and code in candidate_code:
            return idx

    return None