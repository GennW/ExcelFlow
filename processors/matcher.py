import pandas as pd
from typing import Tuple, Optional, List
from utils.date_extractor import extract_date_from_document, determine_quarter, parse_period_to_date_range, extract_date_from_realization
from utils.fuzzy_matcher import normalize_nomenclature, extract_nomenclature_code, fuzzy_match_candidates, partial_keyword_match, match_by_code
from utils.logger_config import setup_logger

class DataMatcher:
    """
    Класс для сопоставления записей между вкладками
    """
    
    def __init__(self, logger=None):
        self.logger = logger or setup_logger()
    
    def match_records(self, df_target: pd.DataFrame, df_source: pd.DataFrame) -> pd.DataFrame:
        """
        Сопоставляет записи между вкладками и добавляет результаты в целевую вкладку.
        
        :param df_target: DataFrame вкладки "СК ТПХ_1 пг" (обработанной)
        :param df_source: DataFrame вкладки "ВП 2024-2025 НЧТЗ" (обработанной)
        :return: Обновленный DataFrame с добавленным столбцом "Рассчитанная себестоимость"
        """
        # Создаем копию целевой таблицы
        df = df_target.copy()
        
        # Добавляем столбец для рассчитанной себестоимости
        df['Рассчитанная себестоимость'] = ""
        
        # Статистика для отчета
        stats = {
            'exact_matches': 0,
            'level1_matches': 0,
            'level2_matches': 0,
            'fuzzy_matches': 0,
            'manual_checks': 0,
            'missing_data': 0
        }
        
        # Индексы столбцов для удобства
        nom_col_idx = 18  # Столбец S - номенклатура
        doc_col_idx = 21  # Столбец V - документ приобретения
        real_col_idx = 3  # Столбец D - дата реализации
        
        # Проходим по каждой строке целевой вкладки
        for idx, row in df.iterrows():
            # Проверяем наличие извлеченной даты приобретения
            acquisition_date = row['Дата приобретения']
            
            if pd.notna(acquisition_date):
                # Уровень 1: Сопоставление с датой
                result = self._level1_match(row, df_source)
                if result is not None:
                    df.at[idx, 'Рассчитанная себестоимость'] = result
                    stats['level1_matches'] += 1
                else:
                    # Если уровень 1 не дал результата, пробуем уровень 2
                    result = self._level2_match(row, df_source)
                    if result is not None:
                        df.at[idx, 'Рассчитанная себестоимость'] = result
                        stats['level2_matches'] += 1
                    else:
                        df.at[idx, 'Рассчитанная себестоимость'] = '*ТРЕБУЕТ РУЧНОЙ ПРОВЕРКИ*'
                        stats['manual_checks'] += 1
            else:
                # Уровень 2: Сопоставление без даты
                result = self._level2_match(row, df_source)
                if result is not None:
                    df.at[idx, 'Рассчитанная себестоимость'] = result
                    stats['level2_matches'] += 1
                else:
                    # Если уровень 2 не дал результата, пробуем размытое сравнение
                    result = self._fuzzy_match(row, df_source)
                    if result is not None:
                        df.at[idx, 'Рассчитанная себестоимость'] = result
                        stats['fuzzy_matches'] += 1
                    else:
                        # Проверяем, есть ли ключевые данные
                        nom_value = row.iloc[nom_col_idx] if nom_col_idx < len(row) else None
                        doc_value = row.iloc[doc_col_idx] if doc_col_idx < len(row) else None
                        real_value = row.iloc[real_col_idx] if real_col_idx < len(row) else None
                        
                        if pd.isna(nom_value) and pd.isna(extract_date_from_document(str(doc_value))) and pd.isna(extract_date_from_realization(str(real_value))):
                            df.at[idx, 'Рассчитанная себестоимость'] = '*ОТСУТСТВУЮТ КЛЮЧЕВЫЕ ДАТЫ*'
                            stats['missing_data'] += 1
                        else:
                            df.at[idx, 'Рассчитанная себестоимость'] = '*ТРЕБУЕТ РУЧНОЙ ПРОВЕРКИ*'
                            stats['manual_checks'] += 1
        
        return df, stats
    
    def _level1_match(self, row, df_source) -> Optional[float]:
        """
        Сопоставление Уровня 1: По номенклатуре и периоду (с датой приобретения).
        
        :param row: Строка из целевой вкладки
        :param df_source: DataFrame справочной вкладки
        :return: Значение себестоимости или None
        """
        # Извлекаем данные из строки
        nom_value = row['Номенклатура_норм']
        acquisition_date = row['Дата приобретения']
        
        if pd.isna(nom_value) or pd.isna(acquisition_date):
            return None
        
        # Фильтруем строки в справочной вкладке по периоду
        matching_rows = []
        for src_idx, src_row in df_source.iterrows():
            period_start = src_row['Период_начало']
            period_end = src_row['Период_конец']
            
            # Проверяем, попадает ли дата приобретения в период
            if pd.notna(period_start) and pd.notna(period_end) and period_start <= acquisition_date <= period_end:
                # Проверяем совпадение по номенклатуре
                if src_row['Номенклатура_норм'] == nom_value:
                    matching_rows.append((src_idx, src_row))
        
        # Если найдено несколько совпадений, применяем приоритизацию
        if len(matching_rows) > 1:
            matching_rows = self._prioritize_matches(matching_rows, acquisition_date)
        
        # Возвращаем значение себестоимости из первого совпадения
        if matching_rows:
            _, src_row = matching_rows[0]
            cost = src_row['Прямая СС на ед'] if 'Прямая СС на ед' in src_row and pd.notna(src_row['Прямая СС на ед']) else None
            if cost is None:
                cost = src_row['Прямая материальная составляющая'] if 'Прямая материальная составляющая' in src_row and pd.notna(src_row['Прямая материальная составляющая']) else None
            return cost
        
        return None
    
    def _level2_match(self, row, df_source) -> Optional[float]:
        """
        Сопоставление Уровня 2: По коду номенклатуры и контрагенту (без даты приобретения).
        
        :param row: Строка из целевой вкладки
        :param df_source: DataFrame справочной вкладки
        :return: Значение себестоимости или None
        """
        # Извлекаем данные из строки
        nom_code = row['Код номенклатуры']
        nom_value = row['Номенклатура_норм']
        realization_date = row['Дата реализации']
        
        matching_rows = []
        
        if nom_code:
            # Поиск по коду номенклатуры
            for src_idx, src_row in df_source.iterrows():
                src_code = src_row['Код номенклатуры']
                if src_code and nom_code in src_code:
                    # Проверяем контрагента на наличие ключевых слов
                    counterparty = src_row['Контрагент_норм'] if 'Контрагент_норм' in src_row else ""
                    if 'ск тпх' in counterparty or 'ск татпром-холдинг' in counterparty:
                        # Если есть дата реализации, проверяем период
                        if pd.notna(realization_date):
                            period_start = src_row['Период_начало']
                            period_end = src_row['Период_конец']
                            if pd.notna(period_start) and pd.notna(period_end) and period_start <= realization_date <= period_end:
                                matching_rows.append((src_idx, src_row))
                        else:
                            matching_rows.append((src_idx, src_row))
        
        # Если не нашли по коду, пробуем по номенклатуре
        if not matching_rows and nom_value:
            for src_idx, src_row in df_source.iterrows():
                # Проверяем контрагента на наличие ключевых слов
                counterparty = src_row['Контрагент_норм'] if 'Контрагент_норм' in src_row else ""
                if 'ск тпх' in counterparty or 'ск татпром-холдинг' in counterparty:
                    # Если есть дата реализации, проверяем период
                    if pd.notna(realization_date):
                        period_start = src_row['Период_начало']
                        period_end = src_row['Период_конец']
                        if pd.notna(period_start) and pd.notna(period_end) and period_start <= realization_date <= period_end:
                            if src_row['Номенклатура_норм'] == nom_value:
                                matching_rows.append((src_idx, src_row))
                    else:
                        if src_row['Номенклатура_норм'] == nom_value:
                            matching_rows.append((src_idx, src_row))
        
        # Если найдено несколько совпадений, применяем приоритизацию
        if len(matching_rows) > 1:
            matching_rows = self._prioritize_matches(matching_rows, realization_date if pd.notna(realization_date) else None)
        
        # Возвращаем значение себестоимости из первого совпадения
        if matching_rows:
            _, src_row = matching_rows[0]
            cost = src_row['Прямая СС на ед'] if 'Прямая СС на ед' in src_row and pd.notna(src_row['Прямая СС на ед']) else None
            if cost is None:
                cost = src_row['Прямая материальная составляющая'] if 'Прямая материальная составляющая' in src_row and pd.notna(src_row['Прямая материальная составляющая']) else None
            return cost
        
        return None
    
    def _fuzzy_match(self, row, df_source) -> Optional[float]:
        """
        Размытое сопоставление по номенклатуре.
        
        :param row: Строка из целевой вкладки
        :param df_source: DataFrame справочной вкладки
        :return: Значение себестоимости или None
        """
        nom_value = row['Номенклатура_норм']
        if pd.isna(nom_value):
            return None
        
        # Используем fuzzy_match_candidates для поиска похожих номенклатур
        matches = fuzzy_match_candidates(nom_value, df_source, 'Номенклатура_норм', threshold=0.6)
        
        if matches:
            # Берем лучшее совпадение
            best_match_idx, similarity = matches[0]
            src_row = df_source.iloc[best_match_idx]
            
            # Проверяем, если коэффициент схожести выше 75%, считаем это успешным совпадением
            if similarity >= 0.75:
                cost = src_row['Прямая СС на ед'] if 'Прямая СС на ед' in src_row and pd.notna(src_row['Прямая СС на ед']) else None
                if cost is None:
                    cost = src_row['Прямая материальная составляющая'] if 'Прямая материальная составляющая' in src_row and pd.notna(src_row['Прямая материальная составляющая']) else None
                return cost
        
        return None
    
    def _prioritize_matches(self, matches: List[Tuple[int, pd.Series]], target_date: Optional[pd.Timestamp] = None) -> List[Tuple[int, pd.Series]]:
        """
        Приоритизирует найденные совпадения по заданным критериям.
        
        :param matches: Список совпадений в формате (индекс, строка)
        :param target_date: Целевая дата для сравнения (опционально)
        :return: Отсортированный список совпадений
        """
        def get_priority(match_tuple):
            idx, row = match_tuple
            
            # Приоритеты:
            # 1. Совпадение по всем трём параметрам (номенклатура + контрагент + период) - не применимо здесь
            # 2. Совпадение по номенклатуре + периоду - для уровня 1
            # 3. Совпадение по коду номенклатуры + периоду - для уровня 2
            # 4. Наиболее близкая дата
            # 5. Наименьшая стоимость закупки
            
            priority = 0
            
            # Если есть целевая дата, проверяем близость
            if target_date and pd.notna(row.get('Период_начало')) and pd.notna(row.get('Период_конец')):
                period_start = row['Период_начало']
                period_end = row['Период_конец']
                
                # Если дата входит в период, это приоритет
                if period_start <= target_date <= period_end:
                    priority += 1000
                
                # Штраф за удалённость даты
                mid_period = period_start + (period_end - period_start) / 2
                date_diff = abs((target_date - mid_period).days)
                priority -= date_diff  # Чем больше разница, тем ниже приоритет
            
            # Приоритет по стоимости (меньше - лучше)
            cost = row.get('Стоимость закупки НЧТ', float('inf'))
            if pd.notna(cost):
                # Инвертируем, чтобы меньшая стоимость имела больший приоритет
                priority -= cost / 100  # Нормализуем для баланса с другими приоритетами
            
            return priority
        
        # Сортируем по приоритету в порядке убывания
        return sorted(matches, key=get_priority, reverse=True)