import pandas as pd
from typing import Tuple, Optional, List
from utils.date_extractor import extract_date_from_document, determine_quarter, parse_period_to_date_range, extract_date_from_realization
from utils.fuzzy_matcher import normalize_nomenclature, extract_nomenclature_code, fuzzy_match_candidates, partial_keyword_match, match_by_code
from utils.logger_config import setup_logger, log_processing_progress
import time

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
        
        # Создаем временные столбцы для сопоставления
        df['nom_key'] = df['Номенклатура_норм']
        df['acquisition_quarter'] = df['Дата приобретения'].apply(determine_quarter)
        df['realization_quarter'] = df['Дата реализации'].apply(determine_quarter)
        df['acquisition_date'] = df['Дата приобретения']
        df['realization_date'] = df['Дата реализации']
        df['nom_code'] = df['Код номенклатуры']
        
        # Обработка по частям для больших файлов
        chunk_size = 1000
        total_rows = len(df)
        self.logger.info(f"Начало сопоставления записей для {total_rows} строк (обработка по {chunk_size} строк)")
        
        for start_idx in range(0, total_rows, chunk_size):
            end_idx = min(start_idx + chunk_size, total_rows)
            df_chunk = df.iloc[start_idx:end_idx].copy()
            
            # Уровень 1: Сопоставление с датой приобретения
            mask_level1 = df_chunk['acquisition_date'].notna()
            if mask_level1.any():
                df_level1 = df_chunk[mask_level1].copy()
                self.logger.info(f"Начало сопоставления уровня 1: {len(df_level1)} записей с датой приобретения (часть {start_idx//chunk_size + 1})")
                matched_level1 = self._level1_match_vectorized(df_level1, df_source)
                
                # Обновляем результаты для уровня 1
                df_chunk.loc[mask_level1, 'Рассчитанная себестоимость'] = matched_level1['Рассчитанная себестоимость'].values
                stats['level1_matches'] += int((matched_level1['Рассчитанная себестоимость'] != "").sum())
                self.logger.info(f"Сопоставление уровня 1 завершено: {int((matched_level1['Рассчитанная себестоимость'] != '').sum())} совпадений")
            
            # Уровень 2: Сопоставление без даты приобретения
            mask_level2 = (df_chunk['acquisition_date'].isna()) & (df_chunk['Рассчитанная себестоимость'] == "")
            if mask_level2.any():
                df_level2 = df_chunk[mask_level2].copy()
                self.logger.info(f"Начало сопоставления уровня 2: {len(df_level2)} записей без даты приобретения (часть {start_idx//chunk_size + 1})")
                matched_level2 = self._level2_match_vectorized(df_level2, df_source)
                
                # Обновляем результаты для уровня 2
                df_chunk.loc[mask_level2, 'Рассчитанная себестоимость'] = matched_level2['Рассчитанная себестоимость'].values
                stats['level2_matches'] += int((matched_level2['Рассчитанная себестоимость'] != "").sum())
                self.logger.info(f"Сопоставление уровня 2 завершено: {int((matched_level2['Рассчитанная себестоимость'] != '').sum())} совпадений")
            
            # Размытое сопоставление только для записей без совпадений
            mask_fuzzy = (df_chunk['Рассчитанная себестоимость'] == "") & (df_chunk['Номенклатура_норм'].notna()) & (df_chunk['Номенклатура_норм'] != "")
            if mask_fuzzy.any():
                df_fuzzy = df_chunk[mask_fuzzy].copy()
                self.logger.info(f"Начало размытого сопоставления: {len(df_fuzzy)} записей (часть {start_idx//chunk_size + 1})")
                matched_fuzzy = self._fuzzy_match_vectorized(df_fuzzy, df_source)
                
                # Обновляем результаты для размытого сопоставления
                df_chunk.loc[mask_fuzzy, 'Рассчитанная себестоимость'] = matched_fuzzy['Рассчитанная себестоимость'].values
                stats['fuzzy_matches'] += int((matched_fuzzy['Рассчитанная себестоимость'] != "").sum())
                self.logger.info(f"Размытое сопоставление завершено: {int((matched_fuzzy['Рассчитанная себестоимость'] != '').sum())} совпадений")
            
            # Отмечаем записи, требующие ручной проверки
            mask_manual = df_chunk['Рассчитанная себестоимость'] == ""
            df_chunk.loc[mask_manual, 'Рассчитанная себестоимость'] = '*ТРЕБУЕТ РУЧНОЙ ПРОВЕРКИ*'
            stats['manual_checks'] += int(mask_manual.sum())
            self.logger.info(f"Записей, требующих ручной проверки в части {start_idx//chunk_size + 1}: {int(mask_manual.sum())}")
            
            # Обновляем основной DataFrame
            df.iloc[start_idx:end_idx] = df_chunk
            
            # Логируем прогресс
            processed = min(end_idx, total_rows)
            self.logger.info(f"Обработано {processed}/{total_rows} строк ({processed/total_rows*100:.1f}%)")
        
        # Удаляем временные столбцы
        df = df.drop(columns=['nom_key', 'acquisition_quarter', 'realization_quarter', 'acquisition_date', 'realization_date', 'nom_code'])
        
        return df, stats
    
    def _level1_match_vectorized(self, df_target: pd.DataFrame, df_source: pd.DataFrame) -> pd.DataFrame:
        """
        Векторизованное сопоставление Уровня 1: По номенклатуре и периоду (с датой приобретения).
        
        :param df_target: DataFrame целевой вкладки
        :param df_source: DataFrame справочной вкладки
        :return: DataFrame с обновленным столбцом "Рассчитанная себестоимость"
        """
        # Создаем копию целевой таблицы
        df = df_target.copy()
        df['Рассчитанная себестоимость'] = ""
        
        # Проверяем наличие необходимых столбцов в df_source
        required_cols = ['Номенклатура_норм', 'Период_начало', 'Период_конец', 'Прямая СС на ед', 'Прямая материальная составляющая']
        available_cols = [col for col in required_cols if col in df_source.columns]
        
        # Проверяем, что все необходимые столбцы присутствуют
        if not all(col in df_source.columns for col in ['Номенклатура_норм', 'Период_начало', 'Период_конец']):
            self.logger.warning("Не найдены необходимые столбцы для сопоставления уровня 1")
            return df
        
        # Подготавливаем данные для сопоставления
        df_clean = df[df['Дата приобретения'].notna()].copy()
        if df_clean.empty:
            self.logger.info("Нет записей с датой приобретения для сопоставления уровня 1")
            return df
        
        # Используем merge для векторизованного сопоставления
        merged_df = pd.merge(
            df_clean[['Номенклатура_норм', 'Дата приобретения']].reset_index(),
            df_source[available_cols],
            on='Номенклатура_норм',
            how='left'
        )
        
        # Фильтруем по дате приобретения в период
        mask = (
            (merged_df['Дата приобретения'].notna()) &
            (merged_df['Период_начало'].notna()) &
            (merged_df['Период_конец'].notna()) &
            (merged_df['Дата приобретения'] >= merged_df['Период_начало']) &
            (merged_df['Дата приобретения'] <= merged_df['Период_конец'])
        )
        filtered_merged = merged_df[mask]
        
        # Выбираем лучшие совпадения (приоритезация)
        filtered_merged = filtered_merged.sort_values(['index', 'Дата приобретения'])
        best_matches = filtered_merged.groupby('index').first()
        
        # Извлекаем стоимость
        if 'Прямая СС на ед' in df_source.columns and 'Прямая материальная составляющая' in df_source.columns:
            best_matches['cost'] = best_matches['Прямая СС на ед'].fillna(best_matches['Прямая материальная составляющая'])
        elif 'Прямая СС на ед' in df_source.columns:
            best_matches['cost'] = best_matches['Прямая СС на ед']
        elif 'Прямая материальная составляющая' in df_source.columns:
            best_matches['cost'] = best_matches['Прямая материальная составляющая']
        else:
            # Если ни один из столбцов с себестоимостью не найден, создаем пустой столбец cost
            best_matches['cost'] = pd.Series([None] * len(best_matches), index=best_matches.index)
        
        # Обновляем целевую таблицу
        for idx in best_matches.index:
            cost = best_matches.loc[idx, 'cost']
            if pd.notna(cost):
                df.loc[df_clean.index[idx], 'Рассчитанная себестоимость'] = cost
        
        # Логируем количество найденных совпадений
        matches_count = int((df['Рассчитанная себестоимость'] != "").sum())
        self.logger.info(f"Уровень 1: найдено {matches_count} совпадений")
        
        return df
    
    def _level2_match_vectorized(self, df_target: pd.DataFrame, df_source: pd.DataFrame) -> pd.DataFrame:
        """
        Векторизованное сопоставление Уровня 2: По коду номенклатуры и контрагенту (без даты приобретения).
        
        :param df_target: DataFrame целевой вкладки
        :param df_source: DataFrame справочной вкладки
        :return: DataFrame с обновленным столбцом "Рассчитанная себестоимость"
        """
        # Создаем копию целевой таблицы
        df = df_target.copy()
        df['Рассчитанная себестоимость'] = ""
        
        # Проверяем наличие столбца контрагента
        if 'Контрагент_норм' not in df_source.columns:
            self.logger.warning("Столбец 'Контрагент_норм' не найден в справочной таблице")
            # Используем все записи из справочной таблицы
            df_source_filtered = df_source.copy()
        else:
            # Фильтруем справочную таблицу по контрагенту
            df_source_filtered = df_source[
                (df_source['Контрагент_норм'].str.contains('ск тпх|ск татпром-холдинг', na=False, case=False))
            ].copy()
            
            if df_source_filtered.empty:
                self.logger.warning("Не найдено записей с контрагентами СК ТПХ или СК Татпром-Холдинг, используем все записи")
                df_source_filtered = df_source.copy()
        
        # Сначала ищем по коду номенклатуры
        mask_with_code = df['Код номенклатуры'].notna()
        if mask_with_code.any():
            df_with_code = df[mask_with_code].copy()
            if not df_with_code.empty:
                # Проверяем наличие необходимых столбцов в df_source_filtered
                required_cols = ['Код номенклатуры', 'Период_начало', 'Период_конец', 'Прямая СС на ед', 'Прямая материальная составляющая', 'Номенклатура_норм']
                available_cols = [col for col in required_cols if col in df_source_filtered.columns]
                
                # Проверяем, что все необходимые столбцы присутствуют
                if all(col in df_source_filtered.columns for col in ['Код номенклатуры', 'Период_начало', 'Период_конец']):
                    # Создаем вспомогательный DataFrame для сопоставления по коду
                    merged_by_code = pd.merge(
                        df_with_code[['Код номенклатуры', 'Дата реализации']].reset_index(),
                        df_source_filtered[available_cols],
                        on='Код номенклатуры',
                        how='left',
                        suffixes=('', '_source')
                    )
                    
                    # Фильтруем по дате реализации в период
                    mask_date = (
                        (merged_by_code['Дата реализации'].notna()) &
                        (merged_by_code['Период_начало'].notna()) &
                        (merged_by_code['Период_конец'].notna()) &
                        (merged_by_code['Дата реализации'] >= merged_by_code['Период_начало']) &
                        (merged_by_code['Дата реализации'] <= merged_by_code['Период_конец'])
                    )
                    filtered_by_code = merged_by_code[mask_date]
                    
                    # Для записей без даты реализации
                    mask_no_date = (
                        (merged_by_code['Дата реализации'].isna()) &
                        (merged_by_code['Код номенклатуры'].notna())
                    )
                    no_date_matches = merged_by_code[mask_no_date]
                    
                    # Объединяем результаты
                    all_code_matches = pd.concat([filtered_by_code, no_date_matches], ignore_index=True)
                    all_code_matches = all_code_matches.drop_duplicates(subset=['index'], keep='first')
                    
                    # Извлекаем стоимость
                    if 'Прямая СС на ед' in df_source_filtered.columns and 'Прямая материальная составляющая' in df_source_filtered.columns:
                        all_code_matches['cost'] = all_code_matches['Прямая СС на ед'].fillna(all_code_matches['Прямая материальная составляющая'])
                    elif 'Прямая СС на ед' in df_source_filtered.columns:
                        all_code_matches['cost'] = all_code_matches['Прямая СС на ед']
                    elif 'Прямая материальная составляющая' in df_source_filtered.columns:
                        all_code_matches['cost'] = all_code_matches['Прямая материальная составляющая']
                    else:
                        # Если ни один из столбцов с себестоимостью не найден, создаем пустой столбец cost
                        all_code_matches['cost'] = pd.Series([None] * len(all_code_matches), index=all_code_matches.index)
                    
                    # Обновляем целевую таблицу
                    for idx in all_code_matches['index']:
                        row_data = all_code_matches[all_code_matches['index'] == idx].iloc[0]
                        cost = row_data['cost']
                        if pd.notna(cost):
                            df.loc[idx, 'Рассчитанная себестоимость'] = cost
        
        # Затем ищем по номенклатуре для записей без стоимости
        mask_no_cost = df['Рассчитанная себестоимость'] == ""
        if mask_no_cost.any():
            df_no_cost = df[mask_no_cost].copy()
            if not df_no_cost.empty:
                # Проверяем наличие необходимых столбцов в df_source_filtered
                required_cols = ['Номенклатура_норм', 'Период_начало', 'Период_конец', 'Прямая СС на ед', 'Прямая материальная составляющая']
                available_cols = [col for col in required_cols if col in df_source_filtered.columns]
                
                # Проверяем, что все необходимые столбцы присутствуют
                if all(col in df_source_filtered.columns for col in ['Номенклатура_норм', 'Период_начало', 'Период_конец']):
                    # Сопоставляем по номенклатуре
                    merged_df = pd.merge(
                        df_no_cost[['Номенклатура_норм', 'Дата реализации']].reset_index(),
                        df_source_filtered[available_cols],
                        on='Номенклатура_норм',
                        how='left'
                    )
                    
                    # Фильтруем по дате реализации в период
                    mask_date = (
                        (merged_df['Дата реализации'].notna()) &
                        (merged_df['Период_начало'].notna()) &
                        (merged_df['Период_конец'].notna()) &
                        (merged_df['Дата реализации'] >= merged_df['Период_начало']) &
                        (merged_df['Дата реализации'] <= merged_df['Период_конец'])
                    )
                    filtered_merged = merged_df[mask_date]
                    
                    # Также добавляем совпадения без даты
                    if 'Прямая СС на ед' in df_source_filtered.columns and 'Прямая материальная составляющая' in df_source_filtered.columns:
                        mask_no_date = (
                            (merged_df['Дата реализации'].isna()) &
                            (merged_df['Прямая СС на ед'].notna() | merged_df['Прямая материальная составляющая'].notna())
                        )
                    elif 'Прямая СС на ед' in df_source_filtered.columns:
                        mask_no_date = (
                            (merged_df['Дата реализации'].isna()) &
                            (merged_df['Прямая СС на ед'].notna())
                        )
                    elif 'Прямая материальная составляющая' in df_source_filtered.columns:
                        mask_no_date = (
                            (merged_df['Дата реализации'].isna()) &
                            (merged_df['Прямая материальная составляющая'].notna())
                        )
                    else:
                        mask_no_date = pd.Series([False] * len(merged_df))
                    
                    no_date_matches = merged_df[mask_no_date]
                    
                    # Объединяем оба результата
                    all_matches = pd.concat([filtered_merged, no_date_matches], ignore_index=True)
                    all_matches = all_matches.drop_duplicates(subset=['index'], keep='first')
                    
                    # Извлекаем стоимость
                    if 'Прямая СС на ед' in df_source_filtered.columns and 'Прямая материальная составляющая' in df_source_filtered.columns:
                        all_matches['cost'] = all_matches['Прямая СС на ед'].fillna(all_matches['Прямая материальная составляющая'])
                    elif 'Прямая СС на ед' in df_source_filtered.columns:
                        all_matches['cost'] = all_matches['Прямая СС на ед']
                    elif 'Прямая материальная составляющая' in df_source_filtered.columns:
                        all_matches['cost'] = all_matches['Прямая материальная составляющая']
                    
                    # Обновляем целевую таблицу
                    for idx in all_matches['index']:
                        row_data = all_matches[all_matches['index'] == idx].iloc[0]
                        cost = row_data['cost']
                        if pd.notna(cost):
                            df.loc[idx, 'Рассчитанная себестоимость'] = cost
        
        # Логируем количество найденных совпадений
        matches_count = int((df['Рассчитанная себестоимость'] != "").sum())
        self.logger.info(f"Уровень 2: найдено {matches_count} совпадений")
        
        return df
    
    def _fuzzy_match_vectorized(self, df_target: pd.DataFrame, df_source: pd.DataFrame) -> pd.DataFrame:
        """
        Векторизованное размытое сопоставление по номенклатуре.
        Применяется только к записям, для которых не найдено точных совпадений.
        
        :param df_target: DataFrame целевой вкладки
        :param df_source: DataFrame справочной вкладки
        :return: DataFrame с обновленным столбцом "Рассчитанная себестоимость"
        """
        # Создаем копию целевой таблицы
        df = df_target.copy()
        df['Рассчитанная себестоимость'] = ""
        
        # Применяем fuzzy matching только к записям с непустой номенклатурой и без стоимости
        mask_to_match = (df['Номенклатура_норм'].notna()) & (df['Номенклатура_норм'] != "") & (df['Рассчитанная себестоимость'] == "")
        df_to_process = df[mask_to_match]
        
        if df_to_process.empty:
            return df
        
        # Ограничиваем количество записей для fuzzy matching для производительности
        if len(df_to_process) > 100:
            self.logger.info(f"Ограничение размытого поиска до 10 записей из {len(df_to_process)}")
            df_to_process = df_to_process.head(10)
        
        # Для каждой записи в целевой таблице ищем наиболее похожую в справочной
        processed_count = 0
        total_to_process = len(df_to_process)
        for idx, row in df_to_process.iterrows():
            nom_value = row['Номенклатура_норм']
            if pd.isna(nom_value) or nom_value == "":
                continue
            
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
                    if cost is not None:
                        df.loc[idx, 'Рассчитанная себестоимость'] = cost
            
            processed_count += 1
            if processed_count % 25 == 0 or processed_count == total_to_process:
                self.logger.info(f"Размытое сопоставление: {processed_count}/{total_to_process} ({processed_count/total_to_process*100:.1f}%)")
        
        # Логируем количество найденных совпадений
        matches_count = 0
        for idx, row in df_to_process.iterrows():
            if row['Рассчитанная себестоимость'] != "":
                matches_count += 1
        
        self.logger.info(f"Размытое сопоставление: найдено {matches_count} совпадений из {processed_count} обработанных записей")
        
        return df
    
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