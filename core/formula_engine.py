"""Движок формул СУММЕСЛИМН (ИСПРАВЛЕННАЯ ВЕРСИЯ с улучшениями)"""
import pandas as pd
from typing import Optional, Dict, List
import logging
import re

logger = logging.getLogger(__name__)


class FormulaEngine:
    """
    Движок для выполнения расчетов по формулам СУММЕСЛИМН.
    Реализует взвешенное среднее на основе количества.
    """

    def __init__(self, df_source: pd.DataFrame, column_indices: Dict[str, int]):
        """
        Инициализация движка формул.

        Args:
            df_source: DataFrame с справочными данными
            column_indices: Словарь с индексами столбцов
        """
        self.df_source = df_source
        self.columns = column_indices
        self._map_columns()
        self._prepare_lookup_index()

    def _map_columns(self) -> None:
        """Маппинг столбцов DataFrame для быстрого доступа"""
        self.col_quantity = self.df_source.iloc[:, self.columns['QUANTITY']]
        self.col_cost_q = self.df_source.iloc[:, self.columns['COST_Q']]
        self.col_cost_r = self.df_source.iloc[:, self.columns['COST_R']]
        self.col_cost_x = self.df_source.iloc[:, self.columns['COST_X']]
        self.col_nomenclature = self.df_source.iloc[:, self.columns['NOMENCLATURE']]
        self.col_period = self.df_source.iloc[:, self.columns['PERIOD_QUARTER']]
        logger.info("✅ Маппинг столбцов выполнен")

    def _prepare_lookup_index(self) -> None:
        """
        УЛУЧШЕНИЕ: Подготовка индекса для быстрого поиска.
        Создает словарь для ускорения поиска совпадений.
        """
        self.lookup_index = {}

        for idx in range(len(self.df_source)):
            nomenclature = self.col_nomenclature.iloc[idx]
            period = self.col_period.iloc[idx]

            # Пропускаем пустые значения
            if pd.isna(nomenclature) or pd.isna(period):
                continue

            # Нормализация ключа
            nomenclature_norm = str(nomenclature).strip()
            period_norm = str(period).strip()

            key = (nomenclature_norm, period_norm)

            if key not in self.lookup_index:
                self.lookup_index[key] = []

            self.lookup_index[key].append(idx)

        logger.info(f"✅ Индекс подготовлен: {len(self.lookup_index)} уникальных комбинаций")

    def _normalize_nomenclature(self, nomenclature: str) -> str:
        """
        УЛУЧШЕНИЕ: Нормализация номенклатуры для лучшего сопоставления.

        Args:
            nomenclature: Исходная номенклатура

        Returns:
            Нормализованная номенклатура
        """
        if pd.isna(nomenclature):
            return ""

        result = str(nomenclature).strip()
        # Удаляем лишние пробелы
        result = re.sub(r'\s+', ' ', result)

        return result

    def sumifs_weighted_avg(self, sum_column_name: str, nomenclature: str, quarter: str) -> Optional[float]:
        """
        Вычисляет взвешенное среднее значение по формуле СУММЕСЛИМН.

        Args:
            sum_column_name: Имя столбца для суммирования ('COST_R', 'COST_Q', 'COST_X')
            nomenclature: Номенклатура для фильтрации
            quarter: Квартал для фильтрации

        Returns:
            Взвешенное среднее или None, если совпадений не найдено
        """
        # Нормализация входных данных
        nomenclature_norm = self._normalize_nomenclature(nomenclature)
        quarter_norm = str(quarter).strip() if pd.notna(quarter) else ""

        # Проверка пустых значений
        if not nomenclature_norm or not quarter_norm:
            return None

        # УЛУЧШЕНИЕ: Используем индекс для быстрого поиска
        key = (nomenclature_norm, quarter_norm)
        indices = self.lookup_index.get(key, [])

        if not indices:
            logger.debug(f"Совпадений не найдено для: '{nomenclature_norm}' | '{quarter_norm}'")
            return None

        # Выбираем нужный столбец для суммирования
        if sum_column_name == 'COST_R':
            sum_column = self.col_cost_r
        elif sum_column_name == 'COST_Q':
            sum_column = self.col_cost_q
        elif sum_column_name == 'COST_X':
            sum_column = self.col_cost_x
        else:
            logger.warning(f"Неизвестный столбец: {sum_column_name}")
            return None

        # Суммируем значения и количество
        total_sum = 0
        total_qty = 0

        for idx in indices:
            cost_val = sum_column.iloc[idx]
            qty_val = self.col_quantity.iloc[idx]

            # Пропускаем некорректные значения
            if pd.isna(cost_val) or pd.isna(qty_val):
                continue

            try:
                cost = float(cost_val)
                qty = float(qty_val)

                if qty > 0:  # УЛУЧШЕНИЕ: Только положительные количества
                    total_sum += cost
                    total_qty += qty
            except (ValueError, TypeError):
                continue

        # Проверка на деление на ноль
        if total_qty == 0 or pd.isna(total_qty):
            logger.debug(f"Нулевое количество для: '{nomenclature_norm}' | '{quarter_norm}'")
            return None

        # Возвращаем взвешенное среднее
        result = total_sum / total_qty
        logger.debug(f"✅ Расчет: '{nomenclature_norm}' | '{quarter_norm}' | {sum_column_name}={result:.2f}")

        return round(result, 2)

    def calculate_aq(self, nomenclature: str, quarter: str) -> Optional[float]:
        """Расчет стоимости закупки (столбец AQ)"""
        return self.sumifs_weighted_avg('COST_R', nomenclature, quarter)

    def calculate_ar(self, nomenclature: str, quarter: str) -> Optional[float]:
        """Расчет прямой себестоимости (столбец AR)"""
        return self.sumifs_weighted_avg('COST_Q', nomenclature, quarter)

    def calculate_as(self, nomenclature: str, quarter: str) -> Optional[float]:
        """Расчет накладных расходов (столбец AS)"""
        return self.sumifs_weighted_avg('COST_X', nomenclature, quarter)
