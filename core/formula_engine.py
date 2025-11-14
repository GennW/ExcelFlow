"""Движок формул СУММЕСЛИМН"""
import pandas as pd
from typing import Optional, Dict
import logging

logger = logging.getLogger(__name__)

class FormulaEngine:
    def __init__(self, df_source: pd.DataFrame, column_indices: Dict[str, int]):
        self.df_source = df_source
        self.columns = column_indices
        self._map_columns()

    def _map_columns(self) -> None:
        self.col_quantity = self.df_source.iloc[:, self.columns['QUANTITY']]
        self.col_cost_q = self.df_source.iloc[:, self.columns['COST_Q']]
        self.col_cost_r = self.df_source.iloc[:, self.columns['COST_R']]
        self.col_cost_x = self.df_source.iloc[:, self.columns['COST_X']]
        self.col_nomenclature = self.df_source.iloc[:, self.columns['NOMENCLATURE']]
        self.col_period = self.df_source.iloc[:, self.columns['PERIOD_QUARTER']]
        logger.info("Маппинг столбцов выполнен")

    def sumifs_weighted_avg(self, sum_column_name: str, nomenclature: str, quarter: str) -> Optional[float]:
        mask = (self.col_nomenclature == nomenclature) & (self.col_period == quarter)
        if mask.sum() == 0:
            return None
        if sum_column_name == 'COST_R':
            sum_column = self.col_cost_r
        elif sum_column_name == 'COST_Q':
            sum_column = self.col_cost_q
        elif sum_column_name == 'COST_X':
            sum_column = self.col_cost_x
        else:
            return None
        total_sum = sum_column[mask].sum()
        total_qty = self.col_quantity[mask].sum()
        if total_qty == 0 or pd.isna(total_qty):
            return None
        return round(total_sum / total_qty, 2)

    def calculate_aq(self, nomenclature: str, quarter: str) -> Optional[float]:
        return self.sumifs_weighted_avg('COST_R', nomenclature, quarter)

    def calculate_ar(self, nomenclature: str, quarter: str) -> Optional[float]:
        return self.sumifs_weighted_avg('COST_Q', nomenclature, quarter)

    def calculate_as(self, nomenclature: str, quarter: str) -> Optional[float]:
        return self.sumifs_weighted_avg('COST_X', nomenclature, quarter)
