"""Модуль для сравнения (ИСПРАВЛЕННАЯ ВЕРСИЯ - ДИНАМИЧЕСКИЙ ПОИСК СТОЛБЦОВ)"""
import pandas as pd
from typing import Dict, Optional
from utils.logger import get_logger
from config import ERROR_CODES, ERROR_DESCRIPTIONS

logger = get_logger(__name__)


class DataComparison:
    """
    ИСПРАВЛЕНИЕ: Динамический поиск эталонных столбцов по именам заголовков,
    а не по фиксированным индексам.
    """

    EXPECTED_NAMES = {
        'AO': 'Дата приобретения',
        'AP': 'Квартал приобретения',
        'AQ': 'Стоимость закупки НЧТЗ 1 ед',
        'AR': 'Прямая СС НЧТЗ 1 ед',
        'AS': 'НР НЧТЗ 1 ед',
    }

    # АЛЬТЕРНАТИВНЫЕ ИМЕНА (на случай вариаций в названиях)
    ALTERNATIVE_NAMES = {
        'AO': ['дата приобретения', 'дата', 'date'],
        'AP': ['квартал приобретения', 'квартал', 'период', 'quarter'],
        'AQ': ['стоимость закупки', 'закупка', 'стоимость'],
        'AR': ['прямая сс', 'себестоимость', 'прямая себестоимость'],
        'AS': ['нр нчтз', 'накладные расходы', 'нр'],
    }

    def __init__(self, df_source: pd.DataFrame):
        """
        Инициализация с динамическим поиском столбцов.

        Args:
            df_source: Исходный DataFrame с данными
        """
        self.df_source = df_source
        self.reference_columns_exist = {}
        self.reference_data = {}
        self._find_reference_columns()

    def _find_column_by_name(self, col_code: str) -> Optional[int]:
        """
        ИСПРАВЛЕНИЕ: Поиск столбца по имени заголовка вместо фиксированного индекса.

        Args:
            col_code: Код столбца (AO, AP, AQ, AR, AS)

        Returns:
            Индекс столбца или None
        """
        expected = self.EXPECTED_NAMES[col_code].lower()
        alternatives = self.ALTERNATIVE_NAMES.get(col_code, [])

        # Ищем по всем столбцам
        for idx, col_name in enumerate(self.df_source.columns):
            col_name_lower = str(col_name).lower().strip()

            # Точное совпадение
            if expected in col_name_lower:
                logger.info(f"  ✅ Найден столбец {col_code}: индекс {idx}, имя '{col_name}'")
                return idx

            # Проверка альтернативных имен
            for alt in alternatives:
                if alt in col_name_lower:
                    logger.info(f"  ✅ Найден столбец {col_code}: индекс {idx}, имя '{col_name}' (альтернатива)")
                    return idx

        logger.warning(f"  ⚠️  Столбец {col_code} не найден")
        return None

    def _find_reference_columns(self) -> None:
        """ИСПРАВЛЕНИЕ: Динамический поиск эталонных столбцов"""
        logger.info("="*80)
        logger.info("ПОИСК ЭТАЛОННЫХ СТОЛБЦОВ (ДИНАМИЧЕСКИЙ)")
        logger.info("="*80)

        for col_code in ['AO', 'AP', 'AQ', 'AR', 'AS']:
            col_idx = self._find_column_by_name(col_code)

            if col_idx is not None:
                self.reference_columns_exist[col_code] = True

                # ИСПРАВЛЕНИЕ: Проверяем первую строку - если это заголовок, пропускаем её
                first_value = self.df_source.iloc[0, col_idx]
                expected_name = self.EXPECTED_NAMES[col_code]

                if pd.notna(first_value) and expected_name.lower() in str(first_value).lower():
                    # Первая строка - заголовок, пропускаем
                    self.reference_data[col_code] = self.df_source.iloc[1:, col_idx].reset_index(drop=True)
                    logger.info(f"      → Данные с 2-й строки (заголовок пропущен)")
                else:
                    # Первая строка - уже данные
                    self.reference_data[col_code] = self.df_source.iloc[:, col_idx].reset_index(drop=True)
                    logger.info(f"      → Данные с 1-й строки")
            else:
                self.reference_columns_exist[col_code] = False
                logger.warning(f"  ❌ Столбец {col_code} не найден в эталоне")

        logger.info("="*80)

    def _is_manual_check_marker(self, ao_val, ap_val, aq_val) -> bool:
        """Проверка на маркеры ручной проверки (без изменений)"""
        ao_empty = pd.isna(ao_val) or str(ao_val).strip() == ""
        ap_empty = pd.isna(ap_val) or str(ap_val).strip() == ""
        aq_zero = pd.isna(aq_val) or str(aq_val).strip() == "" or (isinstance(aq_val, (int, float)) and aq_val == 0)

        if ao_empty and ap_empty and aq_zero:
            return True
        if ao_empty and not ap_empty and "квартал" not in str(ap_val).lower():
            return True

        return False

    def compare_results(self, df_calculated: pd.DataFrame) -> Dict:
        """Сравнение результатов (улучшенная версия)"""
        if not any(self.reference_columns_exist.values()):
            logger.warning("⚠️  Эталонные столбцы не найдены - сравнение невозможно")
            return {'comparison_available': False}

        stats = {
            'comparison_available': True,
            'total_rows': len(df_calculated),
            'columns': {},
            'special_metrics': {
                'human_not_found': 0,
                'program_not_found': 0,
                'both_not_found': 0,
                'human_found_program_not': 0,
                'program_found_human_not': 0
            }
        }

        calc_cols = {
            'AO': 'AO_Дата_приобретения',
            'AP': 'AP_Квартал_приобретения',
            'AQ': 'AQ_Стоимость_закупки',
            'AR': 'AR_Прямая_СС',
            'AS': 'AS_НР'
        }

        for col_code in ['AO', 'AP', 'AQ', 'AR', 'AS']:
            if not self.reference_columns_exist.get(col_code, False):
                stats['columns'][col_code] = {'available': False}
                continue

            ref = self.reference_data[col_code]
            calc = df_calculated[calc_cols[col_code]]

            result = self._compare_column(col_code, ref, calc, stats['special_metrics'])
            stats['columns'][col_code] = result

        return stats

    def _compare_column(self, col_code, ref, calc, sm):
        """Сравнение одного столбца (без изменений)"""
        total = min(len(ref), len(calc))
        exact, mis, empty_ref, empty_calc = 0, 0, 0, 0

        ao_ref = self.reference_data.get('AO')
        ap_ref = self.reference_data.get('AP')
        aq_ref = self.reference_data.get('AQ')

        for i in range(total):
            ref_val = ref.iloc[i]
            calc_val = calc.iloc[i]

            is_manual = False
            if ao_ref is not None and ap_ref is not None and aq_ref is not None:
                is_manual = self._is_manual_check_marker(
                    ao_ref.iloc[i], 
                    ap_ref.iloc[i], 
                    aq_ref.iloc[i]
                )

            calc_empty = (
                pd.isna(calc_val) or 
                str(calc_val).strip() == "" or 
                str(calc_val) in [ERROR_CODES['MANUAL_CHECK'], ERROR_CODES['DATE_NOT_FOUND']]
            )

            ref_empty = pd.isna(ref_val) or str(ref_val).strip() == ""

            if not ref_empty and col_code in ['AQ', 'AR', 'AS']:
                if isinstance(ref_val, (int, float)) and ref_val == 0 and is_manual:
                    ref_empty = True

            if col_code == 'AQ':
                if is_manual:
                    sm['human_not_found'] += 1
                if calc_empty:
                    sm['program_not_found'] += 1
                if is_manual and calc_empty:
                    sm['both_not_found'] += 1
                elif is_manual and not calc_empty:
                    sm['program_found_human_not'] += 1
                elif not is_manual and not ref_empty and calc_empty:
                    sm['human_found_program_not'] += 1

            if ref_empty and calc_empty:
                empty_ref += 1
                empty_calc += 1
                continue

            if ref_empty:
                empty_ref += 1
                continue

            if calc_empty:
                empty_calc += 1
                continue

            # Сравнение значений
            if col_code in ['AQ', 'AR', 'AS']:
                try:
                    ref_float = float(ref_val)
                    calc_float = float(calc_val)
                    if abs(ref_float - calc_float) < 0.01:
                        exact += 1
                    else:
                        mis += 1
                except:
                    mis += 1
            else:
                if str(ref_val).strip() == str(calc_val).strip():
                    exact += 1
                else:
                    mis += 1

        valid = exact + mis
        return {
            'available': True,
            'exact_matches': exact,
            'mismatches': mis,
            'empty_in_reference': empty_ref,
            'empty_in_calculated': empty_calc,
            'exact_match_percent': (exact / valid * 100) if valid > 0 else 0
        }

    def generate_report(self, stats):
        """Генерация отчёта (без изменений)"""
        if not stats.get('comparison_available'):
            return 'Сравнение недоступно - эталонные столбцы не найдены'

        lines = [
            "="*80,
            "ОТЧЁТ О СРАВНЕНИИ",
            "="*80,
            f"Строк: {stats['total_rows']}",
            ""
        ]

        if 'special_metrics' in stats:
            sm = stats['special_metrics']
            lines.extend([
                "ДЕТАЛЬНАЯ СТАТИСТИКА:",
                "-"*80,
                f"  Человек не нашёл: {sm['human_not_found']}",
                f"  Программа не нашла: {sm['program_not_found']}",
                f"  ✅ Оба не нашли: {sm['both_not_found']}",
                f"  ❌ Человек нашёл, программа нет: {sm['human_found_program_not']}",
                f"  ✨ Программа нашла, человек нет: {sm['program_found_human_not']}",
                ""
            ])

        for code in ['AO', 'AP', 'AQ', 'AR', 'AS']:
            cs = stats['columns'].get(code, {})
            lines.extend([
                f"{code}: {self.EXPECTED_NAMES[code]}",
                "-"*80
            ])

            if cs.get('available'):
                lines.extend([
                    f"  Совпадений: {cs['exact_matches']} ({cs['exact_match_percent']:.1f}%)",
                    f"  Несовпадений: {cs['mismatches']}",
                    ""
                ])
            else:
                lines.extend([
                    "  ⚠️  Эталонный столбец не найден",
                    ""
                ])

        lines.append("="*80)
        return "\n".join(lines)
