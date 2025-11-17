"""Модуль для сравнения эталонных и расчётных данных (УЛУЧШЕННАЯ ВЕРСИЯ)"""
import pandas as pd
from typing import Dict, Optional
from utils.logger import get_logger

logger = get_logger(__name__)


class DataComparison:
    """
    УЛУЧШЕННАЯ ВЕРСИЯ: Сочетает динамический поиск столбцов 
    с детальной статистикой из рабочей версии.
    """

    EXPECTED_NAMES = {
        'AO': 'Дата приобретения',
        'AP': 'Квартал приобретения',
        'AQ': 'Стоимость закупки НЧТЗ 1 ед',
        'AR': 'Прямая СС НЧТЗ 1 ед',
        'AS': 'НР НЧТЗ 1 ед',
    }

    # АЛЬТЕРНАТИВНЫЕ ИМЕНА для поиска
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
        self.logger = logger
        self.reference_columns_exist = {}
        self.reference_data = {}
        self._find_reference_columns()

    def _find_column_by_name(self, col_code: str) -> Optional[int]:
        """
        ДИНАМИЧЕСКИЙ ПОИСК: Ищет столбец по имени заголовка.

        Args:
            col_code: Код столбца (AO, AP, AQ, AR, AS)

        Returns:
            Индекс столбца или None
        """
        expected = self.EXPECTED_NAMES[col_code].lower()
        alternatives = self.ALTERNATIVE_NAMES.get(col_code, [])

        # Сначала ищем в названиях столбцов
        for idx, col_name in enumerate(self.df_source.columns):
            col_name_lower = str(col_name).lower().strip()

            # Точное совпадение в названии столбца
            if expected in col_name_lower:
                logger.info(f"  ✅ Найден столбец {col_code} по названию: индекс {idx}, имя '{col_name}'")
                return idx

            # Альтернативные имена
            for alt in alternatives:
                if alt in col_name_lower:
                    logger.info(f"  ✅ Найден столбец {col_code} (альтернатива): индекс {idx}, имя '{col_name}'")
                    return idx

        # Если не нашли в названиях столбцов, ищем в ЗНАЧЕНИЯХ (как в вашем коде)
        # Проверяем строки 0-15 на наличие заголовка
        for row_idx in range(min(15, len(self.df_source))):
            for col_idx in range(len(self.df_source.columns)):
                try:
                    cell_value = str(self.df_source.iloc[row_idx, col_idx]).lower().strip()

                    # Проверяем точное совпадение
                    if expected in cell_value:
                        logger.info(f"  ✅ Найден столбец {col_code} в значениях: индекс {col_idx}, строка {row_idx}, значение '{cell_value}'")
                        return col_idx

                    # Проверяем альтернативы
                    for alt in alternatives:
                        if alt in cell_value:
                            logger.info(f"  ✅ Найден столбец {col_code} (альтернатива) в значениях: индекс {col_idx}, строка {row_idx}")
                            return col_idx
                except:
                    continue

        logger.warning(f"  ⚠️  Столбец {col_code} не найден")
        return None

    def _find_header_row(self, col_idx: int) -> int:
        """
        Определяет строку с заголовком для данного столбца.

        Args:
            col_idx: Индекс столбца

        Returns:
            Индекс строки с заголовком
        """
        # Ищем строку, где находится название столбца
        for row_idx in range(min(15, len(self.df_source))):
            try:
                cell_value = str(self.df_source.iloc[row_idx, col_idx]).lower().strip()

                # Проверяем, есть ли в этой строке какое-то из ожидаемых названий
                for expected in self.EXPECTED_NAMES.values():
                    if expected.lower() in cell_value:
                        logger.info(f"      → Заголовок найден в строке {row_idx}")
                        return row_idx
            except:
                continue

        # По умолчанию считаем, что заголовок в строке 0
        return 0

    def _find_reference_columns(self) -> None:
        """ДИНАМИЧЕСКИЙ ПОИСК эталонных столбцов"""
        logger.info("="*80)
        logger.info("ПОИСК ЭТАЛОННЫХ СТОЛБЦОВ (ДИНАМИЧЕСКИЙ)")
        logger.info("="*80)
        logger.info(f"Всего столбцов: {len(self.df_source.columns)}")
        logger.info(f"Всего строк: {len(self.df_source)}")
        logger.info("")

        for col_code in ['AO', 'AP', 'AQ', 'AR', 'AS']:
            logger.info(f"Поиск столбца {col_code} ({self.EXPECTED_NAMES[col_code]})...")

            col_idx = self._find_column_by_name(col_code)

            if col_idx is not None:
                self.reference_columns_exist[col_code] = True

                # Определяем строку с заголовком
                header_row = self._find_header_row(col_idx)

                # Берём данные начиная со следующей строки после заголовка
                data_start_row = header_row + 1
                self.reference_data[col_code] = self.df_source.iloc[data_start_row:, col_idx].reset_index(drop=True)

                logger.info(f"      → Данные начинаются со строки {data_start_row}")
                logger.info(f"      → Всего значений: {len(self.reference_data[col_code])}")
            else:
                self.reference_columns_exist[col_code] = False
                logger.warning(f"  ❌ Столбец {col_code} не найден в эталоне")

            logger.info("")

        if not any(self.reference_columns_exist.values()):
            logger.warning("⚠️  В исходном файле нет эталонных столбцов для сравнения")

        logger.info("="*80)

    def compare_results(self, df_calculated: pd.DataFrame) -> Dict:
        """Сравнение результатов с детальной статистикой"""
        if not any(self.reference_columns_exist.values()):
            logger.warning("⚠️  Сравнение невозможно - эталонные столбцы не найдены")
            return {
                'comparison_available': False, 
                'message': 'В исходном файле нет эталонных столбцов'
            }

        logger.info("="*80)
        logger.info("СРАВНЕНИЕ ЭТАЛОННЫХ И РАСЧЁТНЫХ ДАННЫХ")
        logger.info("="*80)

        stats = {
            'comparison_available': True, 
            'total_rows': len(df_calculated), 
            'columns': {}
        }

        calculated_columns = {
            'AO': 'AO_Дата_приобретения', 
            'AP': 'AP_Квартал_приобретения',
            'AQ': 'AQ_Стоимость_закупки', 
            'AR': 'AR_Прямая_СС', 
            'AS': 'AS_НР',
        }

        for col_code in ['AO', 'AP', 'AQ', 'AR', 'AS']:
            if not self.reference_columns_exist.get(col_code, False):
                stats['columns'][col_code] = {
                    'available': False, 
                    'message': f'Эталонный столбец {col_code} отсутствует'
                }
                logger.info(f"Столбец {col_code}: эталонные данные отсутствуют")
                continue

            reference_col = self.reference_data[col_code]
            calculated_col = df_calculated[calculated_columns[col_code]]

            comparison_result = self._compare_column(col_code, reference_col, calculated_col)
            stats['columns'][col_code] = comparison_result

            logger.info(f"Столбец {col_code} ({self.EXPECTED_NAMES[col_code]}):")
            logger.info(f"  Полных совпадений: {comparison_result['exact_matches']} ({comparison_result['exact_match_percent']:.1f}%)")
            logger.info(f"  Несовпадений: {comparison_result['mismatches']} ({comparison_result['mismatch_percent']:.1f}%)")

        logger.info("="*80)
        return stats

    def _compare_column(self, col_code: str, reference: pd.Series, calculated: pd.Series) -> Dict:
        """Детальное сравнение одного столбца"""
        total_rows = min(len(reference), len(calculated))
        exact_matches = 0
        mismatches = 0
        empty_in_reference = 0
        empty_in_calculated = 0
        mismatch_examples = []

        for idx in range(total_rows):
            ref_val = reference.iloc[idx] if idx < len(reference) else None
            calc_val = calculated.iloc[idx] if idx < len(calculated) else None

            ref_empty = pd.isna(ref_val) or str(ref_val).strip() == ""
            calc_empty = (
                pd.isna(calc_val) or 
                str(calc_val).strip() == "" or 
                "#РП" in str(calc_val) or 
                "#НД" in str(calc_val)
            )

            if ref_empty:
                empty_in_reference += 1
                continue

            if calc_empty:
                empty_in_calculated += 1
                if len(mismatch_examples) < 10:
                    mismatch_examples.append({
                        'row': idx + 2, 
                        'reference': ref_val, 
                        'calculated': calc_val, 
                        'reason': 'Расчёт не выполнен'
                    })
                continue

            # Сравнение значений
            if col_code in ['AQ', 'AR', 'AS']:
                # Числовое сравнение для стоимостных столбцов
                try:
                    ref_num = float(ref_val)
                    calc_num = float(calc_val)

                    if abs(ref_num - calc_num) < 0.01:
                        exact_matches += 1
                    else:
                        mismatches += 1
                        if len(mismatch_examples) < 10:
                            mismatch_examples.append({
                                'row': idx + 2, 
                                'reference': f"{ref_num:.2f}", 
                                'calculated': f"{calc_num:.2f}", 
                                'difference': f"{calc_num - ref_num:+.2f}", 
                                'reason': 'Разные значения'
                            })
                except (ValueError, TypeError):
                    mismatches += 1
                    if len(mismatch_examples) < 10:
                        mismatch_examples.append({
                            'row': idx + 2,
                            'reference': str(ref_val),
                            'calculated': str(calc_val),
                            'reason': 'Ошибка преобразования в число'
                        })
            else:
                # Текстовое сравнение для дат и кварталов
                if str(ref_val).strip() == str(calc_val).strip():
                    exact_matches += 1
                else:
                    mismatches += 1
                    if len(mismatch_examples) < 10:
                        mismatch_examples.append({
                            'row': idx + 2,
                            'reference': str(ref_val),
                            'calculated': str(calc_val),
                            'reason': 'Разные значения'
                        })

        valid_comparisons = exact_matches + mismatches
        exact_match_percent = (exact_matches / valid_comparisons * 100) if valid_comparisons > 0 else 0
        mismatch_percent = (mismatches / valid_comparisons * 100) if valid_comparisons > 0 else 0

        return {
            'available': True, 
            'total_rows': total_rows, 
            'exact_matches': exact_matches, 
            'mismatches': mismatches,
            'empty_in_reference': empty_in_reference, 
            'empty_in_calculated': empty_in_calculated,
            'exact_match_percent': exact_match_percent, 
            'mismatch_percent': mismatch_percent,
            'mismatch_examples': mismatch_examples
        }

    def generate_report(self, stats: Dict) -> str:
        """Генерация детального отчёта"""
        if not stats.get('comparison_available', False):
            return stats.get('message', 'Сравнение недоступно - эталонные столбцы не найдены')

        report_lines = [
            "="*80,
            "ОТЧЁТ О СРАВНЕНИИ ЭТАЛОННЫХ И РАСЧЁТНЫХ ДАННЫХ",
            "="*80,
            f"Всего строк: {stats['total_rows']}",
            ""
        ]

        for col_code in ['AO', 'AP', 'AQ', 'AR', 'AS']:
            col_stats = stats['columns'].get(col_code, {})
            col_name = self.EXPECTED_NAMES[col_code]

            report_lines.append(f"Столбец {col_code}: {col_name}")
            report_lines.append("-"*80)

            if not col_stats.get('available', False):
                report_lines.append(f"  {col_stats.get('message', 'Данные недоступны')}")
                report_lines.append("")
                continue

            report_lines.append(f"  Полных совпадений: {col_stats['exact_matches']} ({col_stats['exact_match_percent']:.1f}%)")
            report_lines.append(f"  Несовпадений: {col_stats['mismatches']} ({col_stats['mismatch_percent']:.1f}%)")
            report_lines.append(f"  Пустых в эталоне: {col_stats['empty_in_reference']}")
            report_lines.append(f"  Пустых в расчёте: {col_stats['empty_in_calculated']}")

            if col_stats['mismatch_examples']:
                report_lines.append("")
                report_lines.append("  Примеры несовпадений (до 10 шт):")
                for example in col_stats['mismatch_examples'][:10]:
                    report_lines.append(f"    Строка {example['row']}:")
                    report_lines.append(f"      Эталон:  {example['reference']}")
                    report_lines.append(f"      Расчёт:  {example['calculated']}")
                    if 'difference' in example:
                        report_lines.append(f"      Разница: {example['difference']}")
                    report_lines.append(f"      Причина: {example['reason']}")

            report_lines.append("")

        report_lines.append("="*80)
        return "\n".join(report_lines)