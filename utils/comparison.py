"""Модуль для сравнения эталонных и расчётных данных"""
import pandas as pd
import re
from typing import Dict, List
from utils.logger import get_logger

logger = get_logger(__name__)

class DataComparison:
    REFERENCE_COLUMNS = {'AO': 40, 'AP': 41, 'AQ': 42, 'AR': 43, 'AS': 44}
    EXPECTED_NAMES = {
        'AO': 'Дата приобретения',
        'AP': 'Квартал приобретения',
        'AQ': 'Стоимость закупки НЧТЗ 1 ед',
        'AR': 'Прямая СС НЧТЗ 1 ед',
        'AS': 'НР НЧТЗ 1 ед',
    }

    def __init__(self, df_source: pd.DataFrame):
        self.df_source = df_source
        self.logger = logger
        self.reference_columns_exist = {}
        self.reference_data = {}
        self.program_found_human_not_examples = []  # Новое: хранилище примеров
        self._check_reference_columns()

    def _is_excel_not_empty(self, value):
        """
        Проверяет "непустоту" значения так же, как это делает фильтр Excel
        Excel считает ПУСТЫМИ:
        - NaN, None
        - Пустые строки ""
        - Строки только из пробелов "   "
        - Строку "nan" (как текст)
        - Тире "-"
        - Строку "none"
        """
        if pd.isna(value):
            return False

        str_value = str(value).strip().lower()
        empty_values = ["", "nan", "none", "-", "—"]

        return str_value not in empty_values

    def _is_valid_quarter(self, quarter_value) -> bool:
        """Проверяет что квартал имеет валидный формат 'N квартал YYYY' (N=1-4, YYYY=2020-2030)"""
        if not self._is_excel_not_empty(quarter_value):
            return False

        str_value = str(quarter_value).strip()
        pattern = r'^([1-4])\s+квартал\s+(20[2-3][0-9])$'

        return bool(re.match(pattern, str_value))

    def _is_intentional_skip(self, quarter_value) -> bool:
        """Возвращает True если значение НЕ соответствует формату квартала (специально пропущено)"""
        return not self._is_valid_quarter(quarter_value)

    def _check_reference_columns(self) -> None:
        logger.info("Проверка наличия эталонных столбцов...")
        for col_code, col_idx in self.REFERENCE_COLUMNS.items():
            if col_idx < len(self.df_source.columns):
                col_name = self.df_source.columns[col_idx]
                expected_name = self.EXPECTED_NAMES[col_code]
                
                check_row_idx = 9
                if len(self.df_source) > check_row_idx:
                    row_value = str(self.df_source.iloc[check_row_idx, col_idx])
                    if expected_name.lower() in row_value.lower():
                        self.reference_columns_exist[col_code] = True
                        self.reference_data[col_code] = self.df_source.iloc[check_row_idx+1:, col_idx].reset_index(drop=True)
                        logger.info(f"  ✅ Столбец {col_code} ({expected_name}) найден: '{row_value}'")
                    else:
                        self.reference_columns_exist[col_code] = False
                        logger.warning(f" ⚠️ Столбец {col_code}: название не соответствует")
                elif len(self.df_source) > 0:
                    first_row_value = str(self.df_source.iloc[0, col_idx])
                    if expected_name.lower() in first_row_value.lower():
                        self.reference_columns_exist[col_code] = True
                        self.reference_data[col_code] = self.df_source.iloc[1:, col_idx].reset_index(drop=True)
                        logger.info(f" ✅ Столбец {col_code} ({expected_name}) найден: '{first_row_value}'")
                    else:
                        self.reference_columns_exist[col_code] = False
                        logger.warning(f"  ⚠️  Столбец {col_code}: название не соответствует")
                else:
                    self.reference_columns_exist[col_code] = False
                    logger.warning(f"  ❌ Столбец {col_code}: пустой файл")
            else:
                self.reference_columns_exist[col_code] = False
                logger.warning(f"  ❌ Столбец {col_code}: индекс {col_idx} за пределами таблицы")
        if not any(self.reference_columns_exist.values()):
            logger.warning("В исходном файле нет эталонных столбцов для сравнения")

    def compare_results(self, df_calculated: pd.DataFrame) -> Dict:
        if not any(self.reference_columns_exist.values()):
            logger.warning("Сравнение невозможно")
            return {'comparison_available': False, 'message': 'В исходном файле нет эталонных столбцов'}

        logger.info("========== СРАВНЕНИЕ ЭТАЛОННЫХ И РАСЧЁТНЫХ ДАННЫХ ==========")
        stats = {'comparison_available': True, 'total_rows': len(df_calculated), 'columns': {}}
        calculated_columns = {
            'AO': 'AO_**Дата_приобретения**', 'AP': 'AP_**Квартал_приобретения**',
            'AQ': 'AQ_**Стоимость_закупки**', 'AR': 'AR_**Прямая_СС**', 'AS': 'AS_**НР**',
        }

        # Инициализируем специальные метрики
        self.special_metrics = {
            'human_not_found': 0,
            'nothing_found': 0,
            'date_not_found_year_only': 0,
            'date_not_found_quarter_computed': 0,
            'cost_not_found': 0,
            'program_not_found': 0,
            'both_not_found': 0,
            'both_not_found_intentional': 0,
            'both_not_found_real': 0,
            'human_found_program_not': 0,
            'program_found_human_not': 0
        }
        
        # Очищаем список примеров перед новым сравнением
        self.program_found_human_not_examples = []

        for col_code in ['AO', 'AP', 'AQ', 'AR', 'AS']:
            if not self.reference_columns_exist.get(col_code, False):
                stats['columns'][col_code] = {'available': False, 'message': f'Эталонный столбец {col_code} отсутствует'}
                logger.info(f"Столбец {col_code}: эталонные данные отсутствуют")
                continue

            reference_col = self.reference_data[col_code]
            calculated_col = df_calculated[calculated_columns[col_code]]
            comparison_result = self._compare_column(col_code, reference_col, calculated_col, df_calculated)
            stats['columns'][col_code] = comparison_result

            logger.info(f"Столбец {col_code} ({self.EXPECTED_NAMES[col_code]}):")
            logger.info(f"  Полных совпадений: {comparison_result['exact_matches']} ({comparison_result['exact_match_percent']:.1f}%)")
            logger.info(f" Несовпадений: {comparison_result['mismatches']} ({comparison_result['mismatch_percent']:.1f}%)")

        stats['special_metrics'] = self.special_metrics
        stats['program_found_human_not_examples'] = self.program_found_human_not_examples

        return stats

    def _compare_column(self, col_code: str, reference: pd.Series, calculated: pd.Series, df_calculated: pd.DataFrame) -> Dict:
        total_rows = len(reference)
        exact_matches = 0
        mismatches = 0
        empty_in_reference = 0
        empty_in_calculated = 0
        mismatch_examples = []

        for idx in range(total_rows):
            ref_val = reference.iloc[idx]
            calc_val = calculated.iloc[idx]
            
            if col_code == 'AQ':
                is_formula = str(ref_val).startswith('=') if not pd.isna(ref_val) else False
                is_zero = str(ref_val).strip() == "0" if not pd.isna(ref_val) else False
                ref_empty = is_formula or is_zero
            else:
                ref_empty = not self._is_excel_not_empty(ref_val)
            
            calc_empty = pd.isna(calc_val) or str(calc_val).strip() == "" or "*ТРЕБУЕТ РУЧНОЙ ПРОВЕРКИ*" in str(calc_val) or calc_val == "#РП" or calc_val == "#НД"

            if ref_empty:
                empty_in_reference += 1
                
                if col_code == 'AQ':
                    quarter_ref = self.reference_data.get('AP')
                    if quarter_ref is not None and idx < len(quarter_ref):
                        quarter_value = quarter_ref.iloc[idx]
                    else:
                        quarter_value = None
                    
                    date_ref = self.reference_data.get('AO')
                    if date_ref is not None and idx < len(date_ref):
                        date_value = date_ref.iloc[idx]
                    else:
                        date_value = None
                    
                    is_formula = str(ref_val).startswith('=СУММ') if not pd.isna(ref_val) else False
                    
                    self.special_metrics['human_not_found'] += 1
                    
                    if pd.isna(date_value) or str(date_value).strip() == "":
                        if pd.isna(quarter_value) or str(quarter_value).strip() == "":
                            self.special_metrics['nothing_found'] += 1
                        elif self._is_valid_quarter(quarter_value):
                            self.special_metrics['date_not_found_quarter_computed'] += 1
                        else:
                            self.special_metrics['date_not_found_year_only'] += 1
                    else:
                        if not is_formula and (pd.isna(ref_val) or str(ref_val).strip() == "" or str(ref_val).strip() == "0"):
                            self.special_metrics['cost_not_found'] += 1
                    
                    if calc_empty:
                        self.special_metrics['both_not_found'] += 1
                        
                        if pd.isna(date_value) or str(date_value).strip() == "":
                            if pd.isna(quarter_value) or str(quarter_value).strip() == "":
                                self.special_metrics['both_not_found_intentional'] += 1
                            elif self._is_valid_quarter(quarter_value):
                                self.special_metrics['both_not_found_real'] += 1
                            else:
                                self.special_metrics['both_not_found_intentional'] += 1
                        else:
                            self.special_metrics['both_not_found_real'] += 1
                    else:
                        # НОВОЕ: Программа нашла, человек нет - сохраняем пример
                        self.special_metrics['program_found_human_not'] += 1
                        
                        if len(self.program_found_human_not_examples) < 30:
                            # Получаем дополнительную информацию из целевой таблицы
                            # Предполагаем, что столбцы NOMENCLATURE=18, DOCUMENT=21
                            nomenclature = df_calculated.iloc[idx, 18] if 18 < len(df_calculated.columns) else "N/A"
                            document = df_calculated.iloc[idx, 21] if 21 < len(df_calculated.columns) else "N/A"
                            
                            self.program_found_human_not_examples.append({
                                'row': idx + 12,  # +12 т.к. данные начинаются с 12-й строки в Excel
                                'nomenclature': nomenclature,
                                'document': document,
                                'calculated_date': df_calculated.iloc[idx]['AO_**Дата_приобретения**'],
                                'calculated_quarter': df_calculated.iloc[idx]['AP_**Квартал_приобретения**'],
                                'calculated_cost': calc_val
                            })
                 
                continue
            elif calc_empty:
                empty_in_calculated += 1
                self.special_metrics['program_not_found'] += 1
                if col_code == 'AQ':
                    self.special_metrics['human_found_program_not'] += 1
                if len(mismatch_examples) < 10:
                    mismatch_examples.append({'row': idx + 2, 'reference': ref_val, 'calculated': calc_val, 'reason': 'Расчёт не выполнен'})
                continue
            else:
                if col_code in ['AQ', 'AR', 'AS']:
                    try:
                        ref_num = float(ref_val)
                        calc_num = float(calc_val)
                        if abs(ref_num - calc_num) < 0.01:
                            exact_matches += 1
                        else:
                            mismatches += 1
                            if len(mismatch_examples) < 10:
                                mismatch_examples.append({'row': idx + 2, 'reference': f"{ref_num:.2f}", 'calculated': f"{calc_num:.2f}",
                                                         'difference': f"{calc_num - ref_num:+.2f}", 'reason': 'Разные значения'})
                    except (ValueError, TypeError):
                        mismatches += 1
                        if len(mismatch_examples) < 10:
                            mismatch_examples.append({'row': idx + 2, 'reference': ref_val, 'calculated': calc_val, 'reason': 'Невозможно сравнить как числа'})
                else:
                    if str(ref_val).strip() == str(calc_val).strip():
                        exact_matches += 1
                    else:
                        mismatches += 1
                        if len(mismatch_examples) < 10:
                            mismatch_examples.append({'row': idx + 2, 'reference': ref_val, 'calculated': calc_val, 'reason': 'Разные значения'})

        valid_comparisons = exact_matches + mismatches
        exact_match_percent = (exact_matches / valid_comparisons * 100) if valid_comparisons > 0 else 0
        mismatch_percent = (mismatches / valid_comparisons * 100) if valid_comparisons > 0 else 0

        return {
            'available': True, 'total_rows': total_rows, 'exact_matches': exact_matches, 'mismatches': mismatches,
            'empty_in_reference': empty_in_reference, 'empty_in_calculated': empty_in_calculated,
            'exact_match_percent': exact_match_percent, 'mismatch_percent': mismatch_percent,
            'mismatch_examples': mismatch_examples
        }

    def generate_report(self, stats: Dict) -> str:
        if not stats.get('comparison_available', False):
            return stats.get('message', 'Сравнение недоступно')

        report_lines = ["=" * 80, "ОТЧЁТ О СРАВНЕНИИ ЭТАЛОННЫХ И РАСЧЁТНЫХ ДАННЫХ", "=" * 80, f"Всего строк: {stats['total_rows']}", ""]

        if 'special_metrics' in stats:
            special_metrics = stats['special_metrics']
            report_lines.append("ДЕТАЛЬНАЯ СТАТИСТИКА:")
            report_lines.append("-" * 80)
            report_lines.append(f" Человек не нашёл: {special_metrics['human_not_found']}")
            report_lines.append(f"   ├─ Ничего не найдено: {special_metrics.get('nothing_found', 0)}")
            report_lines.append(f"   ├─ Не найдена Дата приобретения только год: {special_metrics.get('date_not_found_year_only', 0)}")
            report_lines.append(f"   ├─ Не найдена Дата приобретения но вычислен квартал: {special_metrics.get('date_not_found_quarter_computed', 0)}")
            report_lines.append(f"   └─ Не найдена Стоимость: {special_metrics.get('cost_not_found', 0)}")
            report_lines.append(f" Программа не нашла: {special_metrics['program_not_found']}")
            report_lines.append(f"  ✅ Оба не нашли (согласие): {special_metrics['both_not_found']}")
            report_lines.append(f"    ├─ Человек не искал: {special_metrics['both_not_found_intentional']}")
            report_lines.append(f"    └─ Человек искал, но не нашёл: {special_metrics['both_not_found_real']}")
            report_lines.append(f" ❌ Человек нашёл, программа нет: {special_metrics['human_found_program_not']}")
            report_lines.append(f"  ✨ Программа нашла, человек нет: {special_metrics['program_found_human_not']}")
            report_lines.append("")
            
            total_records = special_metrics['human_not_found'] + special_metrics['human_found_program_not']
            report_lines.append(f" Всего записей в классификации: {total_records}")
            
            human_total_check = (special_metrics['nothing_found'] +
                                special_metrics['date_not_found_year_only'] +
                                special_metrics['date_not_found_quarter_computed'] +
                                special_metrics['cost_not_found'])
            if special_metrics['human_not_found'] != human_total_check:
                report_lines.append(f" ⚠️ ПРЕДУПРЕЖДЕНИЕ: Сумма подкатегорий 'человек не нашёл' ({human_total_check}) не равна общей сумме ({special_metrics['human_not_found']})")
            report_lines.append("")

        for col_code in ['AO', 'AP', 'AQ', 'AR', 'AS']:
            col_stats = stats['columns'].get(col_code, {})
            col_name = self.EXPECTED_NAMES[col_code]
            report_lines.append(f"Столбец {col_code}: {col_name}")
            report_lines.append("-" * 80)

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

        # НОВОЕ: Добавляем примеры "Программа нашла, человек нет"
        if 'program_found_human_not_examples' in stats and stats['program_found_human_not_examples']:
            report_lines.append("=" * 80)
            report_lines.append("✨ ПРИМЕРЫ: ПРОГРАММА НАШЛА, ЧЕЛОВЕК НЕТ (до 30 шт)")
            report_lines.append("=" * 80)
            
            for idx, example in enumerate(stats['program_found_human_not_examples'][:30], 1):
                report_lines.append(f"\n{idx}. Строка Excel: {example['row']}")
                report_lines.append(f"   Номенклатура: {example['nomenclature']}")
                report_lines.append(f"   Документ: {example['document']}")
                report_lines.append(f"   Программа нашла:")
                report_lines.append(f"     - Дата: {example['calculated_date']}")
                report_lines.append(f"     - Квартал: {example['calculated_quarter']}")
                report_lines.append(f"     - Стоимость: {example['calculated_cost']}")
            
            report_lines.append("")

        report_lines.append("=" * 80)
        return "\n".join(report_lines)
