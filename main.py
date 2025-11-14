"""ExcelCostCalculator - Main Entry Point"""
import argparse
import time
import gc
import pandas as pd
from config import *
from core import load_excel_file, FormulaEngine, write_results
from utils import setup_logger, get_logger, extract_acquisition_date, determine_quarter, DataComparison

logger = None

def process_chunk(df_chunk, engine, chunk_start_idx):
    """Обрабатывает chunk строк"""
    success_count = 0
    error_count = 0
    for idx in range(len(df_chunk)):
        global_idx = chunk_start_idx + idx
        excel_row = global_idx + 2
        document_text = df_chunk.iloc[idx, TARGET_COLUMNS['DOCUMENT']]
        nomenclature = df_chunk.iloc[idx, TARGET_COLUMNS['NOMENCLATURE']]
        acquisition_date = extract_acquisition_date(document_text)
        if not acquisition_date:
            logger.debug(f"Строка {excel_row}: Дата не найдена")
            df_chunk.at[df_chunk.index[idx], 'AO_Дата_приобретения'] = ERROR_CODES['DATE_NOT_FOUND']
            df_chunk.at[df_chunk.index[idx], 'AP_Квартал_приобретения'] = ""
            df_chunk.at[df_chunk.index[idx], 'AQ_Стоимость_закупки'] = ERROR_CODES['MANUAL_CHECK']
            df_chunk.at[df_chunk.index[idx], 'AR_Прямая_СС'] = ERROR_CODES['MANUAL_CHECK']
            df_chunk.at[df_chunk.index[idx], 'AS_НР'] = ERROR_CODES['MANUAL_CHECK']
            error_count += 1
            continue
        df_chunk.at[df_chunk.index[idx], 'AO_Дата_приобретения'] = acquisition_date.strftime(DATE_FORMAT_OUTPUT)
        quarter = determine_quarter(acquisition_date)
        df_chunk.at[df_chunk.index[idx], 'AP_Квартал_приобретения'] = quarter
        aq = engine.calculate_aq(nomenclature, quarter)
        ar = engine.calculate_ar(nomenclature, quarter)
        as_val = engine.calculate_as(nomenclature, quarter)
        if aq is not None:
            df_chunk.at[df_chunk.index[idx], 'AQ_Стоимость_закупки'] = aq
            df_chunk.at[df_chunk.index[idx], 'AR_Прямая_СС'] = ar
            df_chunk.at[df_chunk.index[idx], 'AS_НР'] = as_val
            success_count += 1
        else:
            logger.debug(f"Строка {excel_row}: Нет совпадений")
            df_chunk.at[df_chunk.index[idx], 'AQ_Стоимость_закупки'] = ERROR_CODES['MANUAL_CHECK']
            df_chunk.at[df_chunk.index[idx], 'AR_Прямая_СС'] = ERROR_CODES['MANUAL_CHECK']
            df_chunk.at[df_chunk.index[idx], 'AS_НР'] = ERROR_CODES['MANUAL_CHECK']
            error_count += 1
    return df_chunk, success_count, error_count

def main():
    global logger
    parser = argparse.ArgumentParser(description='ExcelCostCalculator v2')
    parser.add_argument('--input', required=True, help='Входной Excel-файл')
    parser.add_argument('--output', required=True, help='Выходной Excel-файл')
    parser.add_argument('--log-level', default='INFO', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'])
    parser.add_argument('--chunk-size', type=int, default=DEFAULT_CHUNK_SIZE)
    parser.add_argument('--compare', action='store_true', help='Сравнить с эталонными данными')
    args = parser.parse_args()
    setup_logger(args.log_level)
    logger = get_logger(__name__)
    logger.info("="*60)
    logger.info("ExcelCostCalculator v2")
    logger.info("="*60)
    logger.info(f"Входной файл: {args.input}")
    logger.info(f"Выходной файл: {args.output}")
    logger.info(f"Размер chunk: {args.chunk_size} строк")
    logger.info(f"Режим сравнения: {'ВКЛЮЧЕН' if args.compare else 'ВЫКЛЮЧЕН'}")
    start_time = time.time()
    try:
        df_target, df_source = load_excel_file(args.input, SHEET_NAMES['TARGET'], SHEET_NAMES['SOURCE'])
        logger.info(f"Целевая таблица: {len(df_target)} строк")
        logger.info(f"Справочная таблица: {len(df_source)} строк")
        logger.info("Инициализация FormulaEngine...")
        engine = FormulaEngine(df_source, SOURCE_COLUMNS)
        df_result = df_target.copy()
        df_result['AO_Дата_приобретения'] = ""
        df_result['AP_Квартал_приобретения'] = ""
        df_result['AQ_Стоимость_закупки'] = ""
        df_result['AR_Прямая_СС'] = ""
        df_result['AS_НР'] = ""
        logger.info("Начало обработки по частям...")
        total_rows = len(df_result)
        num_chunks = (total_rows + args.chunk_size - 1) // args.chunk_size
        logger.info(f"Общее количество chunks: {num_chunks}")
        total_success = 0
        total_errors = 0
        for chunk_num in range(num_chunks):
            chunk_start = chunk_num * args.chunk_size
            chunk_end = min((chunk_num + 1) * args.chunk_size, total_rows)
            df_chunk = df_result.iloc[chunk_start:chunk_end].copy()
            logger.info(f"Обработка chunk {chunk_num + 1}/{num_chunks} (строки {chunk_start + 1}-{chunk_end})...")
            df_chunk, success, errors = process_chunk(df_chunk, engine, chunk_start)
            df_result.iloc[chunk_start:chunk_end] = df_chunk
            total_success += success
            total_errors += errors
            logger.info(f"Обработано: {chunk_end}/{total_rows} ({chunk_end/total_rows*100:.1f}%) | Успешно: {total_success} | Ошибок: {total_errors}")
            if (chunk_num + 1) % GC_INTERVAL == 0:
                gc.collect()
        logger.info("Финальная очистка памяти...")
        gc.collect()
        logger.info("Сохранение результатов...")
        write_results(args.input, args.output, df_result, SHEET_NAMES['TARGET'])
        if args.compare:
            logger.info("Сравнение с эталонными данными...")
            df_original = pd.read_excel(args.input, sheet_name=SHEET_NAMES['TARGET'])
            comparator = DataComparison(df_original)
            stats = comparator.compare_results(df_result)
            report = comparator.generate_report(stats)
            report_path = f"{args.output.replace('.xlsx', '')}_comparison_report.txt"
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report)
            logger.info(f"Отчёт сравнения сохранён: {report_path}")
            print("\n" + report)
        elapsed = time.time() - start_time
        logger.info("="*60)
        logger.info("ИТОГОВАЯ СТАТИСТИКА")
        logger.info("="*60)
        logger.info(f"Всего обработано записей: {total_rows}")
        logger.info(f"Успешно сопоставлено: {total_success} ({total_success/total_rows*100:.1f}%)")
        logger.info(f"Требует ручной проверки: {total_errors} ({total_errors/total_rows*100:.1f}%)")
        logger.info(f"Время обработки: {elapsed:.1f} сек ({elapsed/60:.1f} мин)")
        logger.info("Обработка завершена успешно!")
        print("\n" + "="*60)
        print("ЛЕГЕНДА КОДОВ ОШИБОК:")
        print("="*60)
        for code, description in ERROR_DESCRIPTIONS.items():
            print(f"  {code} — {description}")
        print("="*60)
    except Exception as e:
        logger.error(f"ОШИБКА: {e}", exc_info=True)
        raise

if __name__ == "__main__":
    main()
