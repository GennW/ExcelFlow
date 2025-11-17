"""Точка входа в приложение ExcelCostCalculator v2"""
import argparse
import sys
import gc
from pathlib import Path
import pandas as pd
from datetime import datetime

from config import TARGET_COLUMNS, SOURCE_COLUMNS, SHEET_NAMES, DATE_FORMAT_OUTPUT, DEFAULT_CHUNK_SIZE, GC_INTERVAL
from core import load_excel_file, FormulaEngine, write_results
from utils import extract_acquisition_date, determine_quarter, setup_logger, get_logger, DataComparison

def get_memory_usage():
    try:
        import psutil
        process = psutil.Process()
        return process.memory_info().rss / 1024 / 1024
    except ImportError:
        return 0

def process_chunk(df_chunk, engine, chunk_start_idx, logger):
    success_count = 0
    error_count = 0

    for idx in range(len(df_chunk)):
        global_idx = chunk_start_idx + idx
        document_text = df_chunk.iloc[idx, TARGET_COLUMNS['DOCUMENT']]
        nomenclature = df_chunk.iloc[idx, TARGET_COLUMNS['NOMENCLATURE']]
        acquisition_date = extract_acquisition_date(document_text)

        if not acquisition_date:
            logger.debug(f"Строка {global_idx+2}: Не удалось извлечь дату")
            df_chunk.at[df_chunk.index[idx], 'AO_**Дата_приобретения**'] = "#НД"
            df_chunk.at[df_chunk.index[idx], 'AP_**Квартал_приобретения**'] = ""
            df_chunk.at[df_chunk.index[idx], 'AQ_**Стоимость_закупки**'] = "#РП"
            df_chunk.at[df_chunk.index[idx], 'AR_**Прямая_СС**'] = "#РП"
            df_chunk.at[df_chunk.index[idx], 'AS_**НР**'] = "#РП"
            error_count += 1
            continue

        df_chunk.at[df_chunk.index[idx], 'AO_**Дата_приобретения**'] = acquisition_date.strftime(DATE_FORMAT_OUTPUT)
        quarter = determine_quarter(acquisition_date)
        df_chunk.at[df_chunk.index[idx], 'AP_**Квартал_приобретения**'] = quarter

        aq = engine.calculate_aq(nomenclature, quarter)
        ar = engine.calculate_ar(nomenclature, quarter)
        as_val = engine.calculate_as(nomenclature, quarter)

        if aq is not None:
            df_chunk.at[df_chunk.index[idx], 'AQ_**Стоимость_закупки**'] = aq
            df_chunk.at[df_chunk.index[idx], 'AR_**Прямая_СС**'] = ar
            df_chunk.at[df_chunk.index[idx], 'AS_**НР**'] = as_val
            success_count += 1
        else:
            logger.debug(f"Строка {global_idx+2}: Нет совпадений")
            df_chunk.at[df_chunk.index[idx], 'AQ_**Стоимость_закупки**'] = "#РП"
            df_chunk.at[df_chunk.index[idx], 'AR_**Прямая_СС**'] = "#РП"
            df_chunk.at[df_chunk.index[idx], 'AS_**НР**'] = "#РП"
            error_count += 1

    return df_chunk, success_count, error_count

def main():
    parser = argparse.ArgumentParser(description='ExcelCostCalculator v2')
    parser.add_argument('--input', required=True, help='Входной Excel-файл')
    parser.add_argument('--output', required=True, help='Выходной Excel-файл')
    parser.add_argument('--log-level', default='INFO', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'])
    parser.add_argument('--chunk-size', type=int, default=DEFAULT_CHUNK_SIZE)
    parser.add_argument('--compare', action='store_true', help='Сравнить с эталонными данными')

    args = parser.parse_args()
    setup_logger(args.log_level)
    logger = get_logger(__name__)

    input_path = Path(args.input)
    if not input_path.exists():
        logger.error(f"Файл не найден: {args.input}")
        sys.exit(1)

    logger.info("========== ExcelCostCalculator v2 ==========")
    logger.info(f"Входной файл: {args.input}")
    logger.info(f"Выходной файл: {args.output}")
    logger.info(f"Размер chunk: {args.chunk_size} строк")
    if args.compare:
        logger.info("Режим сравнения: ВКЛЮЧЕН")

    start_time = datetime.now()
    initial_memory = get_memory_usage()
    if initial_memory > 0:
        logger.info(f"Начальное использование памяти: {initial_memory:.2f} МБ")

    try:
        logger.info("Загрузка Excel-файла...")
        df_target, df_source = load_excel_file(args.input, SHEET_NAMES['TARGET'], SHEET_NAMES['SOURCE'])
        logger.info(f"Целевая таблица: {len(df_target)} строк")
        logger.info(f"Справочная таблица: {len(df_source)} строк")

        after_load_memory = get_memory_usage()
        if after_load_memory > 0:
            logger.info(f"Память после загрузки: {after_load_memory:.2f} МБ")

        logger.info("Инициализация FormulaEngine...")
        engine = FormulaEngine(df_source, SOURCE_COLUMNS)

        df_target['AO_**Дата_приобретения**'] = None
        df_target['AP_**Квартал_приобретения**'] = ""
        df_target['AQ_**Стоимость_закупки**'] = None
        df_target['AR_**Прямая_СС**'] = None
        df_target['AS_**НР**'] = None

        logger.info("Начало обработки по частям...")
        total_rows = len(df_target)
        chunk_size = args.chunk_size
        total_success = 0
        total_errors = 0
        num_chunks = (total_rows + chunk_size - 1) // chunk_size
        logger.info(f"Общее количество chunks: {num_chunks}")

        for chunk_num in range(num_chunks):
            chunk_start = chunk_num * chunk_size
            chunk_end = min(chunk_start + chunk_size, total_rows)
            logger.info(f"Обработка chunk {chunk_num + 1}/{num_chunks} (строки {chunk_start + 2}-{chunk_end + 1})...")

            df_chunk = df_target.iloc[chunk_start:chunk_end].copy()
            df_chunk_processed, success_count, error_count = process_chunk(df_chunk, engine, chunk_start, logger)
            df_target.iloc[chunk_start:chunk_end] = df_chunk_processed

            total_success += success_count
            total_errors += error_count
            progress = chunk_end / total_rows * 100
            logger.info(f"Обработано: {chunk_end}/{total_rows} ({progress:.1f}%) | Успешно: {total_success} | Ошибок: {total_errors}")

            current_memory = get_memory_usage()
            if current_memory > 0:
                logger.debug(f"Память: {current_memory:.2f} МБ")

            if (chunk_num + 1) % GC_INTERVAL == 0:
                logger.debug("Очистка памяти...")
                gc.collect()

        logger.info("Финальная очистка памяти...")
        gc.collect()

        logger.info("Сохранение результатов...")
        write_results(args.input, args.output, df_target, SHEET_NAMES['TARGET'])

        if args.compare:
            logger.info("Сравнение с эталонными данными...")
            df_original = pd.read_excel(args.input, sheet_name=SHEET_NAMES['TARGET'])
            comparator = DataComparison(df_original)
            comparison_stats = comparator.compare_results(df_target)
            report = comparator.generate_report(comparison_stats)
            print("\n" + report)

            report_filename = args.output.replace('.xlsx', '_comparison_report.txt')
            with open(report_filename, 'w', encoding='utf-8') as f:
                f.write(report)
            logger.info(f"Отчёт сравнения сохранён: {report_filename}")

        elapsed = (datetime.now() - start_time).total_seconds()
        final_memory = get_memory_usage()

        logger.info("========== ИТОГОВАЯ СТАТИСТИКА ==========")
        logger.info(f"Всего обработано записей: {total_rows}")
        logger.info(f"Успешно сопоставлено: {total_success} ({total_success/total_rows*100:.1f}%)")
        logger.info(f"Требует ручной проверки: {total_errors} ({total_errors/total_rows*100:.1f}%)")
        logger.info(f"Время обработки: {elapsed:.1f} сек ({elapsed/60:.1f} мин)")

        if final_memory > 0:
            logger.info(f"Финальная память: {final_memory:.2f} МБ")

        logger.info("Обработка завершена успешно!")

    except Exception as e:
        logger.error(f"Ошибка при обработке: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()
