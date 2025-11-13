import logging
import sys
from datetime import datetime
from typing import Optional

def setup_logger(log_level: str = "INFO", log_file: Optional[str] = None) -> logging.Logger:
    """
    Настраивает и возвращает логгер с заданным уровнем и файлом вывода.
    
    :param log_level: Уровень логирования (DEBUG, INFO, WARNING, ERROR)
    :param log_file: Путь к файлу лога (опционально)
    :return: Настроенный объект логгера
    """
    # Преобразуем строку уровня логирования в соответствующий уровень
    numeric_level = getattr(logging, log_level.upper(), None)
    if not isinstance(numeric_level, int):
        raise ValueError(f'Invalid log level: {log_level}')
    
    # Создаем основной логгер
    logger = logging.getLogger('ExcelFlow')
    logger.setLevel(numeric_level)
    
    # Удаляем существующие обработчики, чтобы избежать дублирования
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # Форматтер для логов
    formatter = logging.Formatter(
        '[%(asctime)s] %(levelname)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Обработчик для консоли
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(numeric_level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # Обработчик для файла (если указан)
    if log_file:
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(numeric_level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    return logger


def log_application_start(logger: logging.Logger, version: str = "1.0", 
                         input_file: str = None, output_file: str = None):
    """
    Логирует начало работы приложения.
    
    :param logger: Объект логгера
    :param version: Версия приложения
    :param input_file: Путь к входному файлу
    :param output_file: Путь к выходному файлу
    """
    logger.info(f"============ ExcelFlow v{version} ============")
    if input_file:
        logger.info(f"Входной файл: {input_file}")
    if output_file:
        logger.info(f"Выходной файл: {output_file}")
    logger.info(f"Время запуска: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


def log_sheet_info(logger: logging.Logger, sheet_names: list, sheet_counts: dict):
    """
    Логирует информацию о вкладках файла.
    
    :param logger: Объект логгера
    :param sheet_names: Список имён вкладок
    :param sheet_counts: Словарь с количеством строк в каждой вкладке
    """
    logger.info(f"Найдено вкладок: {len(sheet_names)}")
    for sheet_name in sheet_names:
        logger.info(f"Вкладка '{sheet_name}': {sheet_counts.get(sheet_name, 0)} строк")


def log_processing_progress(logger: logging.Logger, processed: int, total: int):
    """
    Логирует прогресс обработки (каждые 25%).
    
    :param logger: Объект логгера
    :param processed: Количество обработанных записей
    :param total: Общее количество записей
    """
    if total == 0:
        return
    
    percent = (processed / total) * 10
    
    # Логируем на каждом 25% пороге
    if processed == 0 or percent >= 25 or processed == total:
        # Проверяем, нужно ли логировать этот процент
        if (processed == 0 or 
            (percent >= 25 and percent < 50 and processed > 0) or
            (percent >= 50 and percent < 75) or
            (percent >= 75 and percent < 100) or
            processed == total):
            
            logger.info(f"Обработано: {percent:.1f}% ({processed} из {total})")


def log_final_stats(logger: logging.Logger, total_records: int, matched_records: int, 
                   manual_check_records: int, missing_data_records: int,
                   exact_matches: int, level1_matches: int, level2_matches: int, 
                   fuzzy_matches: int, processing_time: float):
    """
    Логирует итоговую статистику обработки.
    
    :param logger: Объект логгера
    :param total_records: Всего обработано записей
    :param matched_records: Успешно сопоставлено
    :param manual_check_records: Требует ручной проверки
    :param missing_data_records: Отсутствуют ключевые данные
    :param exact_matches: Точное совпадение
    :param level1_matches: Совпадение Уровень 1
    :param level2_matches: Совпадение Уровень 2
    :param fuzzy_matches: Размытое совпадение
    :param processing_time: Время обработки в секундах
    """
    logger.info("========== ИТОГОВАЯ СТАТИСТИКА ==========")
    
    if total_records > 0:
        matched_percent = (matched_records / total_records) * 10
        manual_check_percent = (manual_check_records / total_records) * 10
        missing_data_percent = (missing_data_records / total_records) * 100
        
        logger.info(f"Всего обработано записей: {total_records}")
        logger.info(f"Успешно сопоставлено: {matched_records} ({matched_percent:.1f}%)")
        logger.info(f"Требует ручной проверки: {manual_check_records} ({manual_check_percent:.1f}%)")
        logger.info(f"Отсутствуют ключевые данные: {missing_data_records} ({missing_data_percent:.1f}%)")
        
        logger.info(f"Точное совпадение: {exact_matches}")
        logger.info(f"Совпадение Уровень 1: {level1_matches}")
        logger.info(f"Совпадение Уровень 2: {level2_matches}")
        logger.info(f"Размытое совпадение: {fuzzy_matches}")
    
    # Преобразуем время в формат HH:MM:SS
    hours = int(processing_time // 3600)
    minutes = int((processing_time % 3600) // 60)
    seconds = int(processing_time % 60)
    
    logger.info(f"Время обработки: {hours:02d}:{minutes:02d}:{seconds:02d}")