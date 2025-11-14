"""Логирование (БЕЗ ИЗМЕНЕНИЙ)"""
import logging
import sys
from config import LOG_FORMAT, LOG_DATE_FORMAT


def setup_logger(level: str = 'INFO') -> None:
    """
    Настраивает корневой логгер.

    Args:
        level: Уровень логирования ('DEBUG', 'INFO', 'WARNING', 'ERROR')
    """
    log_level = getattr(logging, level.upper(), logging.INFO)

    # Создаем форматтер
    formatter = logging.Formatter(LOG_FORMAT, LOG_DATE_FORMAT)

    # Создаем обработчик для вывода в консоль
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)

    # Настраиваем корневой логгер
    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)
    root_logger.handlers = []  # Очищаем старые обработчики
    root_logger.addHandler(console_handler)


def get_logger(name: str) -> logging.Logger:
    """
    Получает логгер для модуля.

    Args:
        name: Имя модуля (обычно __name__)

    Returns:
        Настроенный логгер
    """
    return logging.getLogger(name)
