"""Запись результатов в Excel (ИСПРАВЛЕННАЯ ВЕРСИЯ)"""
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from utils.logger import get_logger

logger = get_logger(__name__)


def write_results(input_file: str, output_file: str, df_result: pd.DataFrame, target_sheet: str) -> None:
    """
    Записывает результаты расчетов в новый Excel-файл.

    Args:
        input_file: Путь к исходному Excel-файлу
        output_file: Путь к результирующему Excel-файлу
        df_result: DataFrame с результатами
        target_sheet: Имя целевого листа
    """
    try:
        # Загружаем исходный файл
        wb = load_workbook(input_file)
        ws = wb[target_sheet]

        logger.info("="*60)
        logger.info("Поиск строки с заголовками...")

        # Ищем строку с заголовками
        header_row = None
        for row in range(1, 21):
            cell_value = ws.cell(row=row, column=41).value
            if cell_value and "Дата приобретения" in str(cell_value):
                header_row = row
                logger.info(f"  ✅ Найдена строка: {row}")
                break

        if header_row is None:
            logger.warning("  ⚠️  Используем строку 1")
            header_row = 1

        # УЛУЧШЕНИЕ: Определяем начальный столбец для записи
        # Ищем первый пустой столбец после существующих данных
        start_col = ws.max_column + 1

        # Заголовки результирующих столбцов
        headers = [
            'Дата приобретения',
            'Квартал приобретения',
            'Стоимость закупки НЧТЗ 1 ед',
            'Прямая СС НЧТЗ 1 ед',
            'НР НЧТЗ 1 ед'
        ]

        # УЛУЧШЕНИЕ: Записываем заголовки с форматированием
        for i, header in enumerate(headers):
            cell = ws.cell(row=header_row, column=start_col + i, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')

        logger.info(f"  ✅ Заголовки записаны в столбцы {start_col}-{start_col + len(headers) - 1}")

        # Записываем данные
        data_start_row = header_row + 1

        logger.info(f"Запись данных начиная со строки {data_start_row}...")

        for row_idx in range(len(df_result)):
            excel_row = data_start_row + row_idx

            # Записываем значения из DataFrame
            ws.cell(row=excel_row, column=start_col, 
                   value=df_result.iloc[row_idx]['AO_Дата_приобретения'])
            ws.cell(row=excel_row, column=start_col + 1, 
                   value=df_result.iloc[row_idx]['AP_Квартал_приобретения'])
            ws.cell(row=excel_row, column=start_col + 2, 
                   value=df_result.iloc[row_idx]['AQ_Стоимость_закупки'])
            ws.cell(row=excel_row, column=start_col + 3, 
                   value=df_result.iloc[row_idx]['AR_Прямая_СС'])
            ws.cell(row=excel_row, column=start_col + 4, 
                   value=df_result.iloc[row_idx]['AS_НР'])

        logger.info(f"  ✅ Записано {len(df_result)} строк данных")

        # УЛУЧШЕНИЕ: Автоматическая настройка ширины столбцов
        for i in range(len(headers)):
            column_letter = ws.cell(row=header_row, column=start_col + i).column_letter
            ws.column_dimensions[column_letter].width = 20

        # Сохраняем файл
        wb.save(output_file)
        logger.info(f"  ✅ Файл сохранен: {output_file}")
        logger.info("="*60)

    except Exception as e:
        logger.error(f"❌ ОШИБКА при записи результатов: {e}")
        raise
