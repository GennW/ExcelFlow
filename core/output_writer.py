"""Запись результатов"""
import pandas as pd
from openpyxl import load_workbook
from utils.logger import get_logger

logger = get_logger(__name__)

def write_results(input_file: str, output_file: str, df_result: pd.DataFrame, target_sheet: str) -> None:
    try:
        wb = load_workbook(input_file)
        ws = wb[target_sheet]
        logger.info("="*60)
        logger.info("Поиск строки с заголовками...")
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
        start_col = ws.max_column + 1
        headers = ['**Дата приобретения**', '**Квартал приобретения**', '**Стоимость закупки НЧТЗ 1 ед**', '**Прямая СС НЧТЗ 1 ед**', '**НР НЧТЗ 1 ед**']
        for i, header in enumerate(headers):
            ws.cell(row=header_row, column=start_col + i, value=header)
        data_start_row = header_row + 1
        for row_idx in range(len(df_result)):
            excel_row = data_start_row + row_idx
            ws.cell(row=excel_row, column=start_col, value=df_result.iloc[row_idx]['AO_Дата_приобретения'])
            ws.cell(row=excel_row, column=start_col + 1, value=df_result.iloc[row_idx]['AP_Квартал_приобретения'])
            ws.cell(row=excel_row, column=start_col + 2, value=df_result.iloc[row_idx]['AQ_Стоимость_закупки'])
            ws.cell(row=excel_row, column=start_col + 3, value=df_result.iloc[row_idx]['AR_Прямая_СС'])
            ws.cell(row=excel_row, column=start_col + 4, value=df_result.iloc[row_idx]['AS_НР'])
        logger.info(f"  ✅ Записано {len(df_result)} строк")
        wb.save(output_file)
        logger.info(f"  ✅ Сохранено: {output_file}")
        logger.info("="*60)
    except Exception as e:
        logger.error(f"❌ ОШИБКА: {e}")
        raise
