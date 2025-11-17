"""Запись результатов в Excel"""
import pandas as pd
from openpyxl import load_workbook
from utils.logger import get_logger

logger = get_logger(__name__)

def write_results(input_file: str, output_file: str, df_result: pd.DataFrame, target_sheet: str) -> None:
    try:
        wb = load_workbook(input_file)
        ws = wb[target_sheet]
        logger.info(f"Открыт файл для записи: {input_file}")

        start_col = ws.max_column + 1
        headers = ['**Дата приобретения**', '**Квартал приобретения**', '**Стоимость закупки НЧТЗ 1 ед**',
                  '**Прямая СС НЧТЗ 1 ед**', '**НР НЧТЗ 1 ед**']

        for i, header in enumerate(headers):
            ws.cell(row=1, column=start_col + i, value=header)

        logger.info(f"Добавлены заголовки в столбцы {start_col}-{start_col + 4}")

        for row_idx in range(len(df_result)):
            excel_row = row_idx + 2
            ws.cell(row=excel_row, column=start_col, value=df_result.iloc[row_idx]['AO_**Дата_приобретения**'])
            ws.cell(row=excel_row, column=start_col + 1, value=df_result.iloc[row_idx]['AP_**Квартал_приобретения**'])
            ws.cell(row=excel_row, column=start_col + 2, value=df_result.iloc[row_idx]['AQ_**Стоимость_закупки**'])
            ws.cell(row=excel_row, column=start_col + 3, value=df_result.iloc[row_idx]['AR_**Прямая_СС**'])
            ws.cell(row=excel_row, column=start_col + 4, value=df_result.iloc[row_idx]['AS_**НР**'])

        logger.info(f"Записано {len(df_result)} строк данных")
        wb.save(output_file)
        logger.info(f"Файл успешно сохранён: {output_file}")
    except Exception as e:
        logger.error(f"Ошибка при записи: {e}")
        raise
