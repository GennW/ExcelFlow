"""Скрипт для проверки данных в столбцах AO-AS в Excel-файле"""
import pandas as pd
from pathlib import Path

def check_column_data(file_path, sheet_name):
    """Проверяет данные в столбцах AO-AS в указанном листе Excel-файла"""
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    print(f"Файл: {file_path}")
    print(f"Лист: {sheet_name}")
    print(f"Всего строк: {len(df)}")
    print("\nДанные в столбцах AO-AS (индексы 40-44):")
    
    # Выведем названия столбцов и первые несколько значений
    for idx in range(40, 45):
        if idx < len(df.columns):
            col_name = df.columns[idx]
            col_letter = idx_to_excel_col(idx)
            print(f"\n{col_letter} ({idx}): '{col_name}'")
            
            # Выведем первые 5 непустых значений
            col_data = df.iloc[:, idx]
            non_empty_values = col_data.dropna().head(5)
            if len(non_empty_values) > 0:
                print(f"  Непустые значения: {list(non_empty_values.values)}")
            else:
                print(f"  Все значения пустые")
        else:
            col_letter = idx_to_excel_col(idx)
            print(f"  {col_letter} ({idx}): [ИНДЕКС ВНЕ ДИАПАЗОНА]")

def idx_to_excel_col(idx):
    """Преобразует индекс столбца в буквенное обозначение Excel (0 -> A, 1 -> B, ..., 40 -> AO)"""
    result = ""
    idx += 1  # Сдвигаемся к 1-индексации
    
    while idx > 0:
        idx -= 1
        result = chr(idx % 26 + ord('A')) + result
        idx //= 26
    
    return result

if __name__ == "__main__":
    file_path = "Материалы/СК_ТПХпродажи_1_пг_2025.xlsx"
    sheet_name = "СК ТПХ_1 пг"
    
    check_column_data(file_path, sheet_name)