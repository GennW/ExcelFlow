"""Скрипт для отладки названий столбцов в загруженном датафрейме"""
import pandas as pd

def debug_loaded_columns(file_path, sheet_name):
    """Проверяет, как pandas загружает столбцы из файла"""
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    print(f"Файл: {file_path}")
    print(f"Лист: {sheet_name}")
    print(f"Всего столбцов: {len(df.columns)}")
    print("\nНазвания столбцов с индексами AO-AS (40-44):")
    
    # Выведем названия столбцов AO-AS
    for idx in range(40, 45):
        if idx < len(df.columns):
            col_name = df.columns[idx]
            col_letter = idx_to_excel_col(idx)
            print(f" {col_letter} ({idx}): '{col_name}'")
        else:
            col_letter = idx_to_excel_col(idx)
            print(f"  {col_letter} ({idx}): [ИНДЕКС ВНЕ ДИАПАЗОНА]")
    
    print(f"\nПервые 3 строки данных для этих столбцов:")
    print(df.iloc[0:3, 40:45])  # Показываем первую и вторую строки

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
    
    debug_loaded_columns(file_path, sheet_name)