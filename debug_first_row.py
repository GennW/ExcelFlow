"""Скрипт для проверки содержимого первой строки файла Excel"""
import pandas as pd

def debug_first_row(file_path, sheet_name):
    """Проверяет содержимое первой строки файла"""
    # Загружаем файл с header=None, чтобы получить реальные данные из первой строки
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    print(f"Файл: {file_path}")
    print(f"Лист: {sheet_name}")
    print(f"Размер таблицы: {df.shape}")
    print("\nПервая строка (индекс 0) для столбцов AO-AS (индексы 40-44):")
    
    for idx in range(40, 45):
        if idx < len(df.columns):
            value = df.iloc[0, idx]
            col_letter = idx_to_excel_col(idx)
            print(f" {col_letter} ({idx}): '{value}' (тип: {type(value).__name__})")
        else:
            col_letter = idx_to_excel_col(idx)
            print(f"  {col_letter} ({idx}): [ИНДЕКС ВНЕ ДИАПАЗОНА]")
    
    print(f"\nПервые 5 строк для этих столбцов:")
    print(df.iloc[0:5, 40:45])

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
    
    debug_first_row(file_path, sheet_name)