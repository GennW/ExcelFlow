"""Скрипт для расширенной проверки данных в файле Excel"""
import pandas as pd

def debug_extended(file_path, sheet_name):
    """Проверяет содержимое файла с AO-AS более подробно"""
    # Загружаем файл с header=None, чтобы получить реальные данные
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    print(f"Файл: {file_path}")
    print(f"Лист: {sheet_name}")
    print(f"Размер таблицы: {df.shape}")
    print("\nДанные в столбцах AO-AS (40-44) для первых 10 строк:")
    
    for row_idx in range(0, 10):
        print(f"Строка {row_idx}: ", end="")
        for col_idx in range(40, 45):
            if col_idx < len(df.columns):
                value = df.iloc[row_idx, col_idx]
                col_letter = idx_to_excel_col(col_idx)
                print(f"{col_letter}:'{value}' ", end="")
            else:
                col_letter = idx_to_excel_col(col_idx)
                print(f"{col_letter}:[ИНДЕКС ВНЕ ДИАПАЗОНА] ", end="")
        print()
    
    print(f"\nПроверим строки с 10 по 20:")
    for row_idx in range(10, 21):
        print(f"Строка {row_idx}: ", end="")
        for col_idx in range(40, 45):
            if col_idx < len(df.columns):
                value = df.iloc[row_idx, col_idx]
                col_letter = idx_to_excel_col(col_idx)
                print(f"{col_letter}:'{value}' ", end="")
            else:
                col_letter = idx_to_excel_col(col_idx)
                print(f"{col_letter}:[ИНДЕКС ВНЕ ДИАПАЗОНА] ", end="")
        print()

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
    
    debug_extended(file_path, sheet_name)