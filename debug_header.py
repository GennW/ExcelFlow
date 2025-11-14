"""Скрипт для проверки заголовков при загрузке файла Excel"""
import pandas as pd

def debug_header_loading(file_path, sheet_name):
    """Проверяет, как pandas загружает файл с header=0 (по умолчанию)"""
    # Загружаем файл так, как это делает основная программа
    df_with_header = pd.read_excel(file_path, sheet_name=sheet_name)
    
    print(f"Файл: {file_path}")
    print(f"Лист: {sheet_name}")
    print(f"Размер загруженного датафрейма: {df_with_header.shape}")
    print(f"Названия столбцов (первые 5 и последние 5):")
    print(f"  Первые 5: {list(df_with_header.columns[:5])}")
    print(f"  Последние 5: {list(df_with_header.columns[-5:])}")
    
    print(f"\nНазвания столбцов AO-AS (40-44):")
    for idx in range(40, 45):
        if idx < len(df_with_header.columns):
            col_name = df_with_header.columns[idx]
            col_letter = idx_to_excel_col(idx)
            print(f" {col_letter} ({idx}): '{col_name}'")
    
    print(f"\nПервые 3 строки данных для этих столбцов:")
    print(df_with_header.iloc[0:3, 40:45])
    
    print(f"\nДанные в строках 8-12 для этих столбцов:")
    print(df_with_header.iloc[8:13, 40:45])

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
    
    debug_header_loading(file_path, sheet_name)