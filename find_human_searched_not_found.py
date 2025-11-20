"""Скрипт для поиска строк, где человек искал данные, но не нашёл (45 случаев)"""
import pandas as pd
import re

def is_valid_quarter_format(value):
    """Проверяет что квартал имеет валидный формат 'N квартал YYYY' (N=1-4, YYYY=2020-2030)"""
    if pd.isna(value):
        return False
    
    str_value = str(value).strip()
    
    # Паттерн: "1-4 квартал 2020-2030"
    pattern = r'^([1-4])\s+квартал\s+(20[2-3][0-9])$'
    
    return bool(re.match(pattern, str_value))

def is_formula_or_zero(value):
    """Проверяет, является ли значение формулой или нулем"""
    if pd.isna(value):
        return False
    
    str_value = str(value).strip()
    is_formula = str_value.startswith('=') 
    is_zero = str_value == '0'
    
    return is_formula or is_zero

def find_human_searched_not_found():
    """Находит строки, где человек искал данные, но не нашёл"""
    
    # Загружаем результаты сравнения
    try:
        df_result = pd.read_excel('результат.xlsx', sheet_name='СК ТПХ_1 пг', skiprows=10)
        print("Файл: результат.xlsx")
    except:
        df_result = pd.read_excel('результат_обновленный.xlsx', sheet_name='СК ТПХ_1 пг', skiprows=10)
        print("Файл: результат_обновленный.xlsx")
    
    print("Поиск строк, где человек искал данные, но не нашёл (45 случаев)")
    print("="*80)
    
    # Столбцы в файле
    ao_col = 'Дата приобретения'  # AO
    ap_col = 'Квартал приобретения'  # AP
    aq_col = 'Стоимость закупки НЧТЗ 1 ед'  # AQ
    
    # В этом файле расчетные значения могут быть в тех же столбцах, что и эталонные
    # или в специальных столбцах типа #РП, #НД и т.д.
    # Проверим последние столбцы на наличие расчетных данных
    print("Проверяем последние столбцы на наличие расчетных данных...")
    last_cols = df_result.columns[-10:]  # Последние 10 столбцов
    for i, col in enumerate(last_cols):
        print(f"  {len(df_result.columns)-10+i}: {col}")
    
    # В текущей реализации системы сравнения расчетные значения могут быть в тех же столбцах
    # но в других строках, или могут быть в специальных столбцах с маркерами
    # Попробуем найти столбцы, содержащие расчетные значения
    calc_aq_col = None
    for col in df_result.columns:
        if 'Стоимость' in str(col) and 'закупки' in str(col) and 'НЧТЗ' in str(col):
            if col != aq_col:  # Не эталонный столбец
                calc_aq_col = col
                break
    
    # Если не найден отдельный столбец, будем использовать тот же столбец,
    # но интерпретировать значения по-другому
    if calc_aq_col is None:
        calc_aq_col = aq_col
        print(f"Используем тот же столбец для эталона и расчета: {calc_aq_col}")
    
    print(f"Используем столбцы: эталон - {aq_col}, расчет - {calc_aq_col}")
    
    matches = []
    
    matches = []
    
    for idx, row in df_result.iterrows():
        # Проверяем, есть ли валидный квартал
        quarter_value = row[ap_col]
        if is_valid_quarter_format(quarter_value):
            # Проверяем, что человек искал (AQ - формула или 0)
            aq_ref = row[aq_col]
            if is_formula_or_zero(aq_ref):
                # Проверяем, что программа не нашла (расчетное значение - #РП или пусто)
                # В текущей реализации расчетные значения могут быть в том же столбце,
                # но если это так, то нужно понять, как они представлены
                aq_calc = row[calc_aq_col]
                # Проверим, является ли значение в этом столбце результатом расчета
                # Если в столбце AQ есть специальные значения #РП, #НД, то это расчетные
                calc_is_missing = pd.isna(aq_calc) or str(aq_calc) == '#РП' or str(aq_calc) == '#НД' or str(aq_calc).strip() == ''
                
                if calc_is_missing:
                    matches.append((idx, row))
        
        print(f"Найдено {len(matches)} строк, соответствующих критериям:")
        print()
        
        for i, (idx, row) in enumerate(matches[:20], 1):  # Показываем первые 20
            print(f"{i:2d}. Строка {idx + 12} (Excel):")
            print(f"    AO (Дата приобретения): {row[ao_col]}")
            print(f"    AP (Квартал приобретения): {row[ap_col]}")
            print(f"    AQ (Стоимость эталон): {row[aq_col]}")
            print(f"    AQ (Стоимость расчет): {row[calc_aq_col]}")
            print()
        
        if len(matches) > 20:
            print(f"... и еще {len(matches) - 20} строк")

if __name__ == "__main__":
    find_human_searched_not_found()