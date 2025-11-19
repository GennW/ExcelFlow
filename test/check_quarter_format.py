"""
Скрипт для проверки формата значений в столбце "Квартал приобретения"
Версия 2.1 - С пропуском первых 11 строк (начало с 12-й строки)
"""
import pandas as pd
import re
import numpy as np


def is_excel_not_empty(value):
    """
    Проверяет "непустоту" значения так же, как это делает фильтр Excel

    Excel считает ПУСТЫМИ:
    - NaN, None
    - Пустые строки ""
    - Строки только из пробелов "   "
    - Строку "nan" (как текст)
    - Тире "-"
    - Строку "none"
    """
    # NaN или None
    if pd.isna(value):
        return False

    # Приводим к строке и очищаем от пробелов
    str_value = str(value).strip().lower()

    # Список значений, которые считаются пустыми
    empty_values = ["", "nan", "none", "-", "—"]

    return str_value not in empty_values


def is_valid_quarter_format(value):
    """
    Проверяет что значение соответствует формату "N квартал YYYY"
    где N = 1-4, YYYY = 2020-2030

    Примеры валидных:
    - "1 квартал 2024"
    - "2 квартал 2023"
    - "4 квартал 2025"

    Примеры НЕвалидных:
    - "до 2024"
    - "1кв2024"
    - "Q1 2024"
    """
    if not is_excel_not_empty(value):
        return False

    str_value = str(value).strip()

    # Паттерн: "1-4 квартал 2020-2030"
    pattern = r'^([1-4])\s+квартал\s+(20[2-3][0-9])$'

    return bool(re.match(pattern, str_value))


def check_quarter_format(file_path, sheet_name=None, skip_rows=10):
    """
    Анализирует столбец "Квартал приобретения" в Excel файле

    Args:
        file_path: путь к Excel файлу
        sheet_name: имя листа (если None - первый лист)
        skip_rows: количество строк для пропуска (по умолчанию 11 - начинаем с 12-й)
    """
    print("="*80)
    print("АНАЛИЗ ФОРМАТА КВАРТАЛА")
    print("="*80)
    print(f"Файл: {file_path}")

    # Загружаем Excel файл с пропуском первых строк
    if sheet_name:
        df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=skip_rows)
        print(f"Лист: {sheet_name}")
    else:
        df = pd.read_excel(file_path, skiprows=skip_rows)
        print(f"Лист: (первый лист)")

    print(f"Пропущено первых строк: {skip_rows}")
    print(f"Загружено строк: {len(df)}")
    print(f"Данные начинаются с Excel строки: {skip_rows + 1}")
    print("")

    # Находим столбец "Квартал приобретения"
    target_col = None
    for col in df.columns:
        if 'Квартал приобретения' in str(col):
            target_col = col
            break

    if target_col is None:
        print("❌ ОШИБКА: Столбец 'Квартал приобретения' не найден!")
        print("\nДоступные столбцы:")
        for i, col in enumerate(df.columns, 1):
            print(f"  {i}. {col}")
        return

    print(f"✅ Найден столбец: '{target_col}'")
    print("="*80)
    print("")

    # Счетчики
    total_rows = len(df)
    excel_empty = 0
    excel_not_empty = 0
    valid_format = 0
    invalid_format = 0

    # Словари для группировки
    invalid_values = {}  # {значение: количество}

    # Анализируем каждую строку
    for idx, value in df[target_col].items():
        if is_excel_not_empty(value):
            excel_not_empty += 1

            if is_valid_quarter_format(value):
                valid_format += 1
            else:
                invalid_format += 1

                # Сохраняем невалидное значение
                str_val = str(value).strip()
                invalid_values[str_val] = invalid_values.get(str_val, 0) + 1
        else:
            excel_empty += 1

    # Выводим статистику
    print("📊 СТАТИСТИКА (как в Excel фильтре):")
    print("-"*80)
    print(f"  Всего строк (после пропуска): {total_rows:>6}")
    print(f"  Пустых (Excel):               {excel_empty:>6} ({excel_empty/total_rows*100:.1f}%)")
    print(f"  Непустых (Excel):             {excel_not_empty:>6} ({excel_not_empty/total_rows*100:.1f}%)")
    print("")
    print(f"✅ Валидный формат:             {valid_format:>6} ({valid_format/total_rows*100:.1f}%)")
    print(f"   (N квартал YYYY)")
    print("")
    print(f"❌ НЕвалидный формат:           {invalid_format:>6} ({invalid_format/total_rows*100:.1f}%)")
    print(f"   (до 2024, нет данных, etc)")
    print("")
    print("="*80)

    # Показываем примеры невалидных значений
    if invalid_values:
        print("")
        print("🔍 НЕВАЛИДНЫЕ ЗНАЧЕНИЯ (топ-20):")
        print("-"*80)

        # Сортируем по количеству (самые частые сначала)
        sorted_invalid = sorted(invalid_values.items(), key=lambda x: x[1], reverse=True)

        for i, (val, count) in enumerate(sorted_invalid[:20], 1):
            print(f"  {i:2}. '{val}' (встречается {count} раз)")

        if len(sorted_invalid) > 20:
            print(f"  ... и ещё {len(sorted_invalid) - 20} уникальных значений")

        print("")
        print(f"Всего уникальных невалидных значений: {len(invalid_values)}")

    # Диагностика (первые 15 строк)
    print("")
    print("="*80)
    print("🔬 ДИАГНОСТИКА (первые 15 строк данных):")
    print("-"*80)

    for idx in range(min(15, len(df))):
        val = df[target_col].iloc[idx]
        excel_row = skip_rows + 2 + idx  # Номер строки в Excel

        if is_excel_not_empty(val):
            status = "✅" if is_valid_quarter_format(val) else "❌"
        else:
            status = "⚪"

        print(f"  {status} Excel строка {excel_row:4}: {repr(val):<40} (type: {type(val).__name__})")

    print("="*80)

    # Итоговая сводка
    print("")
    print("📋 ВЫВОДЫ:")
    print("-"*80)

    if excel_not_empty == 0:
        print("  ⚠️  Все значения пустые!")
    else:
        valid_percent = (valid_format / excel_not_empty * 100) if excel_not_empty > 0 else 0

        if valid_percent > 90:
            print(f"  ✅ Отлично! {valid_percent:.1f}% значений в правильном формате")
        elif valid_percent > 70:
            print(f"  ⚠️  Хорошо, но есть проблемы: {valid_percent:.1f}% в правильном формате")
        else:
            print(f"  ❌ Много ошибок! Только {valid_percent:.1f}% в правильном формате")

    print("="*80)


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        # По умолчанию
        file_path = "результат.xlsx"

    # Можно передать количество строк для пропуска вторым аргументом
    skip_rows = 10  # По умолчанию пропускаем первые 11 строк (начинаем с 12-й)
    if len(sys.argv) > 2:
        try:
            skip_rows = int(sys.argv[2])
        except ValueError:
            print(f"⚠️  ПРЕДУПРЕЖДЕНИЕ: Неверное значение skip_rows '{sys.argv[2]}', используется 11")

    try:
        check_quarter_format(file_path, skip_rows=skip_rows)
    except FileNotFoundError:
        print(f"❌ ОШИБКА: Файл '{file_path}' не найден!")
        print("\nИспользование:")
        print(f"  python {sys.argv[0]} <путь_к_файлу.xlsx> [количество_строк_для_пропуска]")
        print("\nПримеры:")
        print(f"  python {sys.argv[0]} результат.xlsx")
        print(f"  python {sys.argv[0]} результат.xlsx 11")
    except Exception as e:
        print(f"❌ ОШИБКА: {e}")
        import traceback
        traceback.print_exc()