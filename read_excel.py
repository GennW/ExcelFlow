import pandas as pd

# Считать все листы в виде словаря {имя_листа: DataFrame}
df_dict = pd.read_excel('СК ТПХ продажи_1 пг 2025.xlsx', sheet_name=None)

# Вывести имена всех листов
print(df_dict.keys())

# Пример: вывести первые строки каждого листа
for sheet, df in df_dict.items():
    print(f"=== {sheet} ===")
    print(df.head())
