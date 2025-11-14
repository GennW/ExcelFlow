# ExcelCostCalculator v2

Автоматический расчёт средневзвешенной себестоимости с функцией сравнения.

## Новое в версии 2

✅ Функция сравнения с эталонными данными (--compare)
✅ Детальный отчёт о несовпадениях
✅ Умная проверка наличия эталонных столбцов

## Установка

```bash
pip install -r requirements.txt
```

## Использование

Базовый режим:
```bash
python main.py --input "file.xlsx" --output "result.xlsx"
```

С сравнением:
```bash
python main.py --input "file.xlsx" --output "result.xlsx" --compare

python3 main.py --input "Материалы/СК_ТПХпродажи_1_пг_2025.xlsx" --output "результат.xlsx" --compare
```

## Параметры

- `--input` - входной файл
- `--output` - выходной файл
- `--log-level` - уровень логирования (DEBUG/INFO/WARNING/ERROR)
- `--chunk-size` - размер chunk (по умолчанию 500)
- `--compare` - включить сравнение с эталонными данными
