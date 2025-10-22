# -*- coding: utf-8 -*-
import pandas as pd
import sys
import io

# Установить UTF-8 для вывода
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Читаем файл
file_path = r'c:\dev\ExpertExcelMatcher\Полное_сравнение_всех_методов1.xlsx'
xls = pd.ExcelFile(file_path, engine='openpyxl')

print(f"Всего листов: {len(xls.sheet_names)}\n")
print("Список листов:")
for i, name in enumerate(xls.sheet_names, 1):
    print(f"{i}. {name}")

# Читаем первый лист для примера
print("\n--- Структура первого листа ---")
df_first = pd.read_excel(xls, sheet_name=0, nrows=3)
print(f"Колонки: {df_first.columns.tolist()}")
print(f"\nПервые 3 строки:")
print(df_first)

# Проверяем, есть ли лист статистики
if 'Статистика' in xls.sheet_names or 'Statistics' in xls.sheet_names:
    stat_sheet = 'Статистика' if 'Статистика' in xls.sheet_names else 'Statistics'
    df_stat = pd.read_excel(xls, sheet_name=stat_sheet)
    print(f"\n--- Лист {stat_sheet} ---")
    print(df_stat)
