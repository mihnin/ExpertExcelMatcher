"""Проверка созданных тестовых данных"""
import pandas as pd

# Загрузка данных
df1 = pd.read_excel('Тестовые_данные_Источник1.xlsx')
df2 = pd.read_excel('Тестовые_данные_Источник2.xlsx')

print("="*70)
print("ПРОВЕРКА ТЕСТОВЫХ ДАННЫХ")
print("="*70)

print(f"\n[*] Источник 1: {len(df1)} записей")
print(f"[*] Источник 2: {len(df2)} записей")
print(f"[*] Столбцы: {list(df1.columns)}")

print("\n" + "="*70)
print("ИЗМЕНЕННЫЕ ЗАПИСИ (индексы: 12, 19, 32, 38, 40)")
print("="*70)

for idx in [12, 19, 32, 38, 40]:
    print(f"\n--- Запись {idx+1} ---")
    print(f"Источник 1:")
    print(f"  Номер док: {df1.iloc[idx]['Номер документа']}")
    print(f"  Покупатель: {df1.iloc[idx]['Покупатель']}")
    print(f"Источник 2:")
    print(f"  Номер док: {df2.iloc[idx]['Номер документа']}")
    print(f"  Покупатель: {df2.iloc[idx]['Покупатель']}")

    if df1.iloc[idx]['Номер документа'] != df2.iloc[idx]['Номер документа']:
        print("  => НОМЕР ДОКУМЕНТА ОТЛИЧАЕТСЯ!")
    if df1.iloc[idx]['Покупатель'] != df2.iloc[idx]['Покупатель']:
        print("  => ПОКУПАТЕЛЬ ОТЛИЧАЕТСЯ!")

print("\n" + "="*70)
print("СОВПАДАЮЩИЕ ЗАПИСИ (примеры)")
print("="*70)

matching_count = 0
for idx in range(len(df1)):
    if (df1.iloc[idx]['Номер документа'] == df2.iloc[idx]['Номер документа'] and
        df1.iloc[idx]['Покупатель'] == df2.iloc[idx]['Покупатель']):
        matching_count += 1
        if matching_count <= 3:
            print(f"\nЗапись {idx+1}: {df1.iloc[idx]['Номер документа']} | {df1.iloc[idx]['Покупатель']}")

print(f"\n[*] Всего совпадающих записей: {matching_count}")
print(f"[*] Всего отличающихся записей: {len(df1) - matching_count}")

print("\n" + "="*70)
print("[OK] ПРОВЕРКА ЗАВЕРШЕНА")
print("="*70)
