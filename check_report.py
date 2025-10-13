import pandas as pd

# Загружаем данные
askupo_df = pd.read_excel('Уникальные_ПО_продукты.xlsx')
report_df = pd.read_excel('Полный_отчет_сопоставления3.xlsx')

print('='*60)
print('ПРОВЕРКА ПОЛНОГО ОТЧЕТА')
print('='*60)

print(f'\nЗаписей в АСКУПО: {len(askupo_df)}')
print(f'Записей в отчете: {len(report_df)}')

if len(askupo_df) != len(report_df):
    print(f'\n!!! ОШИБКА: Не все записи попали в отчет!')
    print(f'Пропущено записей: {len(askupo_df) - len(report_df)}')
else:
    print('\nOK: Все записи из АСКУПО есть в отчете')

# Статистика по процентам
print('\n' + '='*60)
print('СТАТИСТИКА ПО ПРОЦЕНТАМ')
print('='*60)

zeros = report_df[report_df['Процент'] == 0.0]
low = report_df[(report_df['Процент'] > 0) & (report_df['Процент'] < 50)]
medium_low = report_df[(report_df['Процент'] >= 50) & (report_df['Процент'] < 70)]
medium = report_df[(report_df['Процент'] >= 70) & (report_df['Процент'] < 90)]
high = report_df[(report_df['Процент'] >= 90) & (report_df['Процент'] < 100)]
perfect = report_df[report_df['Процент'] == 100]

print(f'\n0% (нет совпадений):      {len(zeros)} записей')
print(f'1-49% (очень низкое):    {len(low)} записей')
print(f'50-69% (низкое):         {len(medium_low)} записей')
print(f'70-89% (среднее):        {len(medium)} записей')
print(f'90-99% (высокое):        {len(high)} записей')
print(f'100% (точное):           {len(perfect)} записей')
print(f'\nВсего: {len(zeros) + len(low) + len(medium_low) + len(medium) + len(high) + len(perfect)}')

# Примеры записей с 0%
if len(zeros) > 0:
    print('\n' + '='*60)
    print('ПРИМЕРЫ ЗАПИСЕЙ С 0% СОВПАДЕНИЕМ')
    print('='*60)
    for idx, row in zeros.head(10).iterrows():
        print(f"\nАСКУПО: '{row['АСКУПО']}'")
        print(f"EA Tool: '{row['EA Tool']}'")
        print(f"Процент: {row['Процент']}%")
else:
    print('\nЗаписей с 0% не найдено')

# Поиск пропущенных записей
print('\n' + '='*60)
print('ПРОВЕРКА ПРОПУЩЕННЫХ ЗАПИСЕЙ')
print('='*60)

askupo_names = set(askupo_df.iloc[:, 0].astype(str).tolist())
report_names = set(report_df['АСКУПО'].astype(str).tolist())

missing = askupo_names - report_names
if missing:
    print(f'\nПропущено {len(missing)} записей:')
    for name in list(missing)[:10]:
        print(f"  - '{name}'")
else:
    print('\nОтлично! Все записи присутствуют в отчете')
