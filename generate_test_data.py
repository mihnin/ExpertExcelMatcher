"""
Генератор тестовых данных для Expert Excel Matcher v2.0
Создает 2 источника данных по 50 записей с небольшими различиями
"""

import pandas as pd
import random
from datetime import datetime, timedelta

# Настройка random seed для воспроизводимости
random.seed(42)

# Списки для генерации данных
PRODUCTS = [
    "Ноутбук Lenovo ThinkPad X1",
    "Монитор Dell 27 дюймов",
    "Клавиатура Logitech MX Keys",
    "Мышь Logitech MX Master 3",
    "Принтер HP LaserJet",
    "Сканер Canon CanoScan",
    "Веб-камера Logitech C920",
    "Наушники Sony WH-1000XM4",
    "Роутер TP-Link AC1750",
    "Коммутатор Cisco SG250",
    "Жесткий диск Seagate 2TB",
    "SSD Samsung 1TB",
    "ОЗУ Kingston 16GB",
    "Видеокарта NVIDIA RTX 3060",
    "Процессор Intel Core i7",
    "Материнская плата ASUS",
    "Блок питания Corsair 750W",
    "Корпус Fractal Design",
    "Кулер Noctua NH-D15",
    "USB флешка SanDisk 128GB"
]

CUSTOMERS = [
    "ООО Альфа-Системс",
    "ЗАО БетаТех",
    "ИП Васильев А.П.",
    "ООО ГаммаПро",
    "АО ДельтаСофт",
    "ООО Эпсилон",
    "ИП Жданов М.С.",
    "ООО ЗетаГрупп",
    "АО ЭтаТрейд",
    "ООО ТетаЛаб",
    "ИП Иванов И.И.",
    "ООО КаппаТех",
    "ЗАО ЛямбдаСофт",
    "ООО МюСистемс",
    "АО НюКорп"
]

def generate_source_data(source_num, num_records=50):
    """Генерация данных источника"""
    data = []

    for i in range(num_records):
        doc_num = f"DOC-2024-{1000 + i:04d}"
        product = random.choice(PRODUCTS)
        price = round(random.uniform(5000, 150000), 2)
        customer = random.choice(CUSTOMERS)

        data.append({
            'Номер документа': doc_num,
            'Наименование товара': product,
            'Цена (руб)': price,
            'Покупатель': customer
        })

    return data

# Генерация источника 1
print("Генерация Источника данных 1...")
source1_data = generate_source_data(1, 50)
df_source1 = pd.DataFrame(source1_data)

# Генерация источника 2 (копия источника 1)
print("Генерация Источника данных 2...")
source2_data = source1_data.copy()
df_source2 = pd.DataFrame(source2_data)

# Внесение изменений в 5 записей источника 2
print("\nВнесение различий в 5 записей...")
indices_to_modify = random.sample(range(50), 5)

for idx in indices_to_modify:
    # Изменяем номер документа (меняем год или номер)
    original_doc = df_source2.at[idx, 'Номер документа']
    if random.choice([True, False]):
        # Меняем год
        df_source2.at[idx, 'Номер документа'] = original_doc.replace('2024', '2023')
    else:
        # Меняем номер
        old_num = int(original_doc.split('-')[-1])
        new_num = old_num + random.randint(1, 100)
        df_source2.at[idx, 'Номер документа'] = f"DOC-2024-{new_num:04d}"

    # Изменяем покупателя (выбираем другого)
    current_customer = df_source2.at[idx, 'Покупатель']
    other_customers = [c for c in CUSTOMERS if c != current_customer]
    df_source2.at[idx, 'Покупатель'] = random.choice(other_customers)

    print(f"  Запись {idx + 1}:")
    print(f"    Номер документа: {original_doc} => {df_source2.at[idx, 'Номер документа']}")
    print(f"    Покупатель: {current_customer} => {df_source2.at[idx, 'Покупатель']}")

# Сохранение в XLSX
print("\n" + "="*60)
print("Сохранение файлов...")
print("="*60)

xlsx_file1 = "Тестовые_данные_Источник1.xlsx"
xlsx_file2 = "Тестовые_данные_Источник2.xlsx"

df_source1.to_excel(xlsx_file1, index=False, sheet_name='Продажи')
print(f"[OK] Сохранен: {xlsx_file1}")

df_source2.to_excel(xlsx_file2, index=False, sheet_name='Продажи')
print(f"[OK] Сохранен: {xlsx_file2}")

# Сохранение в CSV
csv_file1 = "Тестовые_данные_Источник1.csv"
csv_file2 = "Тестовые_данные_Источник2.csv"

df_source1.to_csv(csv_file1, index=False, encoding='utf-8-sig')
print(f"[OK] Сохранен: {csv_file1}")

df_source2.to_csv(csv_file2, index=False, encoding='utf-8-sig')
print(f"[OK] Сохранен: {csv_file2}")

# Статистика
print("\n" + "="*60)
print("СТАТИСТИКА")
print("="*60)
print(f"[*] Записей в каждом источнике: {len(df_source1)}")
print(f"[*] Столбцов: {len(df_source1.columns)}")
print(f"[*] Изменено записей в Источнике 2: 5")
print(f"[*] Совпадающих записей: {len(df_source1) - 5}")
print(f"[*] Индексы измененных записей: {sorted(indices_to_modify)}")

print("\n" + "="*60)
print("СТРУКТУРА ДАННЫХ")
print("="*60)
print("\nСтолбцы:")
for i, col in enumerate(df_source1.columns, 1):
    print(f"  {i}. {col}")

print("\n" + "="*60)
print("ПРИМЕРЫ ДАННЫХ (первые 5 записей)")
print("="*60)
print("\nИсточник 1:")
print(df_source1.head())
print("\nИсточник 2:")
print(df_source2.head())

print("\n" + "="*60)
print("[OK] ВСЕ ФАЙЛЫ СОЗДАНЫ УСПЕШНО!")
print("="*60)
print("\nСозданные файлы:")
print(f"  1. {xlsx_file1}")
print(f"  2. {xlsx_file2}")
print(f"  3. {csv_file1}")
print(f"  4. {csv_file2}")
print("\nИспользуйте эти файлы для тестирования Expert Excel Matcher v2.0")
