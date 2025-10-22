# -*- coding: utf-8 -*-
"""
Тестирование 19-го метода BGE-M3
"""

import sys
import io

# Исправляем кодировку консоли для Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import numpy as np

# Проверяем доступность BGE
try:
    from FlagEmbedding import BGEM3FlagModel
    print("✅ FlagEmbedding импортирована успешно")
except ImportError as e:
    print(f"❌ Ошибка импорта FlagEmbedding: {e}")
    sys.exit(1)

print("\n" + "="*80)
print("ТЕСТ 1: Загрузка модели BGE-M3")
print("="*80)

try:
    print("🧠 Загрузка BGE-M3 модели...")
    print("   Это может занять 1-2 минуты при первом запуске (скачивание ~2 GB)")

    model = BGEM3FlagModel(
        'BAAI/bge-m3',
        device='cpu',
        use_fp16=False,
        normalize_embeddings=True
    )
    print("✅ Модель загружена успешно!")

except Exception as e:
    print(f"❌ Ошибка загрузки модели: {e}")
    print(f"   Тип ошибки: {type(e).__name__}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print("\n" + "="*80)
print("ТЕСТ 2: Кодирование строк в векторы")
print("="*80)

test_strings = [
    ("Microsoft Office", "MS Office"),
    ("1C Enterprise", "1С Предприятие"),
    ("Adobe Photoshop", "Photoshop CC"),
    ("Oracle Database", "Oracle DB"),
    ("Google Chrome", "Chrome Browser"),
]

for str1, str2 in test_strings:
    print(f"\nСравнение: '{str1}' vs '{str2}'")

    try:
        # Кодирование
        vec1 = model.encode(str1.lower())['dense_vecs']
        vec2 = model.encode(str2.lower())['dense_vecs']

        print(f"  Размерность vec1: {vec1.shape}")
        print(f"  Размерность vec2: {vec2.shape}")

        # Косинусное сходство
        similarity = float(np.dot(vec1, vec2))
        similarity_percent = similarity * 100.0

        print(f"  Сходство: {similarity:.6f} ({similarity_percent:.2f}%)")

    except Exception as e:
        print(f"  ❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()

print("\n" + "="*80)
print("ТЕСТ 3: Работа с нормализованными строками")
print("="*80)

# Эмулируем нормализацию из приложения
def normalize_string(s):
    """Базовая нормализация"""
    if not s:
        return ""
    return str(s).lower().strip()

normalized_pairs = [
    (normalize_string("ООО Microsoft Office 2021 x64"),
     normalize_string("Microsoft Office Professional")),
    (normalize_string("Adobe Photoshop CC 2019"),
     normalize_string("Photoshop")),
]

for str1, str2 in normalized_pairs:
    print(f"\nНормализованные: '{str1}' vs '{str2}'")

    try:
        vec1 = model.encode(str1)['dense_vecs']
        vec2 = model.encode(str2)['dense_vecs']

        similarity = float(np.dot(vec1, vec2))
        similarity_percent = similarity * 100.0

        print(f"  Сходство: {similarity:.6f} ({similarity_percent:.2f}%)")

    except Exception as e:
        print(f"  ❌ Ошибка: {e}")

print("\n" + "="*80)
print("ТЕСТ 4: Проверка формата выхода encode()")
print("="*80)

test_str = "Microsoft Office"
result = model.encode(test_str)

print(f"Тип результата: {type(result)}")
print(f"Ключи в результате: {result.keys() if isinstance(result, dict) else 'не словарь!'}")

if isinstance(result, dict) and 'dense_vecs' in result:
    dense = result['dense_vecs']
    print(f"Тип dense_vecs: {type(dense)}")
    print(f"Shape dense_vecs: {dense.shape if hasattr(dense, 'shape') else 'нет shape'}")
    print(f"Первые 5 элементов: {dense[:5]}")
    print("✅ Формат правильный!")
else:
    print("❌ Формат НЕ правильный! Ожидался dict с ключом 'dense_vecs'")

print("\n" + "="*80)
print("✅ ТЕСТИРОВАНИЕ ЗАВЕРШЕНО")
print("="*80)
