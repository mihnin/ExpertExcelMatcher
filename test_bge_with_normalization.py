# -*- coding: utf-8 -*-
"""
Тест BGE-M3 с разными уровнями нормализации
"""

import sys
import io

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import numpy as np
from FlagEmbedding import BGEM3FlagModel
from transliterate import translit
import re

print("=" * 80)
print("ТЕСТ: Влияние нормализации на BGE-M3")
print("=" * 80)

# Загрузка модели
print("\n🧠 Загрузка модели...")
model = BGEM3FlagModel('BAAI/bge-m3', device='cpu', use_fp16=False, normalize_embeddings=True)
print("✅ Модель загружена\n")

# Тестовые пары
test_pairs = [
    ("Microsoft Office", "MS Office"),
    ("ООО 1С Предприятие 8.3", "1C Enterprise"),
    ("Adobe Photoshop CC 2019", "Photoshop"),
]

def calculate_similarity(s1, s2):
    vec1 = model.encode(s1)['dense_vecs']
    vec2 = model.encode(s2)['dense_vecs']
    return float(np.dot(vec1, vec2)) * 100.0

def aggressive_normalize(s):
    """Имитация агрессивной нормализации из приложения"""
    s = str(s).lower()

    # Удаление юр.форм
    s = re.sub(r'\bООО\b', '', s, flags=re.IGNORECASE)

    # Удаление версий
    s = re.sub(r'\b(19|20)\d{2}\b', '', s)
    s = re.sub(r'\b[vV]\.?\d+\.?\w*\b', '', s)
    s = re.sub(r'\bCC\b', '', s)
    s = re.sub(r'\b\d+\.\d+\b', '', s)

    # Удаление пунктуации
    s = re.sub(r'[^\w\s]', ' ', s)

    # Транслитерация (критичная часть!)
    try:
        # Пытаемся транслитерировать кириллицу
        if any(ord(c) > 127 for c in s):
            s = translit(s, 'ru', reversed=True)
    except:
        pass

    # Очистка пробелов
    s = re.sub(r'\s+', ' ', s).strip()

    return s

for str1, str2 in test_pairs:
    print(f"Исходные строки: '{str1}' vs '{str2}'")

    # Без нормализации (только lowercase)
    score1 = calculate_similarity(str1.lower(), str2.lower())
    print(f"  Без нормализации: {score1:.2f}%")

    # С агрессивной нормализацией
    norm1 = aggressive_normalize(str1)
    norm2 = aggressive_normalize(str2)
    print(f"  После нормализации: '{norm1}' vs '{norm2}'")
    score2 = calculate_similarity(norm1, norm2)
    print(f"  С нормализацией: {score2:.2f}%")

    diff = score1 - score2
    if diff > 10:
        print(f"  ⚠️ ПОТЕРЯ ТОЧНОСТИ: -{diff:.2f}%")

    print()

print("=" * 80)
print("ВЫВОД:")
print("Если транслитерация снижает точность, значит BGE-M3 НЕ должна")
print("использоваться с агрессивной нормализацией!")
print("=" * 80)
