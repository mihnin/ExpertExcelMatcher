# -*- coding: utf-8 -*-
"""
Тест BGE-M3 с нормализованными строками (как в Notebook)
"""

import sys
import io

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import numpy as np
from FlagEmbedding import BGEM3FlagModel
import pandas as pd
import re
from transliterate import translit

# Импортируем нормализацию из приложения
# (копируем функцию normalize_string)

class NormalizationConstants:
    """Константы для расширенной нормализации текста (подход Архитекторов)"""
    # Стоп-слова (русские и английские)
    RU_STOP = {"и", "в", "во", "не", "на", "но", "при", "для", "к", "из", "от", "с", "со", "о", "а", "у", "по", "над", "под", "до", "без", "или"}
    EN_STOP = {"the", "a", "an", "and", "or", "of", "for", "in", "on", "at", "to", "from", "with", "by", "without", "into", "out", "over", "under", "above", "below"}
    STOP_WORDS = RU_STOP | EN_STOP

    # Регулярки для юридических форм
    LEGAL_PREFIXES = [
        r'\bООО\b', r'\bАО\b', r'\bЗАО\b', r'\bИП\b', r'\bПАО\b', r'\bГК\b',
        r'\bНКО\b', r'\bНПО\b', r'\bНПП\b', r'\bНПФ\b', r'\bОАО\b',
        r'\bLtd\.?\b', r'\bLimited\b', r'\bInc\.?\b', r'\bLLC\b', r'\bGmbH\b',
        r'\bCorp\.?\b', r'\bCo\.?\b', r'\bSARL\b', r'\bS\.?A\.?\b',
        r'\bPLC\b', r'\bGroup\b', r'\bCompany\b', r'\bКомпания\b',
        r'\bИндивидуальный предприниматель\b',
        r'\bОбщество с ограниченной ответственностью\b'
    ]

    # Версионные паттерны
    VERSION_PATTERNS = [
        r'\b(19|20)\d{2}\b',                    # годы (2019, 2021)
        r'\b[vV]\.?\d+\.[xX]\b',                # v.4.x, v4.x
        r'\b\d+\.[xX]\b',                       # 8.x
        r'\b[vV]\.?\d+(\.\d+)*[a-z]*\b',        # v.4, v4, v.1.2
        r'\b\d+\.\d+(\.\d+)*[a-z]*\b',          # 8.1, 2021.1a
        r'\bR\d+\b',                            # R2
        r'\bSP\d+\b',                           # SP1
        r'\b(x64|x86|64[-\s]?bit|32[-\s]?bit)\b',
        r'\bCC\b',                              # Adobe CC
    ]

def normalize_string(s: str,
                    remove_legal: bool = True,
                    remove_versions: bool = True,
                    remove_stopwords: bool = True,
                    transliterate_text: bool = True,
                    remove_punctuation: bool = True) -> str:
    """Нормализация строки"""
    if not s or pd.isna(s):
        return ""
    s = str(s).strip()

    # 1. Удаление юридических форм
    if remove_legal:
        for pattern in NormalizationConstants.LEGAL_PREFIXES:
            s = re.sub(pattern, ' ', s, flags=re.IGNORECASE)

    # 2. Удаление версий
    if remove_versions:
        for pattern in NormalizationConstants.VERSION_PATTERNS:
            s = re.sub(pattern, ' ', s, flags=re.IGNORECASE)

    # 3. Lowercase
    s = s.lower()

    # 4. Удаление пунктуации
    if remove_punctuation:
        s = re.sub(r'[^\w\s]', ' ', s)

    # 5. Удаление стоп-слов
    if remove_stopwords:
        words = s.split()
        words = [w for w in words if w and w not in NormalizationConstants.STOP_WORDS]
        s = ' '.join(words)

    # 6. Транслитерация
    if transliterate_text:
        if re.search(r'[а-яё]', s):
            try:
                s = translit(s, 'ru', reversed=True)
            except Exception:
                pass

    # 7. Очистка множественных пробелов
    s = re.sub(r'\s+', ' ', s).strip()

    return s

print("=" * 80)
print("ТЕСТ: BGE-M3 с полной нормализацией (как в Notebook)")
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
    vec1 = model.encode(s1.lower())['dense_vecs']
    vec2 = model.encode(s2.lower())['dense_vecs']
    return float(np.dot(vec1, vec2)) * 100.0

for str1, str2 in test_pairs:
    print(f"Исходные строки: '{str1}' vs '{str2}'")

    # Полная нормализация (ВСЕ опции включены)
    norm1 = normalize_string(str1,
                            remove_legal=True,
                            remove_versions=True,
                            remove_stopwords=True,
                            transliterate_text=True,
                            remove_punctuation=True)
    norm2 = normalize_string(str2,
                            remove_legal=True,
                            remove_versions=True,
                            remove_stopwords=True,
                            transliterate_text=True,
                            remove_punctuation=True)

    print(f"  После нормализации: '{norm1}' vs '{norm2}'")

    if not norm1 or not norm2:
        print(f"  ⚠️ ПУСТАЯ СТРОКА после нормализации! Score = 0")
    else:
        score = calculate_similarity(norm1, norm2)
        print(f"  BGE-M3 Score: {score:.2f}%")

    print()

print("=" * 80)
print("ВЫВОД:")
print("Если score > 0, значит BGE-M3 работает с нормализованными строками!")
print("=" * 80)
