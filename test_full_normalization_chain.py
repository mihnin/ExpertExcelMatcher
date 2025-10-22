# -*- coding: utf-8 -*-
"""
Полный тест: нормализация → BGE → результат
"""

import sys
import io

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import pandas as pd
import numpy as np
from FlagEmbedding import BGEM3FlagModel
import re
from transliterate import translit

# Копируем класс NormalizationConstants
class NormalizationConstants:
    RU_STOP = {"и", "в", "во", "не", "на", "но", "при", "для", "к", "из", "от", "с", "со", "о", "а", "у", "по", "над", "под", "до", "без", "или"}
    EN_STOP = {"the", "a", "an", "and", "or", "of", "for", "in", "on", "at", "to", "from", "with", "by", "without", "into", "out", "over", "under", "above", "below"}
    STOP_WORDS = RU_STOP | EN_STOP

    LEGAL_PREFIXES = [
        r'\bООО\b', r'\bАО\b', r'\bЗАО\b', r'\bИП\b', r'\bПАО\b', r'\bГК\b',
        r'\bНКО\b', r'\bНПО\b', r'\bНПП\b', r'\bНПФ\b', r'\bОАО\b',
        r'\bLtd\.?\b', r'\bLimited\b', r'\bInc\.?\b', r'\bLLC\b', r'\bGmbH\b',
        r'\bCorp\.?\b', r'\bCo\.?\b', r'\bSARL\b', r'\bS\.?A\.?\b',
        r'\bPLC\b', r'\bGroup\b', r'\bCompany\b', r'\bКомпания\b',
    ]

    VERSION_PATTERNS = [
        r'\b(19|20)\d{2}\b',
        r'\b[vV]\.?\d+\.[xX]\b',
        r'\b\d+\.[xX]\b',
        r'\b[vV]\.?\d+(\.\d+)*[a-z]*\b',
        r'\b\d+\.\d+(\.\d+)*[a-z]*\b',
        r'\bR\d+\b',
        r'\bSP\d+\b',
        r'\b(x64|x86|64[-\s]?bit|32[-\s]?bit)\b',
        r'\bCC\b',
    ]

# Копируем метод normalize_string из приложения
class TestMatcher:
    def __init__(self):
        self.bge_model = None
        # Чекбоксы (ВСЕ включены)
        self.norm_remove_legal = True
        self.norm_remove_versions = True
        self.norm_remove_stopwords = True
        self.norm_transliterate = True
        self.norm_remove_punctuation = True

    def normalize_string(self, s: str) -> str:
        """Копия из expert_matcher.py"""
        if not s or pd.isna(s):
            return ""
        s = str(s).strip()

        # 1. Удаление юридических форм
        if self.norm_remove_legal:
            for pattern in NormalizationConstants.LEGAL_PREFIXES:
                s = re.sub(pattern, ' ', s, flags=re.IGNORECASE)

        # 2. Удаление версий
        if self.norm_remove_versions:
            for pattern in NormalizationConstants.VERSION_PATTERNS:
                s = re.sub(pattern, ' ', s, flags=re.IGNORECASE)

        # 3. Lowercase
        s = s.lower()

        # 4. Удаление пунктуации
        if self.norm_remove_punctuation:
            s = re.sub(r'[^\w\s]', ' ', s)

        # 5. Удаление стоп-слов
        if self.norm_remove_stopwords:
            words = s.split()
            words = [w for w in words if w and w not in NormalizationConstants.STOP_WORDS]
            s = ' '.join(words)

        # 6. Транслитерация
        if self.norm_transliterate:
            if re.search(r'[а-яё]', s):
                try:
                    s = translit(s, 'ru', reversed=True)
                except Exception:
                    pass

        # 7. Очистка пробелов
        s = re.sub(r'\s+', ' ', s).strip()

        return s

    def bge_cosine_similarity(self, s1: str, s2: str) -> float:
        """Копия из expert_matcher.py"""
        if not s1 or not s2 or pd.isna(s1) or pd.isna(s2):
            return 0.0

        s1 = str(s1).strip()
        s2 = str(s2).strip()

        if not s1 or not s2:
            return 0.0

        if self.bge_model is None:
            try:
                print("  🧠 Загрузка BGE-M3...")
                self.bge_model = BGEM3FlagModel(
                    'BAAI/bge-m3',
                    device='cpu',
                    use_fp16=False,
                    normalize_embeddings=True
                )
                print("  ✅ Модель загружена")
            except Exception as e:
                print(f"  ❌ Ошибка загрузки: {e}")
                self.bge_model = False
                return 0.0

        if self.bge_model is False:
            return 0.0

        try:
            vec1 = self.bge_model.encode(str(s1).lower())['dense_vecs']
            vec2 = self.bge_model.encode(str(s2).lower())['dense_vecs']
            similarity = float(np.dot(vec1, vec2))
            return similarity * 100.0
        except Exception as e:
            print(f"  ❌ Ошибка: {e}")
            return 0.0

print("=" * 80)
print("ПОЛНЫЙ ТЕСТ: Исходные строки → Нормализация → BGE-M3")
print("=" * 80)

matcher = TestMatcher()

test_cases = [
    ("Microsoft Office 2021", "MS Office Professional"),
    ("ООО 1С Предприятие 8.3 x64", "1C Enterprise"),
    ("Adobe Photoshop CC 2019", "Photoshop"),
    ("Oracle Database 19c", "Oracle DB"),
]

for original1, original2 in test_cases:
    print(f"\n{'=' * 80}")
    print(f"Исходные строки:")
    print(f"  '{original1}' vs '{original2}'")

    # Нормализация
    norm1 = matcher.normalize_string(original1)
    norm2 = matcher.normalize_string(original2)

    print(f"\nПосле нормализации:")
    print(f"  '{norm1}' vs '{norm2}'")

    if not norm1:
        print(f"  ⚠️ ПРОБЛЕМА: Первая строка стала ПУСТОЙ после нормализации!")
    if not norm2:
        print(f"  ⚠️ ПРОБЛЕМА: Вторая строка стала ПУСТОЙ после нормализации!")

    if not norm1 or not norm2:
        print(f"\nРезультат BGE-M3: 0.00% (пустая строка)")
        continue

    # Вызов BGE
    score = matcher.bge_cosine_similarity(norm1, norm2)
    print(f"\nРезультат BGE-M3: {score:.2f}%")

    if score == 0:
        print("  ⚠️ ПРОБЛЕМА: BGE вернул 0 для непустых строк!")

print("\n" + "=" * 80)
print("ЗАВЕРШЕНО")
print("=" * 80)
