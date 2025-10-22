# -*- coding: utf-8 -*-
"""
Тест: эмуляция вызова BGE-M3 как в приложении
"""

import sys
import io

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import numpy as np
import pandas as pd
from FlagEmbedding import BGEM3FlagModel

print("=" * 80)
print("ТЕСТ: Эмуляция вызова BGE-M3 метода из приложения")
print("=" * 80)

# Создадим класс-обёртку, имитирующий ExpertMatcher
class TestMatcher:
    def __init__(self):
        self.bge_model = None

    def bge_cosine_similarity(self, s1: str, s2: str) -> float:
        """Копия метода из expert_matcher.py"""
        if not s1 or not s2 or pd.isna(s1) or pd.isna(s2):
            print(f"  ❌ Пустая строка: s1='{s1}', s2='{s2}'")
            return 0.0

        # Строки уже нормализованы через normalize_string(), просто очищаем
        s1 = str(s1).strip()
        s2 = str(s2).strip()

        if not s1 or not s2:
            print(f"  ❌ Пустая после strip: s1='{s1}', s2='{s2}'")
            return 0.0

        # Ленивая загрузка модели BGE-M3
        if self.bge_model is None:
            try:
                print("🧠 Загрузка BGE-M3 модели (может занять 1-2 минуты)...")
                self.bge_model = BGEM3FlagModel(
                    'BAAI/bge-m3',
                    device='cpu',
                    use_fp16=False,
                    normalize_embeddings=True
                )
                print("✅ BGE-M3 модель загружена")
            except Exception as e:
                print(f"❌ Ошибка загрузки BGE-M3: {e}")
                import traceback
                traceback.print_exc()
                self.bge_model = False
                return 0.0

        # Если модель не загрузилась ранее
        if self.bge_model is False:
            print("  ❌ Модель не загружена (флаг False)")
            return 0.0

        try:
            print(f"  🔄 Кодирование: '{s1}' vs '{s2}'")
            # Кодирование строк в векторы (1024-мерные)
            vec1 = self.bge_model.encode(str(s1).lower())['dense_vecs']
            vec2 = self.bge_model.encode(str(s2).lower())['dense_vecs']

            # Косинусное сходство через скалярное произведение
            similarity = float(np.dot(vec1, vec2))

            # Преобразование из диапазона [0, 1] в [0, 100]
            result = similarity * 100.0
            print(f"  ✅ Сходство: {result:.2f}%")
            return result

        except Exception as e:
            print(f"❌ Ошибка вычисления BGE сходства для '{s1}' vs '{s2}': {e}")
            import traceback
            traceback.print_exc()
            return 0.0

# Тестирование
matcher = TestMatcher()

print("\n" + "=" * 80)
print("ТЕСТ 1: Нормализованные строки (как их получает метод)")
print("=" * 80)

test_cases = [
    ("microsoft office", "ms office"),
    ("1s predprijatie", "1c enterprise"),
    ("adobe photoshop", "photoshop"),
]

for s1, s2 in test_cases:
    print(f"\nВызов: bge_cosine_similarity('{s1}', '{s2}')")
    score = matcher.bge_cosine_similarity(s1, s2)
    print(f"Результат: {score:.2f}%")
    if score == 0:
        print("  ⚠️ ПРОБЛЕМА: Метод вернул 0!")

print("\n" + "=" * 80)
print("ТЕСТ 2: Проверка с пустыми строками")
print("=" * 80)

empty_cases = [
    ("", "test"),
    ("test", ""),
    ("", ""),
]

for s1, s2 in empty_cases:
    print(f"\nВызов: bge_cosine_similarity('{s1}', '{s2}')")
    score = matcher.bge_cosine_similarity(s1, s2)
    print(f"Результат: {score:.2f}%")

print("\n" + "=" * 80)
print("ЗАВЕРШЕНО")
print("=" * 80)
