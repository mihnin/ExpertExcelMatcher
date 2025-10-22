"""
Движок сопоставления для Expert Excel Matcher

Этот модуль содержит логику нормализации строк и сопоставления данных.
"""

import re
import pandas as pd
from typing import List, Dict
from src.constants import NormalizationConstants

# Проверка доступности транслитерации
try:
    from transliterate import translit
    TRANSLITERATE_AVAILABLE = True
except ImportError:
    TRANSLITERATE_AVAILABLE = False


class NormalizationOptions:
    """Настройки нормализации"""

    def __init__(self,
                 remove_legal: bool = False,
                 remove_versions: bool = False,
                 remove_stopwords: bool = False,
                 transliterate: bool = False,
                 remove_punctuation: bool = True):
        """
        Инициализация настроек нормализации

        Args:
            remove_legal: Удалять юридические формы (ООО, Ltd, Inc...)
            remove_versions: Удалять версии (2021, v4.x, R2, SP1, x64...)
            remove_stopwords: Удалять стоп-слова (и, в, the, a...)
            transliterate: Транслитерация кириллицы → латиница
            remove_punctuation: Удалять пунктуацию (по умолчанию True)
        """
        self.remove_legal = remove_legal
        self.remove_versions = remove_versions
        self.remove_stopwords = remove_stopwords
        self.transliterate = transliterate
        self.remove_punctuation = remove_punctuation


class MatchingEngine:
    """Движок нормализации и сопоставления строк"""

    def __init__(self, normalization_options: NormalizationOptions = None):
        """
        Инициализация движка

        Args:
            normalization_options: Настройки нормализации (если None - используются по умолчанию)
        """
        self.norm_options = normalization_options or NormalizationOptions()

    def normalize_string(self, s: str) -> str:
        """
        Расширенная нормализация строки с учётом настроек

        Применяет различные преобразования в зависимости от настроек:
        - Удаление юридических форм (ООО, Ltd, Inc...)
        - Удаление версий (2021, v4.x, R2, SP1, x64...)
        - Удаление стоп-слов (и, в, the, a...)
        - Транслитерация кириллицы → латиница
        - Удаление пунктуации

        Args:
            s: Строка для нормализации

        Returns:
            Нормализованная строка
        """
        if not s or pd.isna(s):
            return ""

        s = str(s).strip()

        # 1. Удаление юридических префиксов (ООО, Ltd, Inc, GmbH...)
        if self.norm_options.remove_legal:
            for pattern in NormalizationConstants.LEGAL_PREFIXES:
                s = re.sub(pattern, ' ', s, flags=re.IGNORECASE)

        # 2. Удаление версий (2021, v4.x, R2, SP1, x64, Windows 10...)
        if self.norm_options.remove_versions:
            for pattern in NormalizationConstants.VERSION_PATTERNS:
                s = re.sub(pattern, ' ', s, flags=re.IGNORECASE)

        # 3. Приведение к нижнему регистру (всегда)
        s = s.lower()

        # 4. Удаление пунктуации (кроме букв, цифр, пробелов)
        if self.norm_options.remove_punctuation:
            s = re.sub(r'[^a-zа-яё0-9\s]', ' ', s)

        # 5. Удаление стоп-слов (и, в, the, a, and...)
        if self.norm_options.remove_stopwords:
            words = s.split()
            words = [w for w in words if w and w not in NormalizationConstants.STOP_WORDS]
            s = ' '.join(words)

        # 6. Транслитерация кириллицы → латиница
        if self.norm_options.transliterate and TRANSLITERATE_AVAILABLE:
            if re.search(r'[а-яё]', s):
                try:
                    s = translit(s, 'ru', reversed=True)
                except Exception:
                    pass  # Если транслитерация не удалась, оставляем как есть

        # 7. Схлопывание пробелов (всегда)
        s = re.sub(r'\s+', ' ', s).strip()

        return s

    def combine_columns(self, row: pd.Series, columns: List[str]) -> str:
        """
        Объединение значений из нескольких столбцов в одну строку

        Args:
            row: строка DataFrame
            columns: список столбцов для объединения

        Returns:
            Объединенная строка (разделитель: пробел)
        """
        values = []
        for col in columns:
            if col in row.index:
                val = row[col]
                if not pd.isna(val) and str(val).strip():
                    values.append(str(val).strip())

        return " ".join(values) if values else ""

    def prepare_choice_dict(self, df: pd.DataFrame, columns: List[str]) -> Dict[str, str]:
        """
        Подготовка словаря нормализованных строк для быстрого поиска

        Args:
            df: DataFrame с данными
            columns: список столбцов для объединения

        Returns:
            Словарь {нормализованная_строка: оригинальная_строка}
        """
        choice_dict = {}
        for _, row in df.iterrows():
            original = self.combine_columns(row, columns)
            normalized = self.normalize_string(original)
            if normalized:  # Пропускаем пустые строки
                choice_dict[normalized] = original

        return choice_dict

    def calculate_statistics(self, results_df: pd.DataFrame) -> Dict:
        """
        ИСПРАВЛЕННАЯ функция подсчета статистики!
        Теперь считает по КАТЕГОРИЯМ, а не накопительно!

        Args:
            results_df: DataFrame с результатами (должен содержать 'Процент совпадения')

        Returns:
            Словарь со статистикой
        """
        total = len(results_df)

        # Категории (НЕ накопительные!)
        perfect = len(results_df[results_df['Процент совпадения'] == 100])
        high = len(results_df[(results_df['Процент совпадения'] >= 90) & (results_df['Процент совпадения'] < 100)])
        medium = len(results_df[(results_df['Процент совпадения'] >= 70) & (results_df['Процент совпадения'] < 90)])
        low = len(results_df[(results_df['Процент совпадения'] >= 50) & (results_df['Процент совпадения'] < 70)])
        very_low = len(results_df[(results_df['Процент совпадения'] > 0) & (results_df['Процент совпадения'] < 50)])
        none = len(results_df[results_df['Процент совпадения'] == 0])

        # ПРОВЕРКА: сумма должна быть равна total
        check_sum = perfect + high + medium + low + very_low + none
        if check_sum != total:
            print(f"⚠️ ВНИМАНИЕ: Ошибка в статистике! {check_sum} != {total}")

        return {
            'total': total,
            'perfect': perfect,      # 100%
            'high': high,            # 90-99%
            'medium': medium,        # 70-89%
            'low': low,              # 50-69%
            'very_low': very_low,    # 1-49%
            'none': none,            # 0%
            'check_sum': check_sum   # Для проверки
        }
