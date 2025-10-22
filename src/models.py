"""
Модели данных для Expert Excel Matcher

Этот модуль содержит классы данных и модели, используемые в приложении:
- MatchingMethod: Класс метода сопоставления
- MatchResult: Результат сопоставления одной записи (dataclass)
- MethodStatistics: Статистика работы метода (dataclass)
"""

from dataclasses import dataclass, field
from typing import Dict, List, Tuple, Callable, Optional
import pandas as pd

# Флаги доступности библиотек (будут установлены при импорте)
RAPIDFUZZ_AVAILABLE = False
process = None

try:
    from rapidfuzz import process as _process
    RAPIDFUZZ_AVAILABLE = True
    process = _process
except ImportError:
    pass


class MatchingMethod:
    """Класс для описания метода сопоставления"""

    def __init__(self, name: str, func: Callable, library: str,
                 use_process: bool = False, scorer=None, use_original_strings: bool = False):
        """
        Инициализация метода сопоставления

        Args:
            name: Название метода (отображаемое)
            func: Функция сопоставления
            library: Библиотека (rapidfuzz, textdistance, jellyfish, builtin)
            use_process: Использовать ли rapidfuzz.process.extractOne (оптимизация)
            scorer: Scorer для rapidfuzz (если use_process=True)
            use_original_strings: Legacy параметр (не используется)
        """
        self.name = name
        self.func = func
        self.library = library
        self.use_process = use_process
        self.scorer = scorer
        self.use_original_strings = use_original_strings  # Не используется (Legacy)

    def find_best_match(self, query: str, choices: List[str],
                       choice_dict: Dict[str, str]) -> Tuple[str, float]:
        """
        Поиск лучшего совпадения с учетом длины строк

        Args:
            query: Нормализованная строка запроса
            choices: Список нормализованных строк для сравнения
            choice_dict: Словарь {нормализованная_строка: оригинальная_строка}

        Returns:
            Tuple[str, float]: (оригинальная строка совпадения, процент совпадения)
        """
        if not query or not choices:
            return "", 0.0

        try:
            query_len = len(query)

            if self.use_process and RAPIDFUZZ_AVAILABLE and not self.use_original_strings:
                # RapidFuzz process работает только с нормализованными строками
                result = process.extractOne(
                    query,
                    choices,
                    scorer=self.scorer,
                    score_cutoff=50
                )
                if result:
                    match_normalized, score, _ = result
                    original_match = choice_dict.get(match_normalized, "")

                    # Применяем штраф за разницу в длине
                    match_len = len(original_match)
                    length_ratio = min(query_len, match_len) / max(query_len, match_len) if max(query_len, match_len) > 0 else 0

                    # Штраф: если длины очень разные, снижаем score
                    # Для коротких строк (<=3 символа) штраф сильнее
                    if query_len <= 3 or match_len <= 3:
                        # Для очень коротких строк требуем почти точное совпадение длин
                        length_penalty = length_ratio ** 2  # Квадратичный штраф
                    else:
                        # Для длинных строк штраф мягче
                        length_penalty = length_ratio ** 0.5  # Корень квадратный

                    adjusted_score = float(score) * length_penalty

                    # Если после штрафа score < 50, отбрасываем
                    if adjusted_score < 50:
                        return "", 0.0

                    return original_match, adjusted_score
                return "", 0.0
            else:
                # Ручной перебор для других библиотек
                best_match = ""
                best_score = 0.0

                for choice in choices:
                    try:
                        score = self.func(query, choice)
                        # Нормализация score в диапазон 0-100
                        if isinstance(score, float) and 0 <= score <= 1:
                            score = score * 100
                        score = float(score)

                        # Применяем штраф за разницу в длине
                        choice_len = len(choice)
                        length_ratio = min(query_len, choice_len) / max(query_len, choice_len) if max(query_len, choice_len) > 0 else 0

                        if query_len <= 3 or choice_len <= 3:
                            length_penalty = length_ratio ** 2
                        else:
                            length_penalty = length_ratio ** 0.5

                        adjusted_score = score * length_penalty

                        if adjusted_score > best_score:
                            best_score = adjusted_score
                            best_match = choice_dict.get(choice, "")

                            if best_score >= 99.9:
                                break
                    except Exception:
                        continue

                return best_match, best_score
        except Exception:
            return "", 0.0


@dataclass
class MatchResult:
    """Результат сопоставления одной записи"""

    source1_value: str
    """Значение из источника 1 (целевое)"""

    source2_value: str
    """Найденное значение из источника 2"""

    match_score: float
    """Процент совпадения (0-100)"""

    method_name: str
    """Название использованного метода"""

    source1_metadata: Dict[str, any] = field(default_factory=dict)
    """Дополнительные метаданные из источника 1"""

    source2_metadata: Optional[Dict[str, any]] = field(default_factory=dict)
    """Дополнительные метаданные из источника 2"""

    def to_dict(self) -> Dict:
        """Преобразование в словарь для DataFrame"""
        result = {
            'source1_value': self.source1_value,
            'source2_value': self.source2_value,
            'match_score': self.match_score,
            'method_name': self.method_name,
        }
        result.update(self.source1_metadata)
        if self.source2_metadata:
            result.update(self.source2_metadata)
        return result


@dataclass
class MethodStatistics:
    """Статистика работы метода сопоставления"""

    method_name: str
    """Название метода"""

    total: int
    """Общее количество записей"""

    perfect: int = 0
    """Идеальные совпадения (100%)"""

    high: int = 0
    """Высокие совпадения (90-99%)"""

    medium: int = 0
    """Средние совпадения (70-89%)"""

    low: int = 0
    """Низкие совпадения (50-69%)"""

    very_low: int = 0
    """Очень низкие совпадения (1-49%)"""

    none: int = 0
    """Нет совпадения (0%)"""

    avg_score: float = 0.0
    """Средний процент совпадения"""

    processing_time: float = 0.0
    """Время обработки в секундах"""

    @property
    def check_sum(self) -> int:
        """Проверочная сумма: сумма всех категорий должна равняться total"""
        return self.perfect + self.high + self.medium + self.low + self.very_low + self.none

    @property
    def is_valid(self) -> bool:
        """Проверка корректности статистики"""
        return self.check_sum == self.total

    @property
    def sorting_key(self) -> Tuple[int, int, float]:
        """
        Лексикографический ключ для сортировки методов

        Приоритеты:
        1. Максимум идеальных совпадений (100%)
        2. Максимум высоких совпадений (90-99%)
        3. Максимальный средний процент
        """
        return (self.perfect, self.high, self.avg_score)

    def __lt__(self, other: 'MethodStatistics') -> bool:
        """Сравнение для сортировки (больший - лучше)"""
        return self.sorting_key < other.sorting_key

    def __gt__(self, other: 'MethodStatistics') -> bool:
        """Сравнение для сортировки (больший - лучше)"""
        return self.sorting_key > other.sorting_key

    def to_dict(self) -> Dict:
        """Преобразование в словарь"""
        return {
            'method_name': self.method_name,
            'total': self.total,
            'perfect': self.perfect,
            'high': self.high,
            'medium': self.medium,
            'low': self.low,
            'very_low': self.very_low,
            'none': self.none,
            'avg_score': self.avg_score,
            'processing_time': self.processing_time,
            'check_sum': self.check_sum,
            'is_valid': self.is_valid,
        }

    @classmethod
    def from_results_df(cls, method_name: str, results_df: pd.DataFrame,
                       processing_time: float = 0.0) -> 'MethodStatistics':
        """
        Создание статистики из DataFrame результатов

        Args:
            method_name: Название метода
            results_df: DataFrame с результатами (должен содержать 'Процент совпадения')
            processing_time: Время обработки в секундах

        Returns:
            MethodStatistics: Объект статистики
        """
        total = len(results_df)

        # Категории (НЕ накопительные!)
        perfect = len(results_df[results_df['Процент совпадения'] == 100])
        high = len(results_df[(results_df['Процент совпадения'] >= 90) & (results_df['Процент совпадения'] < 100)])
        medium = len(results_df[(results_df['Процент совпадения'] >= 70) & (results_df['Процент совпадения'] < 90)])
        low = len(results_df[(results_df['Процент совпадения'] >= 50) & (results_df['Процент совпадения'] < 70)])
        very_low = len(results_df[(results_df['Процент совпадения'] > 0) & (results_df['Процент совпадения'] < 50)])
        none = len(results_df[results_df['Процент совпадения'] == 0])

        # Средний процент
        avg_score = float(results_df['Процент совпадения'].mean()) if total > 0 else 0.0

        return cls(
            method_name=method_name,
            total=total,
            perfect=perfect,
            high=high,
            medium=medium,
            low=low,
            very_low=very_low,
            none=none,
            avg_score=avg_score,
            processing_time=processing_time
        )
