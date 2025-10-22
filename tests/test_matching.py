"""
Тесты для алгоритмов сопоставления
"""
import sys
from pathlib import Path
import pytest
import tkinter as tk

root_dir = Path(__file__).parent.parent
sys.path.insert(0, str(root_dir))

from expert_matcher import ExpertMatcher, MatchingMethod


class TestMatching:
    """Тесты алгоритмов сопоставления"""

    @pytest.fixture(autouse=True)
    def setup(self):
        """Создание экземпляра для тестов"""
        self.root = tk.Tk()
        self.root.withdraw()
        self.matcher = ExpertMatcher(self.root)
        yield
        self.root.destroy()

    def test_exact_match_function(self):
        """Тест точного совпадения (ВПР)"""
        # Идентичные строки
        assert self.matcher.exact_match_func('Microsoft Office', 'Microsoft Office') == 100.0

        # Разный регистр (должны совпасть после нормализации)
        assert self.matcher.exact_match_func('Microsoft Office', 'microsoft office') == 100.0

        # Разные строки
        assert self.matcher.exact_match_func('Microsoft Office', 'MS Office') == 0.0

        # Пробелы (должны совпасть после нормализации)
        assert self.matcher.exact_match_func('  Office  ', 'Office') == 100.0

    def test_combine_columns(self, sample_data_source1):
        """Тест объединения столбцов"""
        row = sample_data_source1.iloc[0]

        # Один столбец
        result = self.matcher.combine_columns(row, ['Название ПО'])
        assert result == 'Microsoft Office 365'

        # Два столбца
        result = self.matcher.combine_columns(row, ['Название ПО', 'Версия'])
        assert result == 'Microsoft Office 365 2021'

        # Три столбца
        result = self.matcher.combine_columns(row, ['Название ПО', 'Версия', 'Vendor'])
        assert result == 'Microsoft Office 365 2021 Microsoft'

        # Пустые значения игнорируются
        row_with_empty = sample_data_source1.iloc[0].copy()
        row_with_empty['Версия'] = ''
        result = self.matcher.combine_columns(row_with_empty, ['Название ПО', 'Версия'])
        assert result == 'Microsoft Office 365'

    def test_matching_method_find_best_match_exact(self):
        """Тест поиска лучшего совпадения - точное совпадение"""
        choices = ['Microsoft Office', 'Adobe Reader', 'Google Chrome']
        choices_normalized = [self.matcher.normalize_string(c) for c in choices]
        choice_dict = {norm: orig for norm, orig in zip(choices_normalized, choices)}

        # Используем встроенный метод точного совпадения
        exact_method = None
        for method in self.matcher.methods:
            if 'Exact Match' in method.name:
                exact_method = method
                break

        assert exact_method is not None, "Exact Match method should be registered"

        query = self.matcher.normalize_string('Microsoft Office')
        match, score = exact_method.find_best_match(query, choices_normalized, choice_dict)

        assert score == 100.0, "Exact match should return 100%"
        assert match == 'Microsoft Office'

    def test_matching_method_find_best_match_no_match(self):
        """Тест когда нет совпадений"""
        choices = ['Microsoft Office', 'Adobe Reader']
        choices_normalized = [self.matcher.normalize_string(c) for c in choices]
        choice_dict = {norm: orig for norm, orig in zip(choices_normalized, choices)}

        exact_method = None
        for method in self.matcher.methods:
            if 'Exact Match' in method.name:
                exact_method = method
                break

        query = self.matcher.normalize_string('PostgreSQL Database')
        match, score = exact_method.find_best_match(query, choices_normalized, choice_dict)

        assert score == 0.0, "No match should return 0%"
        assert match == ''

    def test_matching_method_empty_inputs(self):
        """Тест с пустыми входными данными"""
        exact_method = None
        for method in self.matcher.methods:
            if 'Exact Match' in method.name:
                exact_method = method
                break

        # Пустой query
        match, score = exact_method.find_best_match('', ['test'], {'test': 'test'})
        assert score == 0.0
        assert match == ''

        # Пустой список choices
        match, score = exact_method.find_best_match('test', [], {})
        assert score == 0.0
        assert match == ''

    def test_length_penalty_mechanism(self):
        """Тест механизма штрафа за разницу в длине"""
        # Этот тест проверяет, что короткие строки не ложно совпадают с длинными
        choices = ['NGINX Web Server Enterprise Edition 2021']
        choices_normalized = [self.matcher.normalize_string(c) for c in choices]
        choice_dict = {norm: orig for norm, orig in zip(choices_normalized, choices)}

        # Проверяем rapidfuzz метод (если доступен)
        if any(method.use_process for method in self.matcher.methods):
            rapidfuzz_method = None
            for method in self.matcher.methods:
                if method.use_process and 'Partial' in method.name:
                    rapidfuzz_method = method
                    break

            if rapidfuzz_method:
                # Короткая строка "R" не должна давать высокий score с длинной строкой
                query = self.matcher.normalize_string('R')
                match, score = rapidfuzz_method.find_best_match(
                    query, choices_normalized, choice_dict
                )

                # После штрафа за длину score должен быть < 50 (rejected)
                assert score < 50, f"Short string should not match long string, got score: {score}"

    def test_methods_registration(self):
        """Тест регистрации методов"""
        methods = self.matcher.methods

        assert len(methods) > 0, "Should have at least one method registered"

        # Проверяем, что есть встроенный метод точного совпадения
        exact_methods = [m for m in methods if 'Exact Match' in m.name]
        assert len(exact_methods) == 1, "Should have exactly one Exact Match method"

        # Проверяем структуру MatchingMethod
        for method in methods:
            assert hasattr(method, 'name')
            assert hasattr(method, 'func')
            assert hasattr(method, 'library')
            assert hasattr(method, 'use_process')
            assert hasattr(method, 'scorer')
            assert hasattr(method, 'find_best_match')

    def test_threshold_rejection(self):
        """Тест порога отклонения (50%)"""
        # Этот тест проверяет, что совпадения ниже 50% отклоняются
        # (проверяется в apply_method_optimized, но можем проверить концепцию)
        from expert_matcher import AppConstants

        assert AppConstants.THRESHOLD_REJECT == 50, "Rejection threshold should be 50%"

        # Совпадения ниже 50% должны быть отклонены
        low_score = 45.0
        assert low_score < AppConstants.THRESHOLD_REJECT
