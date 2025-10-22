"""
Тесты для функции расчета статистики
"""
import sys
from pathlib import Path
import pytest
import pandas as pd
import tkinter as tk

root_dir = Path(__file__).parent.parent
sys.path.insert(0, str(root_dir))

from expert_matcher import ExpertMatcher


class TestStatistics:
    """Тесты расчета статистики"""

    @pytest.fixture(autouse=True)
    def setup(self):
        """Создание экземпляра для тестов"""
        self.root = tk.Tk()
        self.root.withdraw()
        self.matcher = ExpertMatcher(self.root)
        yield
        self.root.destroy()

    def test_statistics_categories(self, statistics_test_data):
        """Тест правильности категорий статистики"""
        stats = self.matcher.calculate_statistics(statistics_test_data)

        assert stats['perfect'] == 3, "Should have 3 perfect matches (100%)"
        assert stats['high'] == 3, "Should have 3 high matches (90-99%)"
        assert stats['medium'] == 3, "Should have 3 medium matches (70-89%)"
        assert stats['low'] == 3, "Should have 3 low matches (50-69%)"
        assert stats['very_low'] == 3, "Should have 3 very low matches (1-49%)"
        assert stats['none'] == 2, "Should have 2 no matches (0%)"

    def test_statistics_sum_check(self, statistics_test_data):
        """Тест проверки суммы (критически важно!)"""
        stats = self.matcher.calculate_statistics(statistics_test_data)

        total = stats['total']
        check_sum = stats['check_sum']

        assert total == len(statistics_test_data), "Total should match dataframe length"
        assert check_sum == total, f"Check sum ({check_sum}) should equal total ({total})"

        # Проверяем, что сумма категорий равна total
        categories_sum = (stats['perfect'] + stats['high'] + stats['medium'] +
                         stats['low'] + stats['very_low'] + stats['none'])
        assert categories_sum == total, f"Sum of categories ({categories_sum}) != total ({total})"

    def test_statistics_non_cumulative(self):
        """Тест некумулятивности статистики (каждая запись в одной категории)"""
        # Создаем специальный датасет с известными значениями
        test_data = pd.DataFrame({
            'Процент совпадения': [100, 95, 85, 65, 45, 0]
        })

        stats = self.matcher.calculate_statistics(test_data)

        # Каждая запись должна быть в одной категории
        assert stats['perfect'] == 1
        assert stats['high'] == 1
        assert stats['medium'] == 1
        assert stats['low'] == 1
        assert stats['very_low'] == 1
        assert stats['none'] == 1
        assert stats['check_sum'] == 6

    def test_statistics_boundary_values(self):
        """Тест граничных значений"""
        test_data = pd.DataFrame({
            'Процент совпадения': [
                100,  # perfect (==100)
                99, 90,  # high (90 <= x < 100)
                89.9, 70,  # medium (70 <= x < 90)
                69.9, 50,  # low (50 <= x < 70)
                49.9, 1, 0.1,  # very_low (0 < x < 50)
                0  # none (==0)
            ]
        })

        stats = self.matcher.calculate_statistics(test_data)

        # Проверяем граничные значения согласно логике в calculate_statistics:
        # perfect: == 100
        # high: >= 90 AND < 100
        # medium: >= 70 AND < 90
        # low: >= 50 AND < 70
        # very_low: > 0 AND < 50
        # none: == 0
        assert stats['perfect'] == 1, f"Expected 1 perfect (100), got {stats['perfect']}"
        assert stats['high'] == 2, f"Expected 2 high (99, 90), got {stats['high']}"
        assert stats['medium'] == 2, f"Expected 2 medium (89.9, 70), got {stats['medium']}"
        assert stats['low'] == 2, f"Expected 2 low (69.9, 50), got {stats['low']}"
        assert stats['very_low'] == 3, f"Expected 3 very_low (49.9, 1, 0.1), got {stats['very_low']}"
        assert stats['none'] == 1, f"Expected 1 none (0), got {stats['none']}"
        # Общая проверка суммы
        total_sum = stats['perfect'] + stats['high'] + stats['medium'] + stats['low'] + stats['very_low'] + stats['none']
        assert total_sum == 11, f"Sum of categories {total_sum} != 11"
        assert stats['check_sum'] == 11

    def test_statistics_empty_dataframe(self):
        """Тест с пустым DataFrame"""
        empty_df = pd.DataFrame({'Процент совпадения': []})
        stats = self.matcher.calculate_statistics(empty_df)

        assert stats['total'] == 0
        assert stats['check_sum'] == 0
        assert all(stats[key] == 0 for key in ['perfect', 'high', 'medium', 'low', 'very_low', 'none'])

    def test_statistics_all_perfect(self):
        """Тест когда все совпадения идеальные"""
        perfect_df = pd.DataFrame({'Процент совпадения': [100, 100, 100, 100]})
        stats = self.matcher.calculate_statistics(perfect_df)

        assert stats['perfect'] == 4
        assert stats['high'] == 0
        assert stats['total'] == 4
        assert stats['check_sum'] == 4

    def test_statistics_all_no_match(self):
        """Тест когда нет совпадений"""
        no_match_df = pd.DataFrame({'Процент совпадения': [0, 0, 0]})
        stats = self.matcher.calculate_statistics(no_match_df)

        assert stats['none'] == 3
        assert stats['perfect'] == 0
        assert stats['total'] == 3
        assert stats['check_sum'] == 3
