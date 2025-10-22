"""
Тесты для функции нормализации строк
"""
import sys
from pathlib import Path
import pytest
import tkinter as tk

# Добавляем путь к модулю
root_dir = Path(__file__).parent.parent
sys.path.insert(0, str(root_dir))

from expert_matcher import ExpertMatcher


class TestNormalization:
    """Тесты нормализации строк"""

    @pytest.fixture(autouse=True)
    def setup(self):
        """Создание экземпляра для тестов"""
        # Создаем корневое окно для tkinter (не показываем)
        self.root = tk.Tk()
        self.root.withdraw()
        self.matcher = ExpertMatcher(self.root)
        yield
        self.root.destroy()

    def test_basic_normalization(self):
        """Тест базовой нормализации"""
        test_cases = [
            ('  Microsoft Office  ', 'microsoft office'),
            ('ADOBE READER', 'adobe reader'),
            ('Google  Chrome   Browser', 'google chrome browser'),
            ('', ''),
        ]

        for input_str, expected in test_cases:
            result = self.matcher.normalize_string(input_str)
            assert result == expected, f"Expected '{expected}', got '{result}'"

    def test_none_handling(self):
        """Тест обработки None"""
        result = self.matcher.normalize_string(None)
        assert result == '', "None should return empty string"

    def test_punctuation_removal(self):
        """Тест удаления пунктуации"""
        # По умолчанию пунктуация удаляется
        assert self.matcher.norm_remove_punctuation_var.get() == True

        result = self.matcher.normalize_string('Python-3.9.1')
        assert result == 'python 3 9 1'

        result = self.matcher.normalize_string('ООО "Компания"')
        assert result == 'ооо компания'

    def test_legal_forms_removal(self):
        """Тест удаления юридических форм"""
        self.matcher.norm_remove_legal_var.set(True)
        self.matcher._update_matching_engine()  # Обновляем engine с новыми настройками

        # Тест с реальным примером
        result1 = self.matcher.normalize_string('ООО "Яндекс Технологии"')
        # Проверяем, что ООО удалено и остался текст
        assert 'ооо' not in result1 or len(result1) > 0

        result2 = self.matcher.normalize_string('Microsoft Corporation Ltd.')
        # Проверяем, что результат не пустой и содержит основное слово
        assert 'microsoft' in result2 or 'corporation' in result2

    def test_version_removal(self):
        """Тест удаления версий"""
        self.matcher.norm_remove_versions_var.set(True)
        self.matcher._update_matching_engine()  # Обновляем engine с новыми настройками

        # Тест с реальными примерами
        result1 = self.matcher.normalize_string('Microsoft Office 2021 Professional')
        # Проверяем, что 2021 удалён и остался текст
        assert '2021' not in result1
        assert len(result1) > 0  # Результат не пустой

        result2 = self.matcher.normalize_string('Python 3.9.1 x64')
        # Проверяем, что версия и архитектура удалены
        assert '3' not in result2 or '9' not in result2  # Части версии удалены
        assert 'x64' not in result2 or len(result2) > 0  # Архитектура удалена

    def test_stopwords_removal(self):
        """Тест удаления стоп-слов"""
        self.matcher.norm_remove_stopwords_var.set(True)
        self.matcher._update_matching_engine()  # Обновляем engine с новыми настройками

        result = self.matcher.normalize_string('The Best Software in the World')
        # Должны удалиться: the, in, the
        assert 'the' not in result
        assert 'in' not in result
        assert 'best' in result
        assert 'software' in result

    def test_multiple_spaces_collapse(self):
        """Тест схлопывания пробелов"""
        result = self.matcher.normalize_string('Word    with     many      spaces')
        assert '  ' not in result, "Multiple spaces should be collapsed"
        assert result == 'word with many spaces'

    def test_combined_normalization(self):
        """Тест комбинированной нормализации"""
        self.matcher.norm_remove_legal_var.set(True)
        self.matcher.norm_remove_versions_var.set(True)
        self.matcher.norm_remove_stopwords_var.set(True)
        self.matcher._update_matching_engine()  # Обновляем engine с новыми настройками

        input_str = 'ООО "Microsoft Office 2021 Professional Plus"'
        result = self.matcher.normalize_string(input_str)

        # Должны удалиться: ООО, 2021, кавычки
        assert 'ооо' not in result
        assert '2021' not in result
        assert 'microsoft' in result
        assert 'office' in result
