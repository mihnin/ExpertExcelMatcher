"""
Общие фикстуры для тестов
"""
import sys
from pathlib import Path
import pandas as pd
import pytest

# Добавляем корневую директорию в путь для импорта
root_dir = Path(__file__).parent.parent
sys.path.insert(0, str(root_dir))


@pytest.fixture
def sample_data_source1():
    """Тестовые данные источника 1"""
    return pd.DataFrame({
        'Название ПО': [
            'Microsoft Office 365',
            'Adobe Acrobat Reader DC',
            'Google Chrome',
            'Mozilla Firefox',
            'Python 3.9',
            'R Programming Language',
            'PostgreSQL',
            'MongoDB'
        ],
        'Версия': ['2021', '22.0', '95', '92', '3.9', '4.1', '13', '5.0'],
        'Vendor': ['Microsoft', 'Adobe', 'Google', 'Mozilla', 'PSF', 'R Foundation', 'PostgreSQL', 'MongoDB Inc']
    })


@pytest.fixture
def sample_data_source2():
    """Тестовые данные источника 2"""
    return pd.DataFrame({
        'Product Name': [
            'MS Office 365',
            'Acrobat Reader',
            'Chrome Browser',
            'Firefox Web Browser',
            'Python',
            'R Statistical Software',
            'Postgres Database',
            'Mongo DB'
        ],
        'Category': ['Office', 'PDF', 'Browser', 'Browser', 'Programming', 'Analytics', 'Database', 'Database']
    })


@pytest.fixture
def normalization_test_cases():
    """Тестовые случаи для нормализации"""
    return [
        # (input, expected_output_basic)
        ('  Microsoft Office  ', 'microsoft office'),
        ('ADOBE READER', 'adobe reader'),
        ('Google  Chrome   Browser', 'google chrome browser'),
        ('Python-3.9.1', 'python 3 9 1'),
        ('ООО "Компания"', 'ооо компания'),
        ('Ltd. Company', 'ltd company'),
        ('', ''),
        (None, ''),
    ]


@pytest.fixture
def statistics_test_data():
    """Тестовые данные для статистики"""
    return pd.DataFrame({
        'Процент совпадения': [
            100, 100, 100,  # 3 perfect
            95, 92, 90,     # 3 high
            85, 75, 70,     # 3 medium
            65, 55, 50,     # 3 low
            45, 25, 10,     # 3 very_low
            0, 0            # 2 none
        ]
    })
