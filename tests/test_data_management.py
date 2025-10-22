"""
Тесты для работы с данными (чтение файлов, валидация)
"""
import sys
from pathlib import Path
import pytest
import pandas as pd
import tkinter as tk
from io import StringIO

root_dir = Path(__file__).parent.parent
sys.path.insert(0, str(root_dir))

from expert_matcher import ExpertMatcher


class TestDataManagement:
    """Тесты управления данными"""

    @pytest.fixture(autouse=True)
    def setup(self, tmp_path):
        """Создание экземпляра и временных файлов"""
        self.root = tk.Tk()
        self.root.withdraw()
        self.matcher = ExpertMatcher(self.root)
        self.tmp_path = tmp_path

        # Создаем тестовые файлы
        self._create_test_files()

        yield
        self.root.destroy()

    def _create_test_files(self):
        """Создание тестовых файлов"""
        # Excel файл
        df_excel = pd.DataFrame({
            'Product': ['MS Office', 'Adobe Reader', 'Chrome'],
            'Version': ['2021', '22.0', '95']
        })
        self.excel_path = self.tmp_path / "test.xlsx"
        df_excel.to_excel(self.excel_path, index=False)

        # CSV файл UTF-8
        df_csv = pd.DataFrame({
            'Software': ['PostgreSQL', 'MongoDB', 'Redis'],
            'Type': ['Database', 'Database', 'Cache']
        })
        self.csv_path = self.tmp_path / "test.csv"
        df_csv.to_csv(self.csv_path, index=False, encoding='utf-8')

        # CSV файл с BOM
        csv_content = '\ufeffName,Category\nPython,Programming\nR,Analytics\n'
        self.csv_bom_path = self.tmp_path / "test_bom.csv"
        with open(self.csv_bom_path, 'w', encoding='utf-8-sig') as f:
            f.write(csv_content)

    def test_read_excel_file(self):
        """Тест чтения Excel файла"""
        df = self.matcher.read_data_file(str(self.excel_path))

        assert df is not None
        assert len(df) == 3
        assert 'Product' in df.columns
        assert df.iloc[0]['Product'] == 'MS Office'

    def test_read_csv_file(self):
        """Тест чтения CSV файла"""
        df = self.matcher.read_data_file(str(self.csv_path))

        assert df is not None
        assert len(df) == 3
        assert 'Software' in df.columns
        assert df.iloc[0]['Software'] == 'PostgreSQL'

    def test_read_csv_with_bom(self):
        """Тест чтения CSV с BOM"""
        df = self.matcher.read_data_file(str(self.csv_bom_path))

        assert df is not None
        assert len(df) == 2
        assert 'Name' in df.columns
        assert df.iloc[0]['Name'] == 'Python'

    def test_read_with_nrows_limit(self):
        """Тест чтения с ограничением строк"""
        df = self.matcher.read_data_file(str(self.excel_path), nrows=2)

        assert df is not None
        assert len(df) == 2

    def test_validate_excel_file_valid(self):
        """Тест валидации корректного файла"""
        result = self.matcher.validate_excel_file(str(self.excel_path))
        # validate_excel_file возвращает tuple (bool, str)
        if isinstance(result, tuple):
            is_valid, message = result
        else:
            is_valid = result
        assert is_valid == True

    def test_validate_excel_file_nonexistent(self):
        """Тест валидации несуществующего файла"""
        result = self.matcher.validate_excel_file(str(self.tmp_path / "nonexistent.xlsx"))
        # validate_excel_file возвращает tuple (bool, str)
        if isinstance(result, tuple):
            is_valid, message = result
        else:
            is_valid = result
        assert is_valid == False

    def test_get_column_display_name_single(self):
        """Тест отображения имени одного столбца"""
        result = self.matcher._get_column_display_name(['Column1'])
        assert result == 'Column1'

    def test_get_column_display_name_multiple(self):
        """Тест отображения имени нескольких столбцов"""
        result = self.matcher._get_column_display_name(['Col1', 'Col2', 'Col3'])
        assert result == 'Col1 + Col2 + Col3'

    def test_get_selected_columns_default(self):
        """Тест получения столбцов по умолчанию"""
        askupo_cols, eatool_cols = self.matcher._get_selected_columns()

        # По умолчанию должны быть пустые списки
        assert askupo_cols == []
        assert eatool_cols == []

    def test_get_selected_columns_with_selection(self):
        """Тест получения выбранных столбцов"""
        self.matcher.selected_askupo_cols = ['Col1', 'Col2']
        self.matcher.selected_eatool_cols = ['ColA']

        askupo_cols, eatool_cols = self.matcher._get_selected_columns()

        assert askupo_cols == ['Col1', 'Col2']
        assert eatool_cols == ['ColA']
