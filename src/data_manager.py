"""
Модуль управления данными для Expert Excel Matcher

Этот модуль содержит класс DataManager для работы с файлами,
столбцами и валидацией данных.
"""

import pandas as pd
from pathlib import Path
from typing import Tuple, List, Optional


class DataManager:
    """Класс для управления данными (файлы, столбцы, валидация)"""

    def __init__(self):
        """Инициализация менеджера данных"""
        # Файлы
        self.source1_file: Optional[str] = None
        self.source2_file: Optional[str] = None

        # Столбцы
        self.source1_columns: List[str] = []
        self.source2_columns: List[str] = []
        self.selected_source1_cols: List[str] = []
        self.selected_source2_cols: List[str] = []

    def read_data_file(self, filename: str, nrows=None) -> pd.DataFrame:
        """
        Универсальное чтение Excel или CSV файла

        Args:
            filename: Путь к файлу
            nrows: Количество строк для чтения (None = все)

        Returns:
            DataFrame с данными
        """
        file_ext = Path(filename).suffix.lower()

        if file_ext == '.csv':
            # Пробуем различные кодировки для CSV
            encodings = ['utf-8-sig', 'utf-8', 'cp1251', 'windows-1251', 'latin1']
            for encoding in encodings:
                try:
                    df = pd.read_csv(filename, encoding=encoding, nrows=nrows)
                    print(f"DEBUG: CSV файл прочитан с кодировкой {encoding}")
                    print(f"DEBUG: Первые 3 значения первого столбца: {df.iloc[:3, 0].tolist()}")
                    return df
                except (UnicodeDecodeError, Exception):
                    continue
            # Если ничего не сработало, пробуем без указания кодировки
            df = pd.read_csv(filename, nrows=nrows)
        else:
            # Excel файлы (.xlsx, .xls)
            df = pd.read_excel(filename, nrows=nrows)
            print(f"DEBUG: Excel файл прочитан")
            print(f"DEBUG: Расширение файла: {file_ext}")
            print(f"DEBUG: Первые 3 значения первого столбца: {df.iloc[:3, 0].tolist()}")

        return df

    def validate_file(self, filename: str) -> Tuple[bool, str]:
        """
        Валидация Excel или CSV файла

        Args:
            filename: Путь к файлу

        Returns:
            Tuple[bool, str]: (успешная_валидация, сообщение)
        """
        try:
            df = self.read_data_file(filename)

            if df.empty:
                return False, "Файл пустой (нет данных)"

            if len(df.columns) == 0:
                return False, "Файл не содержит столбцов"

            if len(df) == 0:
                return False, "Файл не содержит строк с данными"

            # Успешная валидация - показываем информацию о файле
            columns_preview = ', '.join(df.columns[:5])
            if len(df.columns) > 5:
                columns_preview += ' ...'

            return True, f"✅ Файл валидный\n   Записей: {len(df)}\n   Столбцов: {len(df.columns)}\n   Список столбцов: {columns_preview}"

        except Exception as e:
            return False, f"Ошибка чтения файла:\n{str(e)}"

    def set_source1_file(self, filename: str) -> Tuple[bool, str]:
        """
        Установить файл источника 1 (с валидацией)

        Args:
            filename: Путь к файлу

        Returns:
            Tuple[bool, str]: (успешно, сообщение)
        """
        is_valid, message = self.validate_file(filename)

        if is_valid:
            self.source1_file = filename
            # Загружаем столбцы
            try:
                df = self.read_data_file(filename, nrows=0)
                self.source1_columns = list(df.columns)
                # По умолчанию выбираем первый столбец
                if self.source1_columns:
                    self.selected_source1_cols = [self.source1_columns[0]]
            except Exception as e:
                return False, f"Ошибка загрузки столбцов:\n{str(e)}"

        return is_valid, message

    def set_source2_file(self, filename: str) -> Tuple[bool, str]:
        """
        Установить файл источника 2 (с валидацией)

        Args:
            filename: Путь к файлу

        Returns:
            Tuple[bool, str]: (успешно, сообщение)
        """
        is_valid, message = self.validate_file(filename)

        if is_valid:
            self.source2_file = filename
            # Загружаем столбцы
            try:
                df = self.read_data_file(filename, nrows=0)
                self.source2_columns = list(df.columns)
                # По умолчанию выбираем первый столбец
                if self.source2_columns:
                    self.selected_source2_cols = [self.source2_columns[0]]
            except Exception as e:
                return False, f"Ошибка загрузки столбцов:\n{str(e)}"

        return is_valid, message

    def set_selected_source1_columns(self, columns: List[str]) -> bool:
        """
        Установить выбранные столбцы источника 1

        Args:
            columns: Список столбцов (максимум 2)

        Returns:
            bool: Успешно ли
        """
        if len(columns) > 2:
            return False

        # Проверяем, что все столбцы существуют
        for col in columns:
            if col not in self.source1_columns:
                return False

        self.selected_source1_cols = columns
        return True

    def set_selected_source2_columns(self, columns: List[str]) -> bool:
        """
        Установить выбранные столбцы источника 2

        Args:
            columns: Список столбцов (максимум 2)

        Returns:
            bool: Успешно ли
        """
        if len(columns) > 2:
            return False

        # Проверяем, что все столбцы существуют
        for col in columns:
            if col not in self.source2_columns:
                return False

        self.selected_source2_cols = columns
        return True

    def is_ready(self) -> bool:
        """
        Проверка готовности к обработке

        Returns:
            bool: Готовы ли данные (оба файла выбраны)
        """
        return bool(self.source1_file and self.source2_file)

    def get_column_display_name(self, columns: List[str]) -> str:
        """
        Получить отображаемое имя для списка столбцов

        Args:
            columns: Список столбцов

        Returns:
            str: Отображаемое имя
        """
        if not columns:
            return ""
        elif len(columns) == 1:
            return columns[0]
        else:
            return " + ".join(columns)

    def get_selected_columns(self):
        """
        Получить выбранные столбцы для обоих источников

        Returns:
            Tuple[List[str], List[str]]: (столбцы_источника_1, столбцы_источника_2)
        """
        return self.selected_source1_cols, self.selected_source2_cols

    def get_short_filename(self, filename: Optional[str], max_length: int = 50) -> str:
        """
        Получить короткое имя файла для отображения

        Args:
            filename: Полный путь к файлу
            max_length: Максимальная длина

        Returns:
            str: Короткое имя файла
        """
        if not filename:
            return "Файл не выбран"

        display_name = Path(filename).name
        if len(display_name) > max_length:
            display_name = display_name[:max_length - 3] + "..."

        return display_name

    def load_source1_data(self, nrows=None) -> pd.DataFrame:
        """
        Загрузить данные из источника 1

        Args:
            nrows: Количество строк (None = все)

        Returns:
            DataFrame с данными

        Raises:
            ValueError: Если файл не выбран
        """
        if not self.source1_file:
            raise ValueError("Файл источника 1 не выбран")

        return self.read_data_file(self.source1_file, nrows=nrows)

    def load_source2_data(self, nrows=None) -> pd.DataFrame:
        """
        Загрузить данные из источника 2

        Args:
            nrows: Количество строк (None = все)

        Returns:
            DataFrame с данными

        Raises:
            ValueError: Если файл не выбран
        """
        if not self.source2_file:
            raise ValueError("Файл источника 2 не выбран")

        return self.read_data_file(self.source2_file, nrows=nrows)
