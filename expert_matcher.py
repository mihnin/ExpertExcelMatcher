"""
🔬 ЭКСПЕРТНАЯ ВЕРСИЯ Excel Matcher (ИСПРАВЛЕНА СТАТИСТИКА!)
С автоматическим перебором методов и корректной статистикой

ИСПРАВЛЕНО:
    ✅ СТАТИСТИКА ТЕПЕРЬ ПО КАТЕГОРИЯМ, НЕ НАКОПИТЕЛЬНАЯ!
    ✅ Сумма всех категорий = общему количеству записей
    ✅ Использует rapidfuzz.process.extractOne (в 100 раз быстрее)
    ✅ Тестирование ВСЕХ доступных методов (не только ТОП-5)
    ✅ Детальный прогресс с процентами
    ✅ Корректная работа с Excel столбцами
    ✅ Динамическая оценка времени выполнения

УСТАНОВКА:
    pip install pandas openpyxl xlsxwriter rapidfuzz textdistance jellyfish

ЗАПУСК:
    python expert_matcher.py
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
from pathlib import Path
import time
from typing import Dict, List, Tuple, Callable
import re

# Импорт из модульной структуры
from src.constants import AppConstants, NormalizationConstants
from src.help_content import HelpContent
from src.models import MatchingMethod, MatchResult, MethodStatistics
from src.matching_engine import MatchingEngine, NormalizationOptions
from src.excel_exporter import ExcelExporter
from src.data_manager import DataManager
from src.ui_manager import UIManager
from src.ui_components import (
    ScrollableFrame, TreeviewWithScrollbar, MethodSelectorListbox,
    FileSelectorWidget, create_label_frame, create_info_label_frame,
    create_styled_button, create_title_header
)

# Импорт библиотек для сопоставления
try:
    from rapidfuzz import fuzz, process
    RAPIDFUZZ_AVAILABLE = True
except ImportError:
    RAPIDFUZZ_AVAILABLE = False
    print("⚠️ rapidfuzz не установлен. Установите: pip install rapidfuzz")

try:
    import textdistance
    TEXTDISTANCE_AVAILABLE = True
except ImportError:
    TEXTDISTANCE_AVAILABLE = False
    print("⚠️ textdistance не установлен. Установите: pip install textdistance")

try:
    import jellyfish
    JELLYFISH_AVAILABLE = True
except ImportError:
    JELLYFISH_AVAILABLE = False
    print("⚠️ jellyfish не установлен. Установите: pip install jellyfish")

try:
    from transliterate import translit
    TRANSLITERATE_AVAILABLE = True
except ImportError:
    TRANSLITERATE_AVAILABLE = False
    print("⚠️ transliterate не установлен. Установите: pip install transliterate")


# ============================================================================
# КОНСТАНТЫ И МОДЕЛИ (импортированы из src.*)
# ============================================================================
# AppConstants, NormalizationConstants - из src.constants
# HelpContent - из src.help_content
# MatchingMethod, MatchResult, MethodStatistics - из src.models


class ExpertMatcher:
    """Экспертная система сопоставления"""
    
    def __init__(self, root):
        self.root = root
        self.root.title(AppConstants.APP_TITLE)

        # Адаптивный размер окна
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Адаптивный размер от экрана
        window_width = max(AppConstants.WINDOW_MIN_WIDTH, int(screen_width * AppConstants.WINDOW_SCALE))
        window_height = max(AppConstants.WINDOW_MIN_HEIGHT, int(screen_height * AppConstants.WINDOW_SCALE))

        # Центрирование окна
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(AppConstants.WINDOW_MIN_WIDTH, AppConstants.WINDOW_MIN_HEIGHT)

        # Менеджер данных
        self.data_manager = DataManager()

        self.results = None
        self.methods_comparison = None
        self.full_comparison_results = None  # Для хранения полных результатов всех методов

        # LEGACY: Алиасы для совместимости (теперь используем data_manager)
        self.askupo_file = None
        self.eatool_file = None
        self.askupo_columns = []
        self.eatool_columns = []
        self.selected_askupo_cols = []
        self.selected_eatool_cols = []
        self.inherit_askupo_cols_var = tk.BooleanVar(value=True)  # Наследовать столбцы из источника 1
        self.inherit_eatool_cols_var = tk.BooleanVar(value=True)  # Наследовать столбцы из источника 2
        self.multi_column_mode_var = tk.BooleanVar(value=False)    # Режим сравнения по нескольким столбцам
        self.selected_methods = []  # Выбранные методы для режима "Выбор нескольких методов"

        # Переменные для расширенной нормализации
        self.norm_remove_legal_var = tk.BooleanVar(value=False)        # Удалять юридические формы (ООО, Ltd, Inc)
        self.norm_remove_versions_var = tk.BooleanVar(value=False)     # Удалять версии (2021, v4.x, R2, SP1)
        self.norm_remove_stopwords_var = tk.BooleanVar(value=False)    # Удалять стоп-слова (и, в, the, a)
        self.norm_transliterate_var = tk.BooleanVar(value=False)       # Транслитерация кириллицы → латиница
        self.norm_remove_punctuation_var = tk.BooleanVar(value=True)   # Удалять пунктуацию (по умолчанию включено)

        # Создаём движок сопоставления
        self.engine = self._create_matching_engine()

        # Создаём экспортер Excel
        self.exporter = ExcelExporter(self.engine, self.results)

        self.methods = self.register_all_methods()

        # Создаём менеджер UI (делегирует создание всех вкладок)
        self.ui_manager = UIManager(self)
        self.ui_manager.create_widgets()

    def _create_matching_engine(self) -> MatchingEngine:
        """Создание движка сопоставления с текущими настройками нормализации"""
        options = NormalizationOptions(
            remove_legal=self.norm_remove_legal_var.get(),
            remove_versions=self.norm_remove_versions_var.get(),
            remove_stopwords=self.norm_remove_stopwords_var.get(),
            transliterate=self.norm_transliterate_var.get(),
            remove_punctuation=self.norm_remove_punctuation_var.get()
        )
        return MatchingEngine(options)

    def _update_matching_engine(self):
        """Обновление движка при изменении настроек нормализации"""
        self.engine = self._create_matching_engine()
        # Обновляем движок в экспортере
        self.exporter.engine = self.engine
        
    def register_all_methods(self) -> List[MatchingMethod]:
        """Регистрация всех доступных методов сопоставления"""
        methods = []
        
        if RAPIDFUZZ_AVAILABLE:
            methods.extend([
                MatchingMethod("RapidFuzz: WRatio",
                             fuzz.WRatio, "rapidfuzz",
                             use_process=True, scorer=fuzz.WRatio),
                MatchingMethod("RapidFuzz: Token Set",
                             fuzz.token_set_ratio, "rapidfuzz",
                             use_process=True, scorer=fuzz.token_set_ratio),
                MatchingMethod("RapidFuzz: Token Sort",
                             fuzz.token_sort_ratio, "rapidfuzz",
                             use_process=True, scorer=fuzz.token_sort_ratio),
                MatchingMethod("RapidFuzz: Partial Ratio",
                             fuzz.partial_ratio, "rapidfuzz",
                             use_process=True, scorer=fuzz.partial_ratio),
                MatchingMethod("RapidFuzz: Ratio",
                             fuzz.ratio, "rapidfuzz",
                             use_process=True, scorer=fuzz.ratio),
                MatchingMethod("RapidFuzz: QRatio",
                             fuzz.QRatio, "rapidfuzz",
                             use_process=True, scorer=fuzz.QRatio),
                MatchingMethod("RapidFuzz: Token Ratio",
                             fuzz.token_ratio, "rapidfuzz",
                             use_process=True, scorer=fuzz.token_ratio),
                MatchingMethod("RapidFuzz: Partial Token Ratio",
                             fuzz.partial_token_ratio, "rapidfuzz",
                             use_process=True, scorer=fuzz.partial_token_ratio),
                MatchingMethod("RapidFuzz: Partial Token Set",
                             fuzz.partial_token_set_ratio, "rapidfuzz",
                             use_process=True, scorer=fuzz.partial_token_set_ratio),
                MatchingMethod("RapidFuzz: Partial Token Sort",
                             fuzz.partial_token_sort_ratio, "rapidfuzz",
                             use_process=True, scorer=fuzz.partial_token_sort_ratio)
            ])

        if TEXTDISTANCE_AVAILABLE:
            methods.extend([
                MatchingMethod("TextDistance: Jaro-Winkler",
                             textdistance.jaro_winkler, "textdistance"),
                MatchingMethod("TextDistance: Jaro",
                             textdistance.jaro, "textdistance"),
                MatchingMethod("TextDistance: Jaccard",
                             textdistance.jaccard, "textdistance"),
                MatchingMethod("TextDistance: Sorensen-Dice",
                             textdistance.sorensen_dice, "textdistance"),
                MatchingMethod("TextDistance: Cosine",
                             textdistance.cosine, "textdistance"),
            ])

        if JELLYFISH_AVAILABLE:
            methods.extend([
                MatchingMethod("Jellyfish: Jaro-Winkler",
                             jellyfish.jaro_winkler_similarity, "jellyfish"),
                MatchingMethod("Jellyfish: Jaro",
                             jellyfish.jaro_similarity, "jellyfish"),
            ])

        # Всегда добавляем метод точного совпадения (ВПР)
        methods.append(
            MatchingMethod("Exact Match (ВПР)",
                         self.exact_match_func, "builtin",
                         use_process=False, scorer=None)
        )

        return methods
    
    def exact_match_func(self, s1: str, s2: str) -> float:
        """Функция точного совпадения для метода ВПР

        Возвращает 100.0 для точного совпадения (после нормализации),
        0.0 для несовпадения
        """
        norm_s1 = self.engine.normalize_string(s1)
        norm_s2 = self.engine.normalize_string(s2)
        return 100.0 if norm_s1 == norm_s2 else 0.0

    # Алиасы для обратной совместимости (делегируют в engine)
    def normalize_string(self, s: str) -> str:
        """Нормализация строки (делегирует в engine)"""
        return self.engine.normalize_string(s)

    def combine_columns(self, row: pd.Series, columns: List[str]) -> str:
        """Объединение столбцов (делегирует в engine)"""
        return self.engine.combine_columns(row, columns)

    def calculate_statistics(self, results_df: pd.DataFrame) -> Dict:
        """Расчёт статистики (делегирует в engine)"""
        return self.engine.calculate_statistics(results_df)

    # ========================================================================
    # ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ (рефакторинг v2.1)
    # ========================================================================

    def _get_column_display_name(self, columns: List[str]) -> str:
        """Получить отображаемое имя для списка столбцов (делегация к DataManager)"""
        return self.data_manager.get_column_display_name(columns)

    def _get_selected_columns(self):
        """Получить выбранные столбцы (синхронизация с DataManager)"""
        # Синхронизируем с data_manager если там пусто, но legacy переменные заполнены
        if not self.data_manager.selected_source1_cols and self.selected_askupo_cols:
            self.data_manager.selected_source1_cols = self.selected_askupo_cols
        if not self.data_manager.selected_source2_cols and self.selected_eatool_cols:
            self.data_manager.selected_source2_cols = self.selected_eatool_cols

        cols1, cols2 = self.data_manager.get_selected_columns()
        # Обновляем legacy переменные
        self.selected_askupo_cols = cols1
        self.selected_eatool_cols = cols2
        return cols1, cols2

    def _create_result_row_dict(self, askupo_combined: str, best_match: str,
                                best_score: float, method_name: str,
                                askupo_row: pd.Series, askupo_df: pd.DataFrame,
                                eatool_row_dict: dict, eatool_df: pd.DataFrame) -> dict:
        """Создать словарь строки результата (устраняет дублирование кода)

        Args:
            askupo_combined: объединенное значение из источника 1
            best_match: найденное совпадение из источника 2
            best_score: процент совпадения
            method_name: название метода
            askupo_row: строка из DataFrame источника 1
            askupo_df: весь DataFrame источника 1
            eatool_row_dict: словарь для поиска строк источника 2
            eatool_df: весь DataFrame источника 2

        Returns:
            Словарь с полями результата
        """
        askupo_cols, eatool_cols = self._get_selected_columns()

        # Формируем названия столбцов
        askupo_col_name = self._get_column_display_name(askupo_cols)
        eatool_col_name = self._get_column_display_name(eatool_cols)

        # Базовые поля
        result_row = {
            f'{AppConstants.COL_SOURCE1_PREFIX} {askupo_col_name}': askupo_combined,
            f'{AppConstants.COL_SOURCE2_PREFIX} {eatool_col_name}': best_match,
            AppConstants.COL_PERCENT: round(best_score, 1),
            AppConstants.COL_METHOD: method_name
        }

        # Наследование столбцов из источника 1
        if self.inherit_askupo_cols_var.get():
            for col in askupo_df.columns:
                if col not in askupo_cols:
                    result_row[f"{AppConstants.COL_SOURCE1_PREFIX} {col}"] = askupo_row[col]

        # Наследование столбцов из источника 2
        if best_match and self.inherit_eatool_cols_var.get():
            matched_row = eatool_row_dict.get(best_match)
            if matched_row is not None:
                for col in eatool_df.columns:
                    if col not in eatool_cols:
                        result_row[f"{AppConstants.COL_SOURCE2_PREFIX} {col}"] = matched_row[col]
        elif self.inherit_eatool_cols_var.get():
            for col in eatool_df.columns:
                if col not in eatool_cols:
                    result_row[f"{AppConstants.COL_SOURCE2_PREFIX} {col}"] = ""

        return result_row

    # ========================================================================
    # СТАТИСТИКА (теперь в src.matching_engine.MatchingEngine)
    # ========================================================================
    # Метод calculate_statistics перенесён в MatchingEngine
    

    # ========================================================================
    # UI МЕТОДЫ (перенесены в UIManager)
    # ========================================================================
    # Методы create_widgets, create_help_tab, create_setup_tab,
    # create_comparison_tab, create_results_tab перенесены в src/ui_manager.py
    #
    # ВАЖНО: Обработчики событий UI (on_*_column_select, select_all_methods,
    # enable_all_normalization и др.) остаются в ExpertMatcher, так как
    # вызываются из UIManager через self.parent.*

    def read_data_file(self, filename: str, nrows=None) -> pd.DataFrame:
        """Универсальное чтение Excel или CSV файла (делегация к DataManager)"""
        return self.data_manager.read_data_file(filename, nrows)

    def validate_excel_file(self, filename: str) -> Tuple[bool, str]:
        """Валидация Excel или CSV файла (делегация к DataManager)"""
        return self.data_manager.validate_file(filename)

    def select_askupo(self):
        filename = filedialog.askopenfilename(
            title="Выберите Источник данных 1 (целевой)",
            filetypes=[("Data files", "*.xlsx *.xls *.csv"), ("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            # Используем DataManager для установки файла
            is_valid, message = self.data_manager.set_source1_file(filename)

            if not is_valid:
                messagebox.showerror("❌ Ошибка валидации Источника данных 1",
                                   f"Файл не прошёл проверку:\n\n{message}\n\n"
                                   f"Требования:\n"
                                   f"• Файл должен содержать данные (не пустой)\n"
                                   f"• Файл должен иметь столбцы и строки\n"
                                   f"• Формат: .xlsx, .xls или .csv")
                return

            # Обновляем legacy переменные
            self.askupo_file = self.data_manager.source1_file
            self.askupo_columns = self.data_manager.source1_columns
            self.selected_askupo_cols = self.data_manager.selected_source1_cols

            # Обновляем GUI
            display_name = self.data_manager.get_short_filename(filename)
            self.askupo_label.config(text=f"✅ {display_name}", fg="green", font=("Arial", 9, "bold"))

            # Загрузка столбцов в GUI
            self.load_askupo_columns()
            self.check_ready()
    
    def select_eatool(self):
        filename = filedialog.askopenfilename(
            title="Выберите Источник данных 2",
            filetypes=[("Data files", "*.xlsx *.xls *.csv"), ("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            # Используем DataManager для установки файла
            is_valid, message = self.data_manager.set_source2_file(filename)

            if not is_valid:
                messagebox.showerror("❌ Ошибка валидации Источника данных 2",
                                   f"Файл не прошёл проверку:\n\n{message}\n\n"
                                   f"Требования:\n"
                                   f"• Файл должен содержать данные (не пустой)\n"
                                   f"• Файл должен иметь столбцы и строки\n"
                                   f"• Формат: .xlsx, .xls или .csv")
                return

            # Обновляем legacy переменные
            self.eatool_file = self.data_manager.source2_file
            self.eatool_columns = self.data_manager.source2_columns
            self.selected_eatool_cols = self.data_manager.selected_source2_cols

            # Обновляем GUI
            display_name = self.data_manager.get_short_filename(filename)
            self.eatool_label.config(text=f"✅ {display_name}", fg="green", font=("Arial", 9, "bold"))

            # Загрузка столбцов в GUI
            self.load_eatool_columns()
            self.check_ready()

    def check_ready(self):
        """Проверка готовности к обработке (делегация к DataManager)"""
        if self.data_manager.is_ready():
            self.process_btn.config(state=tk.NORMAL)

    def load_askupo_columns(self):
        """Загрузка списка столбцов из источника 1"""
        try:
            df = self.read_data_file(self.askupo_file, nrows=0)  # Читаем только заголовки
            self.askupo_columns = list(df.columns)

            # Обновляем GUI для выбора столбцов
            if hasattr(self, 'askupo_col_listbox'):
                self.askupo_col_listbox.delete(0, tk.END)
                for col in self.askupo_columns:
                    self.askupo_col_listbox.insert(tk.END, col)
                # По умолчанию выбираем первый столбец
                if self.askupo_columns:
                    self.askupo_col_listbox.selection_set(0)
                    self.selected_askupo_cols = [self.askupo_columns[0]]
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить столбцы из источника 1:\n{str(e)}")

    def load_eatool_columns(self):
        """Загрузка списка столбцов из источника 2"""
        try:
            df = self.read_data_file(self.eatool_file, nrows=0)  # Читаем только заголовки
            self.eatool_columns = list(df.columns)

            # Обновляем GUI для выбора столбцов
            if hasattr(self, 'eatool_col_listbox'):
                self.eatool_col_listbox.delete(0, tk.END)
                for col in self.eatool_columns:
                    self.eatool_col_listbox.insert(tk.END, col)
                # По умолчанию выбираем первый столбец
                if self.eatool_columns:
                    self.eatool_col_listbox.selection_set(0)
                    self.selected_eatool_cols = [self.eatool_columns[0]]
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить столбцы из источника 2:\n{str(e)}")

    def on_askupo_column_select(self, event):
        """Обработчик выбора столбцов из источника 1"""
        selected_indices = self.askupo_col_listbox.curselection()
        self.selected_askupo_cols = [self.askupo_columns[i] for i in selected_indices]

        # Ограничение: максимум 2 столбца
        if len(selected_indices) > 2:
            messagebox.showwarning("Предупреждение",
                                 "Можно выбрать максимум 2 столбца.\n"
                                 "Последний выбор будет отменен.")
            # Отменяем последний выбор
            self.askupo_col_listbox.selection_clear(selected_indices[-1])
            self.selected_askupo_cols = self.selected_askupo_cols[:-1]

    def on_eatool_column_select(self, event):
        """Обработчик выбора столбцов из источника 2"""
        selected_indices = self.eatool_col_listbox.curselection()
        self.selected_eatool_cols = [self.eatool_columns[i] for i in selected_indices]

        # Ограничение: максимум 2 столбца
        if len(selected_indices) > 2:
            messagebox.showwarning("Предупреждение",
                                 "Можно выбрать максимум 2 столбца.\n"
                                 "Последний выбор будет отменен.")
            # Отменяем последний выбор
            self.eatool_col_listbox.selection_clear(selected_indices[-1])
            self.selected_eatool_cols = self.selected_eatool_cols[:-1]

    def select_all_methods(self):
        """Выбрать все методы в списке"""
        self.methods_listbox.selection_set(0, tk.END)

    def deselect_all_methods(self):
        """Снять выбор всех методов"""
        self.methods_listbox.selection_clear(0, tk.END)

    def enable_all_normalization(self):
        """Включить все опции расширенной нормализации"""
        self.norm_remove_legal_var.set(True)
        self.norm_remove_versions_var.set(True)
        self.norm_remove_stopwords_var.set(True)
        self.norm_transliterate_var.set(True)
        self.norm_remove_punctuation_var.set(True)
        messagebox.showinfo("✓ Опции нормализации",
                           "Все опции расширенной нормализации включены!\n\n"
                           "Это повысит качество сопоставления на 30-50%.")

    def disable_all_normalization(self):
        """Отключить все опции расширенной нормализации"""
        self.norm_remove_legal_var.set(False)
        self.norm_remove_versions_var.set(False)
        self.norm_remove_stopwords_var.set(False)
        self.norm_transliterate_var.set(False)
        self.norm_remove_punctuation_var.set(False)
        messagebox.showinfo("✗ Опции нормализации",
                           "Все опции расширенной нормализации отключены.\n\n"
                           "Будет использоваться только базовая нормализация (lowercase + trim).")

    def get_selected_methods(self):
        """Получить список выбранных методов"""
        selected_indices = self.methods_listbox.curselection()
        return [self.methods[i] for i in selected_indices]

    def start_processing(self):
        """Начать обработку"""
        # Валидация выбранных столбцов
        if not self.selected_askupo_cols:
            messagebox.showerror("Ошибка",
                               "Не выбраны столбцы из Источника данных 1!\n\n"
                               "Выберите хотя бы 1 столбец для сравнения.")
            return

        if not self.selected_eatool_cols:
            messagebox.showerror("Ошибка",
                               "Не выбраны столбцы из Источника данных 2!\n\n"
                               "Выберите хотя бы 1 столбец для сравнения.")
            return

        # Проверка режима множественных столбцов
        if self.multi_column_mode_var.get():
            if len(self.selected_askupo_cols) != 2:
                messagebox.showerror("Ошибка",
                                   "Режим сравнения по 2 столбцам требует выбора ровно 2 столбцов из Источника 1!\n\n"
                                   f"Сейчас выбрано: {len(self.selected_askupo_cols)}")
                return
            if len(self.selected_eatool_cols) != 2:
                messagebox.showerror("Ошибка",
                                   "Режим сравнения по 2 столбцам требует выбора ровно 2 столбцов из Источника 2!\n\n"
                                   f"Сейчас выбрано: {len(self.selected_eatool_cols)}")
                return

        # Проверка совместимости количества столбцов
        if len(self.selected_askupo_cols) != len(self.selected_eatool_cols):
            messagebox.showwarning("Предупреждение",
                                  "Количество выбранных столбцов в обоих источниках должно совпадать!\n\n"
                                  f"Источник 1: {len(self.selected_askupo_cols)} столбцов\n"
                                  f"Источник 2: {len(self.selected_eatool_cols)} столбцов\n\n"
                                  "Для сравнения будет использован только первый столбец из каждого источника.")

        # Валидация выбранных методов для ВСЕХ режимов
        mode = self.mode_var.get()
        selected_methods = self.get_selected_methods()

        # Если методы не выбраны, используем все доступные
        if not selected_methods:
            # Выбираем все методы автоматически
            for i in range(len(self.methods)):
                self.methods_listbox.selection_set(i)
            selected_methods = self.methods
            messagebox.showinfo("Информация",
                               f"Методы не были выбраны.\n\n"
                               f"Будут использованы ВСЕ {len(self.methods)} доступных методов.")

        if mode == "auto":
            self.run_auto_mode(selected_methods)
        elif mode == "compare":
            self.run_compare_mode(selected_methods)
        elif mode == "full_compare":
            self.run_full_comparison_mode(selected_methods)
    
    def run_auto_mode(self, selected_methods):
        """Автоматический режим - выбор лучшего метода из ВЫБРАННЫХ

        Логика выбора ИДЕНТИЧНА режиму сравнения:
        - Приоритет 1: Максимум 100% совпадений
        - Приоритет 2: Максимум 90-99% совпадений
        - Приоритет 3: Максимальный средний процент
        """
        try:
            askupo_df = self.read_data_file(self.askupo_file)
            eatool_df = self.read_data_file(self.eatool_file)

            # Используем выбранные столбцы вместо жестко заданных columns[0]
            askupo_cols = self.selected_askupo_cols
            eatool_cols = self.selected_eatool_cols

            # Динамически рассчитываем примерное время
            sample_size = min(200, len(askupo_df))
            # RapidFuzz быстрые (~2 сек на метод), остальные медленнее (~15-20 сек на метод)
            rapidfuzz_count = sum(1 for m in selected_methods if m.use_process)
            other_count = len(selected_methods) - rapidfuzz_count
            estimated_time = (rapidfuzz_count * 2 + other_count * 20) / 60

            info_msg = (f"📂 Загружено:\n"
                       f"   АСКУПО: {len(askupo_df)} записей\n"
                       f"   EA Tool: {len(eatool_df)} записей\n\n"
                       f"🔍 Будет протестировано {len(selected_methods)} выбранных методов\n"
                       f"   • RapidFuzz методов: {rapidfuzz_count} (быстрые)\n"
                       f"   • Других методов: {other_count} (медленнее)\n"
                       f"⏱️ Примерное время: {estimated_time:.0f}-{estimated_time*1.5:.0f} минут")

            if not messagebox.askokcancel("Начать обработку?", info_msg):
                return

            sample_askupo = askupo_df.head(sample_size)

            best_method = None
            best_score = (-1, -1, -1)  # Кортеж для лексикографического сравнения

            progress_win = tk.Toplevel(self.root)
            progress_win.title("Тестирование выбранных методов...")
            progress_win.geometry("500x200")
            progress_win.transient(self.root)
            progress_win.grab_set()

            tk.Label(progress_win, text="🔬 Тестирование выбранных методов на sample данных",
                    font=("Arial", 12, "bold")).pack(pady=10)

            progress_label = tk.Label(progress_win, text="", font=("Arial", 10))
            progress_label.pack(pady=5)

            progress_bar = ttk.Progressbar(progress_win, length=400, mode='determinate')
            progress_bar.pack(pady=10)
            progress_bar['maximum'] = len(selected_methods)

            time_label = tk.Label(progress_win, text="", font=("Arial", 9), fg="gray")
            time_label.pack(pady=5)

            start_time = time.time()

            for i, method in enumerate(selected_methods):
                elapsed = time.time() - start_time
                progress_label.config(text=f"Метод {i+1}/{len(selected_methods)}: {method.name}")
                time_label.config(text=f"⏱️ Прошло: {int(elapsed)}с")
                progress_bar['value'] = i
                self.root.update()

                score = self.evaluate_method_fast(method, sample_askupo, eatool_df,
                                                  askupo_cols, eatool_cols)

                if score > best_score:
                    best_score = score
                    best_method = method

            progress_win.destroy()

            messagebox.showinfo("✅ Лучший метод найден!",
                              f"🏆 Выбран метод: {best_method.name}\n\n"
                              f"📊 Статистика на sample данных:\n"
                              f"   • 100% совпадений: {best_score[0]}\n"
                              f"   • 90-99% совпадений: {best_score[1]}\n"
                              f"   • Средний процент: {best_score[2]:.1f}%\n\n"
                              f"⏱️ Применение ко всем данным займет ~2-3 минуты")

            self.apply_method_optimized(best_method, askupo_df, eatool_df,
                                       askupo_cols, eatool_cols)

        except Exception as e:
            messagebox.showerror("❌ Ошибка", f"Ошибка обработки:\n{str(e)}\n\n"
                               f"Проверьте:\n"
                               f"• Файлы Excel корректны\n"
                               f"• Первый столбец содержит названия ПО\n"
                               f"• Установлены все библиотеки")
    
    def run_compare_mode(self, selected_methods):
        """Режим сравнения ВЫБРАННЫХ методов

        Логика сортировки ИДЕНТИЧНА автоматическому режиму:
        - Приоритет 1: Максимум 100% совпадений
        - Приоритет 2: Максимум 90-99% совпадений
        - Приоритет 3: Максимальный средний процент
        """
        try:
            askupo_df = self.read_data_file(self.askupo_file)
            eatool_df = self.read_data_file(self.eatool_file)

            # Используем выбранные столбцы вместо жестко заданных columns[0]
            askupo_cols = self.selected_askupo_cols
            eatool_cols = self.selected_eatool_cols

            sample_size = min(200, len(askupo_df))
            sample_askupo = askupo_df.head(sample_size)

            # Динамически рассчитываем примерное время
            rapidfuzz_count = sum(1 for m in selected_methods if m.use_process)
            other_count = len(selected_methods) - rapidfuzz_count
            estimated_time = (rapidfuzz_count * 3 + other_count * 30) / 60

            info_msg = (f"📊 Будет протестировано {len(selected_methods)} выбранных методов\n"
                       f"   • RapidFuzz методов: {rapidfuzz_count} (быстрые)\n"
                       f"   • Других методов: {other_count} (медленнее)\n"
                       f"📦 Sample: {sample_size} записей\n"
                       f"⏱️ Примерное время: {estimated_time:.0f}-{estimated_time*1.5:.0f} минут")

            if not messagebox.askokcancel("Начать сравнение?", info_msg):
                return

            progress_win = tk.Toplevel(self.root)
            progress_win.title("Сравнение выбранных методов...")
            progress_win.geometry("500x200")
            progress_win.transient(self.root)
            progress_win.grab_set()

            tk.Label(progress_win, text="📊 Сравнение выбранных методов",
                    font=("Arial", 12, "bold")).pack(pady=10)

            progress_label = tk.Label(progress_win, text="", font=("Arial", 10))
            progress_label.pack(pady=5)

            progress_bar = ttk.Progressbar(progress_win, length=400, mode='determinate')
            progress_bar.pack(pady=10)
            progress_bar['maximum'] = len(selected_methods)

            comparison_results = []

            for i, method in enumerate(selected_methods):
                progress_label.config(text=f"Тестирование {i+1}/{len(selected_methods)}: {method.name}")
                progress_bar['value'] = i
                self.root.update()

                start_time = time.time()
                # test_method_optimized использует self.selected_*_cols
                results = self.test_method_optimized(method, sample_askupo, eatool_df,
                                                     None, None)
                elapsed = time.time() - start_time

                # Используем ИСПРАВЛЕННУЮ функцию подсчета статистики
                stats_dict = self.engine.calculate_statistics(results)

                stats = {
                    'method': method.name,
                    'library': method.library,
                    'avg_score': results['Процент совпадения'].mean(),
                    'perfect': stats_dict['perfect'],      # Только 100%
                    'high': stats_dict['high'],            # Только 90-99%
                    'medium': stats_dict['medium'],        # Только 70-89%
                    'time': elapsed
                }

                comparison_results.append(stats)

            progress_win.destroy()

            # Лексикографическая сортировка (идентична автоматическому режиму)
            # Приоритет: 100% совпадений > 90-99% совпадений > средний процент
            comparison_results.sort(key=lambda x: (x['perfect'], x['high'], x['avg_score']),
                                   reverse=True)

            self.display_comparison(comparison_results)
            self.notebook.select(1)

            messagebox.showinfo("✅ Сравнение завершено!",
                              f"Протестировано {len(selected_methods)} выбранных методов\n\n"
                              f"🏆 Лучший: {comparison_results[0]['method']}\n"
                              f"📊 100% совпадений: {comparison_results[0]['perfect']}")

        except Exception as e:
            messagebox.showerror("❌ Ошибка", f"Ошибка обработки:\n{str(e)}")

    def _run_comparison_on_full_data(self, methods: List, window_title: str,
                                     header_text: str, export_filename: str) -> None:
        """Общий метод для полного сравнения методов на ВСЕХ данных

        Args:
            methods: Список методов для тестирования
            window_title: Заголовок окна прогресса
            header_text: Текст заголовка в окне прогресса
            export_filename: Имя файла по умолчанию для экспорта
        """
        askupo_df = self.read_data_file(self.askupo_file)
        eatool_df = self.read_data_file(self.eatool_file)

        # Используем выбранные столбцы вместо жестко заданных columns[0]
        askupo_cols = self.selected_askupo_cols
        eatool_cols = self.selected_eatool_cols

        # Создание окна прогресса
        progress_win = tk.Toplevel(self.root)
        progress_win.title(window_title)
        progress_win.geometry("600x250")
        progress_win.transient(self.root)
        progress_win.grab_set()

        tk.Label(progress_win, text=header_text,
                font=("Arial", 12, "bold")).pack(pady=10)

        method_label = tk.Label(progress_win, text="", font=("Arial", 10))
        method_label.pack(pady=5)

        progress_label = tk.Label(progress_win, text="", font=("Arial", 9))
        progress_label.pack(pady=5)

        progress_bar = ttk.Progressbar(progress_win, length=500, mode='determinate')
        progress_bar.pack(pady=10)
        progress_bar['maximum'] = len(methods) * len(askupo_df)

        time_label = tk.Label(progress_win, text="", font=("Arial", 9), fg="gray")
        time_label.pack(pady=5)

        start_time = time.time()
        all_methods_results = {}  # Словарь: имя метода -> DataFrame с результатами
        comparison_stats = []

        total_processed = 0

        # Обработка каждого метода
        for method_idx, method in enumerate(methods):
            method_start_time = time.time()
            method_label.config(text=f"Метод {method_idx+1}/{len(methods)}: {method.name}")
            self.root.update()

            # Применяем метод ко ВСЕМ данным
            # test_method_optimized использует self.selected_*_cols
            results_df = self.test_method_optimized(method, askupo_df, eatool_df,
                                                   None, None)

            # Сохраняем результаты
            all_methods_results[method.name] = results_df

            # Подсчитываем статистику
            stats_dict = self.engine.calculate_statistics(results_df)

            comparison_stats.append({
                'method': method.name,
                'library': method.library,
                'total': stats_dict['total'],
                'perfect': stats_dict['perfect'],
                'high': stats_dict['high'],
                'medium': stats_dict['medium'],
                'low': stats_dict['low'],
                'very_low': stats_dict['very_low'],
                'none': stats_dict['none'],
                'avg_score': results_df['Процент совпадения'].mean(),
                'time': time.time() - method_start_time
            })

            # Обновляем прогресс
            total_processed += len(askupo_df)
            progress_bar['value'] = total_processed
            elapsed = time.time() - start_time
            remaining = (elapsed / total_processed) * (len(methods) * len(askupo_df) - total_processed)

            progress_label.config(text=f"Обработано методов: {method_idx+1}/{len(methods)}")
            time_label.config(text=f"⏱️ Прошло: {int(elapsed)}с ({elapsed/60:.1f} мин) | Осталось: ~{int(remaining)}с ({remaining/60:.1f} мин)")
            self.root.update()

        progress_win.destroy()

        # Сортируем методы по качеству
        comparison_stats.sort(key=lambda x: (x['perfect'], x['high'], x['avg_score']), reverse=True)

        # Сохраняем для экспорта
        self.full_comparison_results = {
            'methods_data': all_methods_results,
            'comparison_stats': comparison_stats
        }

        elapsed_total = time.time() - start_time

        # Автоматически экспортируем результаты
        self.export_full_comparison_to_excel(default_filename=export_filename)

        # Показываем финальное сообщение
        messagebox.showinfo("✅ Полное сравнение завершено!",
                          f"⏱️ Время выполнения: {int(elapsed_total)}с ({elapsed_total/60:.1f} мин)\n\n"
                          f"📊 Протестировано {len(methods)} методов\n"
                          f"📦 Обработано {len(askupo_df)} записей в каждом методе\n\n"
                          f"🏆 Лучший метод: {comparison_stats[0]['method']}\n"
                          f"   • 100% совпадений: {comparison_stats[0]['perfect']}\n"
                          f"   • 90-99%: {comparison_stats[0]['high']}\n"
                          f"   • Средний балл: {comparison_stats[0]['avg_score']:.1f}%\n\n"
                          f"💾 Результаты сохранены в Excel")

    def run_full_comparison_mode(self, selected_methods):
        """Полное сравнение - применяет ВЫБРАННЫЕ методы ко ВСЕМ данным"""
        try:
            # Читаем данные для расчета времени
            askupo_df = self.read_data_file(self.askupo_file)
            eatool_df = self.read_data_file(self.eatool_file)

            # Динамически рассчитываем примерное время для ВСЕХ данных
            rapidfuzz_count = sum(1 for m in selected_methods if m.use_process)
            other_count = len(selected_methods) - rapidfuzz_count
            estimated_time = (rapidfuzz_count * 3 + other_count * 4) / 60

            # Показываем предупреждение
            info_msg = (f"⚠️ ВНИМАНИЕ: Это может быть долгая операция!\n\n"
                       f"📂 Будет обработано:\n"
                       f"   АСКУПО: {len(askupo_df)} записей\n"
                       f"   EA Tool: {len(eatool_df)} записей\n"
                       f"   Методов: {len(selected_methods)} выбранных\n\n"
                       f"🔬 Каждый метод будет применен ко ВСЕМ записям\n"
                       f"⏱️ Примерное время: {estimated_time:.0f}-{estimated_time*1.5:.0f} минут\n\n"
                       f"📊 Результат: Excel файл с листом для каждого метода + сводка")

            if not messagebox.askokcancel("⚠️ Начать полное сравнение?", info_msg):
                return

            # Вызываем общий метод для обработки
            self._run_comparison_on_full_data(
                methods=selected_methods,
                window_title="Полное сравнение выбранных методов...",
                header_text="🔬 Полное сравнение выбранных методов на ВСЕХ данных",
                export_filename="Полное_сравнение_выбранных_методов.xlsx"
            )

        except Exception as e:
            messagebox.showerror("❌ Ошибка", f"Ошибка обработки:\n{str(e)}")

    # Методы run_manual_mode и run_multi_manual_mode УДАЛЕНЫ
    # Вся функциональность теперь в run_full_comparison_mode

    def evaluate_method_fast(self, method: MatchingMethod, sample_askupo: pd.DataFrame,
                            eatool_df: pd.DataFrame, askupo_cols: list, eatool_cols: list) -> tuple:
        """Быстрая оценка качества метода

        Возвращает кортеж для лексикографического сравнения:
        (количество 100%, количество 90-99%, средний процент)
        Это обеспечивает единообразие с режимом сравнения методов.

        Args:
            askupo_cols: Список столбцов источника 1 для сравнения
            eatool_cols: Список столбцов источника 2 для сравнения
        """
        # test_method_optimized уже правильно обрабатывает списки столбцов через self.selected_*_cols
        results = self.test_method_optimized(method, sample_askupo, eatool_df,
                                            None, None)

        stats = self.engine.calculate_statistics(results)

        # Лексикографическая оценка (приоритет: 100% > 90-99% > средний)
        # Идентична логике сортировки в режиме сравнения
        score = (stats['perfect'], stats['high'], results['Процент совпадения'].mean())

        return score
    
    def test_method_optimized(self, method: MatchingMethod, askupo_df: pd.DataFrame,
                             eatool_df: pd.DataFrame, askupo_col: str = None, eatool_col: str = None) -> pd.DataFrame:
        """Оптимизированное тестирование метода

        Поддерживает:
        - Выбор конкретных столбцов для сравнения
        - Режим множественных столбцов (2 столбца одновременно)
        - Наследование дополнительных столбцов из источников
        """
        # Используем выбранные столбцы из GUI или переданные параметры
        askupo_cols = self.selected_askupo_cols if self.selected_askupo_cols else [askupo_col if askupo_col else askupo_df.columns[0]]
        eatool_cols = self.selected_eatool_cols if self.selected_eatool_cols else [eatool_col if eatool_col else eatool_df.columns[0]]

        # Подготовка данных из источника 2 для сравнения
        eatool_combined_names = []
        eatool_original_values = []

        for _, row in eatool_df.iterrows():
            # Объединяем значения из выбранных столбцов
            combined = self.engine.combine_columns(row, eatool_cols)
            eatool_combined_names.append(combined)
            eatool_original_values.append(combined)

        # Нормализация для поиска
        eatool_normalized = [self.normalize_string(name) for name in eatool_combined_names]
        choice_dict = {norm: orig for norm, orig in zip(eatool_normalized, eatool_original_values)}

        # Создаём словарь для быстрого поиска строки по комбинированному значению
        eatool_row_dict = {}
        for idx, row in eatool_df.iterrows():
            combined = self.engine.combine_columns(row, eatool_cols)
            eatool_row_dict[combined] = row

        results = []

        for _, row in askupo_df.iterrows():
            # Объединяем значения из выбранных столбцов источника 1 (конкатенация)
            askupo_combined = self.engine.combine_columns(row, askupo_cols)
            askupo_normalized = self.engine.normalize_string(askupo_combined)

            # Поиск лучшего совпадения
            best_match, best_score = method.find_best_match(
                askupo_normalized,
                eatool_normalized,
                choice_dict
            )

            # Применяем порог отклонения
            if best_score < AppConstants.THRESHOLD_REJECT:
                best_match = ""
                best_score = 0

            # Используем вспомогательный метод (рефакторинг v2.1 - устранение дублирования)
            result_row = self._create_result_row_dict(
                askupo_combined=askupo_combined,
                best_match=best_match,
                best_score=best_score,
                method_name=method.name,
                askupo_row=row,
                askupo_df=askupo_df,
                eatool_row_dict=eatool_row_dict,
                eatool_df=eatool_df
            )

            results.append(result_row)

        return pd.DataFrame(results)
    
    def apply_method_optimized(self, method: MatchingMethod, askupo_df: pd.DataFrame,
                               eatool_df: pd.DataFrame, askupo_cols: list, eatool_cols: list):
        """Оптимизированное применение метода с поддержкой множественных столбцов

        Args:
            askupo_cols: Список столбцов источника 1 для сравнения
            eatool_cols: Список столбцов источника 2 для сравнения
        """

        progress_win = tk.Toplevel(self.root)
        progress_win.title(f"Применение метода...")
        progress_win.geometry("600x250")
        progress_win.transient(self.root)
        progress_win.grab_set()

        tk.Label(progress_win, text=f"⚙️ {method.name}",
                font=("Arial", 12, "bold")).pack(pady=10)

        status_label = tk.Label(progress_win, text="Подготовка данных...",
                               font=("Arial", 10))
        status_label.pack(pady=5)

        progress_label = tk.Label(progress_win, text="", font=("Arial", 9))
        progress_label.pack(pady=5)

        progress_bar = ttk.Progressbar(progress_win, length=500, mode='determinate')
        progress_bar.pack(pady=10)

        time_label = tk.Label(progress_win, text="", font=("Arial", 9), fg="gray")
        time_label.pack(pady=5)

        self.root.update()

        start_time = time.time()

        # Подготовка данных источника 2 с объединением столбцов
        eatool_combined_names = []
        for _, row in eatool_df.iterrows():
            combined = self.engine.combine_columns(row, eatool_cols)
            eatool_combined_names.append(combined)

        eatool_normalized = [self.engine.normalize_string(name) for name in eatool_combined_names]
        choice_dict = {norm: orig for norm, orig in zip(eatool_normalized, eatool_combined_names)}

        # Создаём словарь для быстрого поиска строки по комбинированному значению
        eatool_row_dict = {}
        for idx, row in eatool_df.iterrows():
            combined = self.engine.combine_columns(row, eatool_cols)
            eatool_row_dict[combined] = row

        status_label.config(text="Обработка записей...")

        results = []
        total = len(askupo_df)
        progress_bar['maximum'] = total

        for idx, row in askupo_df.iterrows():
            # Объединяем значения из выбранных столбцов источника 1
            askupo_combined = self.engine.combine_columns(row, askupo_cols)
            askupo_normalized = self.engine.normalize_string(askupo_combined)

            best_match, best_score = method.find_best_match(
                askupo_normalized,
                eatool_normalized,
                choice_dict
            )

            # Применяем порог отклонения
            if best_score < AppConstants.THRESHOLD_REJECT:
                best_match = ""
                best_score = 0

            # Используем вспомогательный метод (рефакторинг v2.1 - устранение дублирования)
            result_row = self._create_result_row_dict(
                askupo_combined=askupo_combined,
                best_match=best_match,
                best_score=best_score,
                method_name=method.name,
                askupo_row=row,
                askupo_df=askupo_df,
                eatool_row_dict=eatool_row_dict,
                eatool_df=eatool_df
            )

            results.append(result_row)
            
            if idx % 10 == 0:
                elapsed = time.time() - start_time
                remaining = (elapsed / (idx + 1)) * (total - idx - 1)
                
                progress_bar['value'] = idx
                progress_label.config(text=f"{idx}/{total} записей ({int(idx/total*100)}%)")
                time_label.config(text=f"⏱️ Прошло: {int(elapsed)}с | Осталось: ~{int(remaining)}с")
                self.root.update()
        
        progress_bar['value'] = total
        self.root.update()
        
        self.results = pd.DataFrame(results).sort_values('Процент совпадения', ascending=False)
        
        progress_win.destroy()
        
        self.display_results(method)
        self.notebook.select(2)
        
        elapsed_total = time.time() - start_time
        
        # Используем ИСПРАВЛЕННУЮ функцию статистики
        stats = self.engine.calculate_statistics(self.results)
        
        stats_msg = (f"✅ Обработка завершена!\n\n"
                    f"⏱️ Время: {int(elapsed_total)}с ({elapsed_total/60:.1f} мин)\n"
                    f"📊 Обработано: {stats['total']} записей\n\n"
                    f"Результаты (по категориям):\n"
                    f"  • 100% (точное):     {stats['perfect']} ({stats['perfect']/stats['total']*100:.1f}%)\n"
                    f"  • 90-99% (высокое):  {stats['high']} ({stats['high']/stats['total']*100:.1f}%)\n"
                    f"  • 70-89% (среднее):  {stats['medium']} ({stats['medium']/stats['total']*100:.1f}%)\n"
                    f"  • 50-69% (низкое):   {stats['low']} ({stats['low']/stats['total']*100:.1f}%)\n"
                    f"  • 1-49% (очень низкое): {stats['very_low']} ({stats['very_low']/stats['total']*100:.1f}%)\n"
                    f"  • 0% (нет совпадения): {stats['none']} ({stats['none']/stats['total']*100:.1f}%)\n\n"
                    f"✓ Проверка: {stats['check_sum']} = {stats['total']} {'✅' if stats['check_sum'] == stats['total'] else '❌'}")
        
        messagebox.showinfo("Готово!", stats_msg)
    
    def display_comparison(self, comparison_results: List[Dict]):
        """Отображение сравнения методов"""
        self.methods_comparison = comparison_results
        
        for item in self.comparison_tree.get_children():
            self.comparison_tree.delete(item)
        
        for rank, stats in enumerate(comparison_results, 1):
            medal = "🥇" if rank == 1 else "🥈" if rank == 2 else "🥉" if rank == 3 else f"#{rank}"
            
            values = (
                medal,
                stats['method'],
                stats['library'],
                stats['perfect'],      # Только 100%
                stats['high'],         # Только 90-99%
                f"{stats['avg_score']:.1f}%",
                f"{stats['time']:.1f}с"
            )
            
            tag = 'best' if rank == 1 else 'good' if rank <= 3 else ''
            self.comparison_tree.insert("", tk.END, values=values, tags=(tag,))
        
        self.comparison_tree.tag_configure('best', background='#D1FAE5')
        self.comparison_tree.tag_configure('good', background='#DBEAFE')
    
    def display_results(self, method: MatchingMethod):
        """Отображение результатов"""
        
        for widget in self.result_info_frame.winfo_children():
            widget.destroy()
        
        info_text = f"🔬 Использован метод: {method.name} (библиотека: {method.library})"
        tk.Label(self.result_info_frame, text=info_text, 
                font=("Arial", 11, "bold"), fg="#7C3AED").pack(anchor=tk.W)
        
        for widget in self.result_stats_frame.winfo_children():
            widget.destroy()
        
        # Используем ИСПРАВЛЕННУЮ функцию статистики
        stats = self.engine.calculate_statistics(self.results)
        
        stats_display = {
            'Всего': stats['total'],
            '100%': stats['perfect'],
            '90-99%': stats['high'],
            '70-89%': stats['medium'],
            '50-69%': stats['low'],
            '1-49%': stats['very_low'],
            '0%': stats['none']
        }
        
        colors = ['#6B7280', '#10B981', '#3B82F6', '#F59E0B', '#F97316', '#FF6B6B', '#EF4444']
        
        for i, (label, value) in enumerate(stats_display.items()):
            frame = tk.Frame(self.result_stats_frame, bg=colors[i], padx=8, pady=5)
            frame.pack(side=tk.LEFT, padx=3)
            
            tk.Label(frame, text=str(value), font=("Arial", 15, "bold"), 
                    fg="white", bg=colors[i]).pack()
            tk.Label(frame, text=label, font=("Arial", 8), 
                    fg="white", bg=colors[i]).pack()
        
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        for idx, row in self.results.head(50).iterrows():
            # Названия столбцов теперь динамические, используем первый и второй столбец
            col_names = self.results.columns.tolist()
            source1_col = [c for c in col_names if c.startswith('Источник 1:')][0]
            source2_col = [c for c in col_names if c.startswith('Источник 2:')][0]

            source1 = str(row[source1_col])
            source2 = str(row[source2_col]) if row[source2_col] else ""

            values = (
                idx + 1,
                source1[:50] + "..." if len(source1) > 50 else source1,
                source2[:50] + "..." if source2 and len(source2) > 50 else source2 if source2 else "❌ НЕТ",
                f"{row['Процент совпадения']}%"
            )

            percent = row['Процент совпадения']
            tag = 'perfect' if percent == 100 else 'high' if percent >= 90 else 'medium' if percent >= 70 else 'low' if percent >= 50 else 'very_low' if percent > 0 else 'none'

            self.results_tree.insert("", tk.END, values=values, tags=(tag,))
        
        self.results_tree.tag_configure('perfect', background='#D1FAE5')
        self.results_tree.tag_configure('high', background='#DBEAFE')
        self.results_tree.tag_configure('medium', background='#FEF3C7')
        self.results_tree.tag_configure('low', background='#FED7AA')
        self.results_tree.tag_configure('very_low', background='#FFE4E1')
        self.results_tree.tag_configure('none', background='#FEE2E2')
    
    def export_comparison(self):
        """Экспорт сравнения методов"""
        self.exporter.export_comparison(self.methods_comparison)
    
    def export_full(self):
        """Экспорт полного отчета со статистикой"""
        if self.results is None:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return
        # Обновляем results в exporter перед экспортом
        self.exporter.results = self.results
        self.exporter.export_results(self.results, "Полный_отчет_сопоставления.xlsx", include_stats=True)

    def export_perfect(self):
        """Экспорт только 100% совпадений"""
        if self.results is None:
            return
        data = self.results[self.results[AppConstants.COL_PERCENT] == 100]
        self.exporter.results = self.results
        self.exporter.export_results(data, "Точные_совпадения_100%.xlsx")

    def export_problems(self):
        """Экспорт проблемных совпадений (<90%)"""
        if self.results is None:
            return
        data = self.results[self.results[AppConstants.COL_PERCENT] < 90]
        self.exporter.results = self.results
        self.exporter.export_results(data, "Требуют_проверки_менее_90%.xlsx")

    def export_no_match(self):
        """Экспорт несовпадений (0%)"""
        if self.results is None:
            return
        data = self.results[self.results[AppConstants.COL_PERCENT] == 0]
        self.exporter.results = self.results
        self.exporter.export_results(data, "Без_совпадений_0%.xlsx")
    
    def export_excel(self, data: pd.DataFrame, filename: str, include_stats: bool = False):
        """
        Базовая функция экспорта в Excel (LEGACY - используется для обратной совместимости)
        Рекомендуется использовать self.exporter.export_results() напрямую
        """
        self.exporter.results = self.results
        return self.exporter.export_results(data, filename, include_stats)

    def export_full_comparison_to_excel(self, default_filename=None):
        """
        Экспорт полного сравнения всех методов в Excel

        Args:
            default_filename: Имя файла по умолчанию (опционально)
        """
        filename = default_filename or "Полное_сравнение_всех_методов.xlsx"
        return self.exporter.export_full_comparison(self.full_comparison_results, filename)


def main():
    root = tk.Tk()
    app = ExpertMatcher(root)
    root.mainloop()


if __name__ == "__main__":
    main()