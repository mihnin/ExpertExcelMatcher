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


# ============================================================================
# КОНСТАНТЫ
# ============================================================================

class AppConstants:
    """Константы приложения"""
    # Версия
    VERSION = "2.2.0"
    APP_TITLE = f"🔬 Expert Excel Matcher v{VERSION}"

    # Названия столбцов в результатах
    COL_SOURCE1_PREFIX = "Источник 1:"
    COL_SOURCE2_PREFIX = "Источник 2:"
    COL_PERCENT = "Процент совпадения"
    COL_METHOD = "Метод"

    # Пороги совпадения
    THRESHOLD_PERFECT = 100
    THRESHOLD_HIGH = 90
    THRESHOLD_MEDIUM = 70
    THRESHOLD_LOW = 50
    THRESHOLD_REJECT = 50  # Ниже этого порога - отклоняем

    # UI константы
    WINDOW_MIN_WIDTH = 1000
    WINDOW_MIN_HEIGHT = 700
    WINDOW_SCALE = 0.8  # 80% от размера экрана

    # Размеры sample для тестирования
    SAMPLE_SIZE = 200


# ============================================================================
# КЛАССЫ
# ============================================================================

class MatchingMethod:
    """Класс для описания метода сопоставления"""

    def __init__(self, name: str, func: Callable, library: str,
                 use_process: bool = False, scorer=None):
        self.name = name
        self.func = func
        self.library = library
        self.use_process = use_process
        self.scorer = scorer
        
    def find_best_match(self, query: str, choices: List[str],
                       choice_dict: Dict[str, str]) -> Tuple[str, float]:
        """Поиск лучшего совпадения с учетом длины строк"""
        if not query or not choices:
            return "", 0.0

        try:
            query_len = len(query)

            if self.use_process and RAPIDFUZZ_AVAILABLE:
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
                best_match = ""
                best_score = 0.0

                for choice in choices:
                    try:
                        score = self.func(query, choice)
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
                    except:
                        continue

                return best_match, best_score
        except Exception as e:
            return "", 0.0


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

        self.askupo_file = None
        self.eatool_file = None
        self.results = None
        self.methods_comparison = None
        self.full_comparison_results = None  # Для хранения полных результатов всех методов

        # Новые переменные для работы со столбцами
        self.askupo_columns = []  # Список всех столбцов из источника 1
        self.eatool_columns = []  # Список всех столбцов из источника 2
        self.selected_askupo_cols = []  # Выбранные столбцы для сравнения из источника 1
        self.selected_eatool_cols = []  # Выбранные столбцы для сравнения из источника 2
        self.inherit_askupo_cols_var = tk.BooleanVar(value=True)  # Наследовать столбцы из источника 1
        self.inherit_eatool_cols_var = tk.BooleanVar(value=True)  # Наследовать столбцы из источника 2
        self.multi_column_mode_var = tk.BooleanVar(value=False)    # Режим сравнения по нескольким столбцам
        self.selected_methods = []  # Выбранные методы для режима "Выбор нескольких методов"

        self.methods = self.register_all_methods()

        self.create_widgets()
        
    def register_all_methods(self) -> List[MatchingMethod]:
        """Регистрация всех доступных методов сопоставления"""
        methods = []
        
        if RAPIDFUZZ_AVAILABLE:
            methods.extend([
                MatchingMethod("🥇 RapidFuzz: WRatio (рекомендуется)",
                             fuzz.WRatio, "rapidfuzz",
                             use_process=True, scorer=fuzz.WRatio),
                MatchingMethod("🥈 RapidFuzz: Token Set",
                             fuzz.token_set_ratio, "rapidfuzz",
                             use_process=True, scorer=fuzz.token_set_ratio),
                MatchingMethod("🥉 RapidFuzz: Token Sort",
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
            MatchingMethod("📊 Exact Match (ВПР)",
                         self.exact_match_func, "builtin",
                         use_process=False, scorer=None)
        )

        return methods
    
    def exact_match_func(self, s1: str, s2: str) -> float:
        """Функция точного совпадения для метода ВПР

        Возвращает 100.0 для точного совпадения (после нормализации),
        0.0 для несовпадения
        """
        norm_s1 = self.normalize_string(s1)
        norm_s2 = self.normalize_string(s2)
        return 100.0 if norm_s1 == norm_s2 else 0.0

    def normalize_string(self, s: str) -> str:
        """Нормализация строки"""
        if not s or pd.isna(s):
            return ""
        s = str(s).lower().strip()
        s = re.sub(r'\s+', ' ', s)
        return s

    def combine_columns(self, row: pd.Series, columns: List[str]) -> str:
        """Объединение значений из нескольких столбцов в одну строку

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

    # ========================================================================
    # ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ (рефакторинг v2.1)
    # ========================================================================

    def _get_column_display_name(self, columns: List[str]) -> str:
        """Получить отображаемое имя для списка столбцов

        Args:
            columns: список имен столбцов

        Returns:
            Строка вида "Col1" или "Col1 + Col2"
        """
        return " + ".join(columns) if len(columns) > 1 else columns[0]

    def _get_selected_columns(self):
        """Получить выбранные столбцы или дефолтные

        Returns:
            Tuple[List[str], List[str]]: (askupo_cols, eatool_cols)
        """
        askupo_cols = self.selected_askupo_cols if self.selected_askupo_cols else []
        eatool_cols = self.selected_eatool_cols if self.selected_eatool_cols else []
        return askupo_cols, eatool_cols

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
    # СТАТИСТИКА
    # ========================================================================

    def calculate_statistics(self, results_df: pd.DataFrame) -> Dict:
        """
        ИСПРАВЛЕННАЯ функция подсчета статистики!
        Теперь считает по КАТЕГОРИЯМ, а не накопительно!
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
    
    def create_widgets(self):
        """Создание интерфейса"""
        
        title_frame = tk.Frame(self.root, bg="#7C3AED", pady=15)
        title_frame.pack(fill=tk.X)
        
        tk.Label(
            title_frame,
            text="🔬 Expert Excel Matcher v1.0",
            font=("Arial", 18, "bold"),
            fg="white",
            bg="#7C3AED"
        ).pack()
        
        tk.Label(
            title_frame,
            text=f"⚡ В 100 раз быстрее! • {len(self.methods)} методов • Корректная статистика",
            font=("Arial", 10),
            fg="white",
            bg="#7C3AED"
        ).pack()
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.setup_tab = tk.Frame(self.notebook)
        self.notebook.add(self.setup_tab, text="1. Загрузка и настройка")
        self.create_setup_tab()
        
        self.comparison_tab = tk.Frame(self.notebook)
        self.notebook.add(self.comparison_tab, text="2. Сравнение методов")
        self.create_comparison_tab()
        
        self.results_tab = tk.Frame(self.notebook)
        self.notebook.add(self.results_tab, text="3. Результаты")
        self.create_results_tab()

        self.help_tab = tk.Frame(self.notebook)
        self.notebook.add(self.help_tab, text="❓ Справка")
        self.create_help_tab()

    def create_help_tab(self):
        """Вкладка справки"""
        main_frame = tk.Frame(self.help_tab, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок
        tk.Label(main_frame, text="📖 Справка по работе с приложением",
                font=("Arial", 14, "bold"), fg="#7C3AED").pack(pady=(0, 15))

        # Создаём canvas для прокрутки
        canvas = tk.Canvas(main_frame, bg="white")
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Раздел 1: Требования к файлам
        section1 = tk.LabelFrame(scrollable_frame, text="📂 Требования к входным файлам (Excel/CSV)",
                                font=("Arial", 11, "bold"), padx=15, pady=10, bg="white")
        section1.pack(fill=tk.X, pady=(0, 15))

        help_text_files = """
✅ ОБЩИЕ ТРЕБОВАНИЯ:
   • Формат файлов: .xlsx, .xls или .csv (NEW в v2.2!)
   • Кодировка CSV: автоматическое определение (UTF-8-BOM, UTF-8, CP1251, Windows-1251, Latin1)
   • Разделитель CSV: запятая (стандарт)
   • Название файла: ЛЮБОЕ (без ограничений)

✅ СТРУКТУРА ФАЙЛОВ (v2.2 - гибкий выбор столбцов):
   • Можно выбрать ЛЮБОЙ столбец(цы) из каждого источника для сравнения
   • Поддержка сравнения по 1 или 2 столбцам одновременно
   • Чекбоксы для наследования остальных столбцов в результат
   • По умолчанию: оба источника наследуют все столбцы

📋 ПРИМЕРЫ ДОПУСТИМЫХ СТРУКТУР:

   Файл 1 (АСКУПО):
   ┌────────────────────────────┬──────────┬─────────┐
   │ Название ПО                │ Версия   │ Vendor  │
   ├────────────────────────────┼──────────┼─────────┤
   │ Microsoft Office 365       │ 2021     │ MS      │
   │ Adobe Acrobat Reader DC    │ 22.0     │ Adobe   │
   └────────────────────────────┴──────────┴─────────┘

   Файл 2 (EA Tool):
   ┌────────────────────────────┬──────────┐
   │ Product Name               │ Category │
   ├────────────────────────────┼──────────┤
   │ MS Office 365              │ Office   │
   │ Acrobat Reader             │ PDF      │
   └────────────────────────────┴──────────┘

⚠️ ВАЖНО:
   • Выбранные столбцы могут содержать ЛЮБЫЕ данные (текст, числа, даты)
   • Все данные автоматически преобразуются в текст для сравнения
   • Пустые строки в выбранных столбцах будут пропущены
   • Регистр букв не важен (всё приводится к нижнему регистру)
   • При выборе 2 столбцов их значения объединяются через ПРОБЕЛ в единую строку
   • Сравнение идет по объединенной строке как единому целому
"""

        tk.Label(section1, text=help_text_files, font=("Consolas", 9),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        # Раздел 2: Режимы работы
        section2 = tk.LabelFrame(scrollable_frame, text="⚙️ Режимы работы приложения",
                                font=("Arial", 11, "bold"), padx=15, pady=10, bg="white")
        section2.pack(fill=tk.X, pady=(0, 15))

        help_text_modes = f"""
1️⃣ АВТОМАТИЧЕСКИЙ РЕЖИМ (рекомендуется):
   • Тестирует ВЫБРАННЫЕ методы на образце данных (~200 записей)
   • Автоматически выбирает лучший метод из выбранных
   • Применяет его ко всем данным и создает результирующий Excel
   • Выбор методов: удерживайте Ctrl для множественного выбора
   • Если методы не выбраны - используются ВСЕ {len(self.methods)} доступных
   • Время: зависит от количества выбранных методов

2️⃣ РЕЖИМ СРАВНЕНИЯ (для анализа):
   • Тестирует ВЫБРАННЫЕ методы на выборке (не более первых 200 записей)
   • Показывает статистику по каждому выбранному методу
   • Выбирает лучший метод по критериям (100% > 90-99% > средний балл)
   • Позволяет увидеть и сравнить результаты перед применением
   • Если методы не выбраны - используются ВСЕ {len(self.methods)} доступных
   • Время: зависит от количества выбранных методов

3️⃣ ПОЛНОЕ СРАВНЕНИЕ (может быть долгая операция):
   • Выбор одного или нескольких методов для создания результирующего Excel
   • Применяет КАЖДЫЙ выбранный метод ко ВСЕМ данным
   • Создаёт Excel файл с листом для каждого метода + сводная таблица
   • Выбор методов: удерживайте Ctrl для множественного выбора
   • Если методы не выбраны - используются ВСЕ {len(self.methods)} доступных
   • Время: зависит от количества выбранных методов и объёма данных

💡 СОВЕТ: Выбирайте только нужные методы для ускорения работы!
"""

        tk.Label(section2, text=help_text_modes, font=("Consolas", 9),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        # Раздел 2.5: Алгоритм выбора лучшего метода (NEW)
        section2_5 = tk.LabelFrame(scrollable_frame, text="🧮 Алгоритм автоматического выбора метода (NEW в v2.2)",
                                font=("Arial", 11, "bold"), padx=15, pady=10, bg="white")
        section2_5.pack(fill=tk.X, pady=(0, 15))

        help_text_algorithm = """
📊 КАК РАБОТАЕТ АВТОМАТИЧЕСКИЙ ВЫБОР:

Автоматический режим и режим Сравнения используют ОДИНАКОВЫЙ алгоритм для
определения лучшего метода. Это гарантирует, что автоматический режим всегда
выбирает тот же метод, который будет показан как #1 в режиме Сравнения.

🔢 ЛЕКСИКОГРАФИЧЕСКАЯ СОРТИРОВКА (Lexicographic Sorting):

Каждый метод получает 3 показателя после тестирования на образце данных:
   1. Количество совпадений 100% (идеальные совпадения)
   2. Количество совпадений 90-99% (высокое качество)
   3. Средний процент совпадения по всем записям

Методы сравниваются в следующем порядке приоритетов:

   ┌────────────────────────────────────────────────────────────┐
   │ ПРИОРИТЕТ 1: Максимум идеальных совпадений (100%)          │
   │ Чем больше точных совпадений - тем лучше метод             │
   └────────────────────────────────────────────────────────────┘
                            ↓ если равны
   ┌────────────────────────────────────────────────────────────┐
   │ ПРИОРИТЕТ 2: Максимум высоких совпадений (90-99%)          │
   │ Чем больше почти точных совпадений - тем лучше метод       │
   └────────────────────────────────────────────────────────────┘
                            ↓ если равны
   ┌────────────────────────────────────────────────────────────┐
   │ ПРИОРИТЕТ 3: Максимальный средний процент совпадения       │
   │ Общее качество сопоставления всех записей                  │
   └────────────────────────────────────────────────────────────┘

📈 ПРИМЕР СРАВНЕНИЯ:

   Метод A: (50 × 100%, 30 × 90-99%, avg=85%)
   Метод B: (48 × 100%, 40 × 90-99%, avg=90%)

   ПОБЕДИТЕЛЬ: Метод A
   ПРИЧИНА: 50 > 48 в первом приоритете (100% совпадений)

   Метод C: (50 × 100%, 35 × 90-99%, avg=82%)
   Метод D: (50 × 100%, 35 × 90-99%, avg=87%)

   ПОБЕДИТЕЛЬ: Метод D
   ПРИЧИНА: Равны в приоритетах 1 и 2, но 87% > 82% в приоритете 3

✅ ПОЧЕМУ ТАКОЙ АЛГОРИТМ:

   • Идеальные совпадения важнее всего - их можно использовать сразу
   • Высокие совпадения (90-99%) требуют минимальной проверки
   • Средний процент показывает общее качество, но менее важен
   • Лексикографический порядок гарантирует стабильность выбора
"""

        tk.Label(section2_5, text=help_text_algorithm, font=("Consolas", 8),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        # Раздел 3: Выбор столбцов (NEW в v2.2)
        section3 = tk.LabelFrame(scrollable_frame, text="🎯 Выбор столбцов для сравнения (NEW в v2.2)",
                                font=("Arial", 11, "bold"), padx=15, pady=10, bg="white")
        section3.pack(fill=tk.X, pady=(0, 15))

        help_text_columns = """
🔹 ВЫБОР СТОЛБЦОВ:
   • После загрузки каждого файла появляется список его столбцов
   • Можно выбрать 1 или 2 столбца для сравнения (множественный выбор)
   • При выборе 2 столбцов их значения объединяются через ПРОБЕЛ в единую строку
   • Сравнение проходит по объединенной строке как единому целому
   • Если не выбрано ничего - используется первый столбец по умолчанию

🔹 НАСЛЕДОВАНИЕ СТОЛБЦОВ (чекбоксы):
   • "Наследовать остальные столбцы источника 1" - ВКЛ по умолчанию
   • "Наследовать остальные столбцы источника 2" - ВКЛ по умолчанию
   • Наследуемые столбцы получают префикс: "Источник 1: [имя]"
   • Отключите чекбокс, если не нужны дополнительные столбцы

📋 ПРИМЕР:
   Источник 1 имеет столбцы: [Название ПО, Версия, Vendor]
   Источник 2 имеет столбцы: [Product Name, Category]

   Вы выбираете для сравнения:
   • Источник 1: "Название ПО"
   • Источник 2: "Product Name"

   В результате будут столбцы:
   • Источник 1: Название ПО
   • Источник 1: Версия (если чекбокс ВКЛ)
   • Источник 1: Vendor (если чекбокс ВКЛ)
   • Источник 2: Product Name
   • Источник 2: Category (если чекбокс ВКЛ)
   • Процент совпадения
   • Метод
"""

        tk.Label(section3, text=help_text_columns, font=("Consolas", 9),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        # Раздел 4: Экспорт результатов
        section4 = tk.LabelFrame(scrollable_frame, text="💾 Экспорт результатов",
                                font=("Arial", 11, "bold"), padx=15, pady=10, bg="white")
        section4.pack(fill=tk.X, pady=(0, 15))

        help_text_export = """
📊 ПОЛНЫЙ ОТЧЁТ:
   • Все результаты сопоставления
   • Отдельный лист со статистикой
   • Цветовая раскраска по качеству совпадения

✅ ТОЧНЫЕ СОВПАДЕНИЯ (100%):
   • Только записи с точным совпадением
   • Готово к импорту без проверки

⚠️ ТРЕБУЮТ ПРОВЕРКИ (<90%):
   • Записи с совпадением ниже 90%
   • Рекомендуется ручная проверка

❌ БЕЗ СОВПАДЕНИЙ (0%):
   • Записи, для которых не найдено совпадений
   • Требуется ручной поиск или добавление
"""

        tk.Label(section4, text=help_text_export, font=("Consolas", 9),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        # Раздел 5: Библиотеки
        section5 = tk.LabelFrame(scrollable_frame, text="📚 Используемые библиотеки",
                                font=("Arial", 11, "bold"), padx=15, pady=10, bg="white")
        section5.pack(fill=tk.X, pady=(0, 15))

        help_text_libs = """
🏠 BUILTIN (встроенные методы):
   • Точное совпадение (ВПР) - аналог Excel VLOOKUP
   • Возвращает 100% при точном совпадении, 0% при несовпадении
   • Мгновенная работа, не требует внешних библиотек
   • Полезно для точного поиска без приблизительного сопоставления

🔬 RAPIDFUZZ (рекомендуется):
   • Самая быстрая библиотека (в 100 раз быстрее аналогов)
   • Методы: WRatio, Token Set, Token Sort, Partial Ratio и др.
   • Оптимизирована для больших датасетов

📊 TEXTDISTANCE:
   • Научные метрики расстояния
   • Методы: Jaro-Winkler, Jaro, Jaccard, Sorensen-Dice, Cosine
   • Медленнее, но иногда точнее

🔊 JELLYFISH:
   • Фонетическое сравнение (для похоже звучащих слов)
   • Методы: Jaro-Winkler, Jaro
   • Полезно для имён и названий с опечатками
"""

        tk.Label(section5, text=help_text_libs, font=("Consolas", 9),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        # Раздел 6: Подробное описание методов
        section6 = tk.LabelFrame(scrollable_frame, text="🔍 Подробное описание всех методов сопоставления",
                                font=("Arial", 11, "bold"), padx=15, pady=10, bg="white")
        section6.pack(fill=tk.X, pady=(0, 15))

        help_text_methods = """
═══════════════════════════════════════════════════════════════════
🏠 ВСТРОЕННЫЕ МЕТОДЫ (1 метод) - NEW в v2.2
═══════════════════════════════════════════════════════════════════

0️⃣ Точное совпадение (ВПР) 📌
   ОПИСАНИЕ: Аналог функции VLOOKUP в Excel - точное сопоставление
   КАК РАБОТАЕТ: Сравнивает нормализованные строки на полное совпадение
   ПРИМЕР:
      "Microsoft Office" vs "microsoft office" → 100% (игнор регистра)
      "Microsoft Office" vs "MS Office" → 0% (не совпадает)
      "Chrome" vs "Chrome Browser" → 0% (не совпадает)
   КОГДА ИСПОЛЬЗОВАТЬ: Когда нужно ТОЧНОЕ совпадение без приближения
   ОСОБЕННОСТИ: Мгновенная работа, возвращает только 100% или 0%

═══════════════════════════════════════════════════════════════════
📚 БИБЛИОТЕКА RAPIDFUZZ (10 методов)
═══════════════════════════════════════════════════════════════════

1️⃣ WRatio (рекомендуется) 🥇
   ОПИСАНИЕ: Взвешенный алгоритм, автоматически выбирает лучшую стратегию
   КАК РАБОТАЕТ: Комбинирует несколько методов и выбирает лучший результат
   ПРИМЕР:
      "Microsoft Office 365" vs "MS Office 365" → 95%
      "Adobe Reader DC" vs "Acrobat Reader" → 85%
   КОГДА ИСПОЛЬЗОВАТЬ: Универсальный метод, подходит для большинства случаев

2️⃣ Token Set
   ОПИСАНИЕ: Сравнивает уникальные слова, игнорируя порядок и дубликаты
   КАК РАБОТАЕТ: Разбивает на слова, убирает повторы, сравнивает множества
   ПРИМЕР:
      "SQL Server Management Studio" vs "Management Studio SQL Server" → 100%
      "Microsoft SQL Server" vs "SQL Server Microsoft" → 100%
   КОГДА ИСПОЛЬЗОВАТЬ: Когда порядок слов не важен

3️⃣ Token Sort
   ОПИСАНИЕ: Сортирует слова по алфавиту перед сравнением
   КАК РАБОТАЕТ: Разбивает на слова, сортирует, затем сравнивает
   ПРИМЕР:
      "Visual Studio Code" vs "Code Studio Visual" → 100%
      "Google Chrome Browser" vs "Browser Chrome Google" → 100%
   КОГДА ИСПОЛЬЗОВАТЬ: Когда названия содержат одни слова в разном порядке

4️⃣ Partial Ratio
   ОПИСАНИЕ: Ищет лучшее совпадение подстрок
   КАК РАБОТАЕТ: Проверяет, содержится ли короткая строка в длинной
   ПРИМЕР:
      "Office" vs "Microsoft Office 365 ProPlus" → 90%
      "Chrome" vs "Google Chrome Browser" → 85%
   КОГДА ИСПОЛЬЗОВАТЬ: Когда одно название короткое, другое полное

5️⃣ Ratio
   ОПИСАНИЕ: Простое посимвольное сравнение Левенштейна
   КАК РАБОТАЕТ: Считает минимальное число правок для совпадения
   ПРИМЕР:
      "Firefox" vs "Firefoх" (русская 'х') → 95%
      "Chrome" vs "Chromium" → 85%
   КОГДА ИСПОЛЬЗОВАТЬ: Для точного посимвольного сравнения

6️⃣ QRatio
   ОПИСАНИЕ: Быстрая версия Ratio с оптимизациями
   КАК РАБОТАЕТ: Упрощенный алгоритм для ускорения
   ПРИМЕР: Аналогично Ratio, но работает быстрее
   КОГДА ИСПОЛЬЗОВАТЬ: Когда нужна скорость, а не максимальная точность

7️⃣ Token Ratio
   ОПИСАНИЕ: Комбинация Ratio и Token Sort
   КАК РАБОТАЕТ: Сортирует токены, затем применяет Ratio
   ПРИМЕР:
      "Adobe Acrobat Reader" vs "Reader Acrobat Adobe" → 100%
   КОГДА ИСПОЛЬЗОВАТЬ: Средний вариант между Token Sort и Ratio

8️⃣ Partial Token Ratio
   ОПИСАНИЕ: Комбинация Partial Ratio и Token Ratio
   КАК РАБОТАЕТ: Частичное совпадение с сортировкой токенов
   ПРИМЕР:
      "MS Office" vs "Microsoft Office Professional Plus 2021" → 85%
   КОГДА ИСПОЛЬЗОВАТЬ: Короткое название vs полное с сортировкой

9️⃣ Partial Token Set
   ОПИСАНИЕ: Комбинация Partial Ratio и Token Set
   КАК РАБОТАЕТ: Частичное совпадение уникальных токенов
   ПРИМЕР:
      "SQL Server" vs "Microsoft SQL Server Enterprise Edition" → 90%
   КОГДА ИСПОЛЬЗОВАТЬ: Подмножество слов в большом названии

🔟 Partial Token Sort
   ОПИСАНИЕ: Комбинация Partial Ratio и Token Sort
   КАК РАБОТАЕТ: Частичное совпадение с сортированными токенами
   ПРИМЕР:
      "Chrome Browser" vs "Google Chrome Web Browser Latest" → 85%
   КОГДА ИСПОЛЬЗОВАТЬ: Частичное совпадение с игнорированием порядка

═══════════════════════════════════════════════════════════════════
📊 БИБЛИОТЕКА TEXTDISTANCE (5 методов)
═══════════════════════════════════════════════════════════════════

1️⃣1️⃣ Jaro-Winkler
   ОПИСАНИЕ: Фонетическое сравнение с акцентом на начало строки
   КАК РАБОТАЕТ: Больший вес символам в начале названия
   ПРИМЕР:
      "Microsoft" vs "Mikrosoft" → 95%
      "PostgreSQL" vs "Postgres" → 90%
   КОГДА ИСПОЛЬЗОВАТЬ: Когда начало названия важнее конца

1️⃣2️⃣ Jaro
   ОПИСАНИЕ: Базовая версия Jaro без акцента на начало
   КАК РАБОТАЕТ: Равномерное сравнение по всей строке
   ПРИМЕР:
      "Oracle" vs "Orakle" → 92%
   КОГДА ИСПОЛЬЗОВАТЬ: Для фонетического сравнения без приоритета

1️⃣3️⃣ Jaccard
   ОПИСАНИЕ: Сравнение множеств символов/биграмм
   КАК РАБОТАЕТ: Считает пересечение и объединение множеств
   ПРИМЕР:
      "MySQL" vs "MariaDB" → 40% (мало общих символов)
      "Python" vs "Pythan" → 85%
   КОГДА ИСПОЛЬЗОВАТЬ: Для оценки общих элементов

1️⃣4️⃣ Sorensen-Dice
   ОПИСАНИЕ: Похож на Jaccard, но с другой формулой
   КАК РАБОТАЕТ: Удвоенное пересечение / сумма размеров
   ПРИМЕР:
      "Redis" vs "Reddis" → 80%
   КОГДА ИСПОЛЬЗОВАТЬ: Альтернатива Jaccard, чуть мягче

1️⃣5️⃣ Cosine
   ОПИСАНИЕ: Косинусное расстояние векторов символов
   КАК РАБОТАЕТ: Преобразует строки в векторы, считает угол
   ПРИМЕР:
      "MongoDB" vs "Mongo DB" → 95%
   КОГДА ИСПОЛЬЗОВАТЬ: Для математически точного сравнения

═══════════════════════════════════════════════════════════════════
🔊 БИБЛИОТЕКА JELLYFISH (2 метода)
═══════════════════════════════════════════════════════════════════

1️⃣6️⃣ Jaro-Winkler (Jellyfish)
   ОПИСАНИЕ: Аналог TextDistance Jaro-Winkler
   КАК РАБОТАЕТ: Та же логика, другая реализация
   ПРИМЕР: Аналогичные результаты TextDistance версии
   КОГДА ИСПОЛЬЗОВАТЬ: Альтернативная реализация

1️⃣7️⃣ Jaro (Jellyfish)
   ОПИСАНИЕ: Аналог TextDistance Jaro
   КАК РАБОТАЕТ: Та же логика, другая реализация
   ПРИМЕР: Аналогичные результаты TextDistance версии
   КОГДА ИСПОЛЬЗОВАТЬ: Альтернативная реализация

═══════════════════════════════════════════════════════════════════
💡 СОВЕТЫ ПО ВЫБОРУ МЕТОДА
═══════════════════════════════════════════════════════════════════

📌 Точное совпадение → Точное совпадение (ВПР)
🎯 Общий случай → WRatio (автоматический выбор стратегии)
📝 Разный порядок слов → Token Set / Token Sort
✂️ Короткое vs полное → Partial Ratio / Partial Token Set
🔤 Опечатки → Jaro-Winkler / Ratio
⚡ Нужна скорость → QRatio
🎲 Не уверены → Запустите режим "Сравнение методов"!
"""

        tk.Label(section6, text=help_text_methods, font=("Consolas", 8),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def create_setup_tab(self):
        """Вкладка загрузки файлов с прокруткой"""
        # Создаем Canvas для прокрутки
        canvas = tk.Canvas(self.setup_tab)
        scrollbar = tk.Scrollbar(self.setup_tab, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, padx=20, pady=20)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Bind mousewheel для прокрутки
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        main_frame = scrollable_frame
        
        # Описание функциональности приложения
        info_frame = tk.LabelFrame(main_frame, text="🎯 Возможности приложения",
                                   font=("Arial", 11, "bold"), padx=15, pady=15, bg="#F0F9FF")
        info_frame.pack(fill=tk.X, pady=(0, 20))

        description_text = """
🔍 Интеллектуальное сопоставление данных из двух источников

✨ ОСНОВНЫЕ ВОЗМОЖНОСТИ:

  • Поддержка Excel (.xlsx, .xls) и CSV файлов
  • Сравнение по 1 или сразу 2 столбцам одновременно
  • 18 методов сопоставления: от простого ВПР до продвинутых алгоритмов
  • Гибкий выбор методов: работа только с выбранными или со всеми сразу
  • Выбор любых столбцов из каждого источника для сравнения

🎯 РЕЖИМЫ РАБОТЫ:

  1. Автоматический - находит лучший метод из выбранных
  2. Сравнение - показывает статистику по выбранным методам
  3. Полное сравнение - применяет выбранные методы ко всем данным

  📚 Подробное описание всех режимов и методов — в разделе "Справка"

⚡ РЕЗУЛЬТАТ: Точное сопоставление данных даже при неполном совпадении названий!
"""

        tk.Label(info_frame, text=description_text,
                font=("Consolas", 9), justify=tk.LEFT, anchor="w",
                bg="#F0F9FF", fg="#1E40AF").pack(fill=tk.X)
        
        files_frame = tk.LabelFrame(main_frame, text="📁 Входные файлы (Excel / CSV)",
                                    font=("Arial", 11, "bold"), padx=10, pady=10)
        files_frame.pack(fill=tk.X, pady=(0, 20))
        
        askupo_frame = tk.Frame(files_frame)
        askupo_frame.pack(fill=tk.X, pady=5)
        tk.Label(askupo_frame, text="1️⃣ Источник данных 1 (целевой):",
                font=("Arial", 10, "bold")).pack(anchor=tk.W)
        self.askupo_label = tk.Label(askupo_frame, text="📂 Файл не выбран", fg="gray", font=("Arial", 9))
        self.askupo_label.pack(anchor=tk.W, padx=20)
        tk.Button(askupo_frame, text="📁 Выбрать Источник данных 1 (целевой)", command=self.select_askupo,
                 bg="#10B981", fg="white", font=("Arial", 9, "bold"),
                 padx=15, pady=5).pack(anchor=tk.W, padx=20, pady=3)

        eatool_frame = tk.Frame(files_frame)
        eatool_frame.pack(fill=tk.X, pady=5)
        tk.Label(eatool_frame, text="2️⃣ Источник данных 2:",
                font=("Arial", 10, "bold")).pack(anchor=tk.W)
        self.eatool_label = tk.Label(eatool_frame, text="📂 Файл не выбран", fg="gray", font=("Arial", 9))
        self.eatool_label.pack(anchor=tk.W, padx=20)
        tk.Button(eatool_frame, text="📁 Выбрать Источник данных 2", command=self.select_eatool,
                 bg="#10B981", fg="white", font=("Arial", 9, "bold"),
                 padx=15, pady=5).pack(anchor=tk.W, padx=20, pady=3)
        
        settings_frame = tk.LabelFrame(main_frame, text="Настройки обработки", 
                                       font=("Arial", 11, "bold"), padx=10, pady=10)
        settings_frame.pack(fill=tk.X, pady=(0, 20))
        
        mode_frame = tk.Frame(settings_frame)
        mode_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(mode_frame, text="Режим работы:", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        
        self.mode_var = tk.StringVar(value="auto")
        
        tk.Radiobutton(mode_frame,
                      text="🤖 Автоматический - тестирует ВЫБРАННЫЕ методы и выбрав лучший создает результирующий эксель",
                      variable=self.mode_var, value="auto",
                      font=("Arial", 9)).pack(anchor=tk.W, padx=20)
        tk.Radiobutton(mode_frame,
                      text="📊 Сравнение выбранных методов - тестирует на выборке (не более первых 200 записей) и выбирает лучший",
                      variable=self.mode_var, value="compare",
                      font=("Arial", 9)).pack(anchor=tk.W, padx=20)
        tk.Radiobutton(mode_frame,
                      text="🔬 Полное сравнение - выбор одного или нескольких методов для создания результирующего эксель",
                      variable=self.mode_var, value="full_compare",
                      font=("Arial", 9)).pack(anchor=tk.W, padx=20)

        # Фрейм для выбора методов
        self.method_selector_frame = tk.Frame(settings_frame)
        self.method_selector_frame.pack(fill=tk.X, pady=5)

        tk.Label(self.method_selector_frame, text="Выберите методы (для всех режимов):",
                font=("Arial", 9, "bold")).pack(anchor=tk.W, padx=20)

        tk.Label(self.method_selector_frame,
                text="💡 Удерживайте Ctrl для выбора нескольких методов",
                font=("Arial", 8), fg="gray").pack(anchor=tk.W, padx=20)

        # Listbox с прокруткой для выбора методов
        listbox_frame = tk.Frame(self.method_selector_frame)
        listbox_frame.pack(anchor=tk.W, padx=20, pady=5)

        methods_scrollbar = tk.Scrollbar(listbox_frame)
        methods_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.methods_listbox = tk.Listbox(listbox_frame,
                                          selectmode=tk.MULTIPLE,
                                          height=8,
                                          width=80,
                                          yscrollcommand=methods_scrollbar.set,
                                          exportselection=False)
        self.methods_listbox.pack(side=tk.LEFT, fill=tk.BOTH)
        methods_scrollbar.config(command=self.methods_listbox.yview)

        # Заполняем список методами
        for method in self.methods:
            self.methods_listbox.insert(tk.END, method.name)

        # Выбираем первый метод по умолчанию
        if self.methods:
            self.methods_listbox.selection_set(0)

        # Кнопка "Выбрать все методы"
        button_frame = tk.Frame(self.method_selector_frame)
        button_frame.pack(anchor=tk.W, padx=20, pady=5)

        tk.Button(button_frame, text="✓ Выбрать все",
                 command=self.select_all_methods,
                 font=("Arial", 8), padx=10, pady=3).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="✗ Снять выбор",
                 command=self.deselect_all_methods,
                 font=("Arial", 8), padx=10, pady=3).pack(side=tk.LEFT, padx=5)

        # ==== НОВАЯ СЕКЦИЯ: Выбор столбцов для сравнения ====
        columns_frame = tk.LabelFrame(main_frame, text="Выбор столбцов для сравнения",
                                      font=("Arial", 11, "bold"), padx=10, pady=10)
        columns_frame.pack(fill=tk.X, pady=(0, 10))

        # Контейнер для двух источников
        columns_container = tk.Frame(columns_frame)
        columns_container.pack(fill=tk.X)

        # Источник 1 (левая колонка)
        source1_frame = tk.Frame(columns_container)
        source1_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        tk.Label(source1_frame, text="📂 Источник данных 1 (целевой):",
                font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))

        tk.Label(source1_frame, text="Выберите столбцы для сравнения (1-2 столбца):",
                font=("Arial", 9)).pack(anchor=tk.W)

        listbox_frame1 = tk.Frame(source1_frame)
        listbox_frame1.pack(fill=tk.BOTH, expand=True)

        scrollbar1 = tk.Scrollbar(listbox_frame1)
        scrollbar1.pack(side=tk.RIGHT, fill=tk.Y)

        self.askupo_col_listbox = tk.Listbox(listbox_frame1, selectmode=tk.MULTIPLE,
                                             height=5, yscrollcommand=scrollbar1.set,
                                             exportselection=False)
        self.askupo_col_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar1.config(command=self.askupo_col_listbox.yview)
        self.askupo_col_listbox.bind('<<ListboxSelect>>', self.on_askupo_column_select)

        tk.Checkbutton(source1_frame, text="Наследовать остальные столбцы источника 1",
                      variable=self.inherit_askupo_cols_var,
                      font=("Arial", 9)).pack(anchor=tk.W, pady=(5, 0))

        # Источник 2 (правая колонка)
        source2_frame = tk.Frame(columns_container)
        source2_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        tk.Label(source2_frame, text="📂 Источник данных 2:",
                font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))

        tk.Label(source2_frame, text="Выберите столбцы для сравнения (1-2 столбца):",
                font=("Arial", 9)).pack(anchor=tk.W)

        listbox_frame2 = tk.Frame(source2_frame)
        listbox_frame2.pack(fill=tk.BOTH, expand=True)

        scrollbar2 = tk.Scrollbar(listbox_frame2)
        scrollbar2.pack(side=tk.RIGHT, fill=tk.Y)

        self.eatool_col_listbox = tk.Listbox(listbox_frame2, selectmode=tk.MULTIPLE,
                                             height=5, yscrollcommand=scrollbar2.set,
                                             exportselection=False)
        self.eatool_col_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar2.config(command=self.eatool_col_listbox.yview)
        self.eatool_col_listbox.bind('<<ListboxSelect>>', self.on_eatool_column_select)

        tk.Checkbutton(source2_frame, text="Наследовать остальные столбцы источника 2",
                      variable=self.inherit_eatool_cols_var,
                      font=("Arial", 9)).pack(anchor=tk.W, pady=(5, 0))

        # Чекбокс для режима множественных столбцов
        tk.Checkbutton(columns_frame,
                      text="🔗 Режим сравнения по 2 столбцам одновременно (требует выбора 2 столбцов в каждом источнике)",
                      variable=self.multi_column_mode_var,
                      font=("Arial", 9, "bold"), fg="#7C3AED").pack(anchor=tk.W, pady=(10, 0))

        tk.Label(columns_frame,
                text="💡 Подсказка: После выбора файлов, столбцы появятся в списках. Выберите 1-2 столбца для сравнения.",
                font=("Arial", 8), fg="gray", wraplength=700, justify=tk.LEFT).pack(anchor=tk.W, pady=(5, 0))

        self.process_btn = tk.Button(main_frame, text="🚀 Начать обработку",
                 command=self.start_processing, bg="#7C3AED", fg="white",
                 font=("Arial", 13, "bold"), padx=50, pady=12,
                 state=tk.DISABLED)
        self.process_btn.pack(pady=20)
        
    def create_comparison_tab(self):
        """Вкладка сравнения методов"""
        frame = tk.Frame(self.comparison_tab, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(frame, text="📊 Сравнение производительности методов",
                font=("Arial", 13, "bold")).pack(pady=10)
        
        tree_frame = tk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        scroll_y = ttk.Scrollbar(tree_frame)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.comparison_tree = ttk.Treeview(
            tree_frame,
            columns=("rank", "method", "library", "perfect", "high", "avg_score", "time"),
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set,
            height=15
        )
        
        scroll_y.config(command=self.comparison_tree.yview)
        scroll_x.config(command=self.comparison_tree.xview)
        
        headers = [
            ("rank", "🏆", 50),
            ("method", "Метод", 300),
            ("library", "Библиотека", 120),
            ("perfect", "100%", 80),
            ("high", "90-99%", 80),
            ("avg_score", "Средний %", 100),
            ("time", "Время", 90),
        ]
        
        for col, text, width in headers:
            self.comparison_tree.heading(col, text=text)
            self.comparison_tree.column(col, width=width, 
                                       anchor=tk.CENTER if col != "method" else tk.W)
        
        self.comparison_tree.pack(fill=tk.BOTH, expand=True)
        
        btn_frame = tk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        tk.Button(btn_frame, text="💾 Экспортировать сравнение в Excel",
                 command=self.export_comparison, bg="#3B82F6", fg="white",
                 font=("Arial", 10, "bold"), padx=20, pady=5).pack(side=tk.LEFT, padx=5)
        
    def create_results_tab(self):
        """Вкладка результатов"""
        frame = tk.Frame(self.results_tab, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        self.result_info_frame = tk.Frame(frame)
        self.result_info_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.result_stats_frame = tk.Frame(frame)
        self.result_stats_frame.pack(fill=tk.X, pady=(0, 10))
        
        export_frame = tk.Frame(frame)
        export_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(export_frame, text="💾 Экспорт результатов в Excel:", 
                font=("Arial", 11, "bold")).pack(anchor=tk.W)
        
        btn_container = tk.Frame(export_frame)
        btn_container.pack(fill=tk.X, pady=5)
        
        export_buttons = [
            ("📊 Полный отчет", self.export_full, "#4F46E5"),
            ("✅ Точные (100%)", self.export_perfect, "#10B981"),
            ("⚠️ Требуют проверки (<90%)", self.export_problems, "#F59E0B"),
            ("❌ Без совпадений (0%)", self.export_no_match, "#EF4444"),
        ]
        
        for text, command, color in export_buttons:
            tk.Button(btn_container, text=text, command=command, bg=color, fg="white",
                     font=("Arial", 10, "bold"), padx=18, pady=6).pack(side=tk.LEFT, padx=3)
        
        tk.Label(frame, text="📋 Результаты сопоставления (первые 50 записей):",
                font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        
        tree_frame = tk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        scroll_y = ttk.Scrollbar(tree_frame)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.results_tree = ttk.Treeview(
            tree_frame,
            columns=("num", "askupo", "eatool", "percent"),
            show="headings",
            yscrollcommand=scroll_y.set,
            height=15
        )
        scroll_y.config(command=self.results_tree.yview)
        
        headers = [
            ("num", "№", 50),
            ("askupo", "Источник 1 (сравниваемый столбец)", 350),
            ("eatool", "Источник 2 (сопоставленный столбец)", 350),
            ("percent", "Процент совпадения", 120),
        ]
        
        for col, text, width in headers:
            self.results_tree.heading(col, text=text)
            self.results_tree.column(col, width=width, 
                                    anchor=tk.CENTER if col in ["num", "percent"] else tk.W)
        
        self.results_tree.pack(fill=tk.BOTH, expand=True)
        
    def read_data_file(self, filename: str, nrows=None) -> pd.DataFrame:
        """Универсальное чтение Excel или CSV файла

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
                    return df
                except (UnicodeDecodeError, Exception):
                    continue
            # Если ничего не сработало, пробуем без указания кодировки
            df = pd.read_csv(filename, nrows=nrows)
        else:
            # Excel файлы (.xlsx, .xls)
            df = pd.read_excel(filename, nrows=nrows)

        return df

    def validate_excel_file(self, filename: str) -> Tuple[bool, str]:
        """Валидация Excel или CSV файла - упрощенная проверка (v2.1)"""
        try:
            df = self.read_data_file(filename)

            if df.empty:
                return False, "Файл пустой (нет данных)"

            if len(df.columns) == 0:
                return False, "Файл не содержит столбцов"

            if len(df) == 0:
                return False, "Файл не содержит строк с данными"

            # Успешная валидация - показываем информацию о файле
            return True, f"✅ Файл валидный\n   Записей: {len(df)}\n   Столбцов: {len(df.columns)}\n   Список столбцов: {', '.join(df.columns[:5])}{' ...' if len(df.columns) > 5 else ''}"

        except Exception as e:
            return False, f"Ошибка чтения файла:\n{str(e)}"

    def select_askupo(self):
        filename = filedialog.askopenfilename(
            title="Выберите Источник данных 1 (целевой)",
            filetypes=[("Data files", "*.xlsx *.xls *.csv"), ("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            # Валидация файла
            is_valid, message = self.validate_excel_file(filename)

            if not is_valid:
                messagebox.showerror("❌ Ошибка валидации Источника данных 1",
                                   f"Файл не прошёл проверку:\n\n{message}\n\n"
                                   f"Требования:\n"
                                   f"• Файл должен содержать данные (не пустой)\n"
                                   f"• Файл должен иметь столбцы и строки\n"
                                   f"• Формат: .xlsx, .xls или .csv")
                return

            self.askupo_file = filename
            display_name = Path(filename).name
            if len(display_name) > 50:
                display_name = display_name[:47] + "..."
            self.askupo_label.config(text=f"✅ {display_name}", fg="green", font=("Arial", 9, "bold"))

            # Загрузка столбцов из файла
            self.load_askupo_columns()
            self.check_ready()
    
    def select_eatool(self):
        filename = filedialog.askopenfilename(
            title="Выберите Источник данных 2",
            filetypes=[("Data files", "*.xlsx *.xls *.csv"), ("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            # Валидация файла
            is_valid, message = self.validate_excel_file(filename)

            if not is_valid:
                messagebox.showerror("❌ Ошибка валидации Источника данных 2",
                                   f"Файл не прошёл проверку:\n\n{message}\n\n"
                                   f"Требования:\n"
                                   f"• Файл должен содержать данные (не пустой)\n"
                                   f"• Файл должен иметь столбцы и строки\n"
                                   f"• Формат: .xlsx, .xls или .csv")
                return

            self.eatool_file = filename
            display_name = Path(filename).name
            if len(display_name) > 50:
                display_name = display_name[:47] + "..."
            self.eatool_label.config(text=f"✅ {display_name}", fg="green", font=("Arial", 9, "bold"))

            # Загрузка столбцов из файла
            self.load_eatool_columns()
            self.check_ready()
    
    def check_ready(self):
        if self.askupo_file and self.eatool_file:
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

            askupo_col = askupo_df.columns[0]
            eatool_col = eatool_df.columns[0]

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
                                                  askupo_col, eatool_col)

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
                                       askupo_col, eatool_col)

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

            askupo_col = askupo_df.columns[0]
            eatool_col = eatool_df.columns[0]

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
                results = self.test_method_optimized(method, sample_askupo, eatool_df,
                                                     askupo_col, eatool_col)
                elapsed = time.time() - start_time

                # Используем ИСПРАВЛЕННУЮ функцию подсчета статистики
                stats_dict = self.calculate_statistics(results)

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

        askupo_col = askupo_df.columns[0]
        eatool_col = eatool_df.columns[0]

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
            results_df = self.test_method_optimized(method, askupo_df, eatool_df,
                                                   askupo_col, eatool_col)

            # Сохраняем результаты
            all_methods_results[method.name] = results_df

            # Подсчитываем статистику
            stats_dict = self.calculate_statistics(results_df)

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

    # Методы run_manual_mode и run_multi_manual_mode УДАЛЕНЫ в v2.2
    # Вся функциональность теперь в run_full_comparison_mode

    def evaluate_method_fast(self, method: MatchingMethod, sample_askupo: pd.DataFrame,
                            eatool_df: pd.DataFrame, askupo_col: str, eatool_col: str) -> tuple:
        """Быстрая оценка качества метода

        Возвращает кортеж для лексикографического сравнения:
        (количество 100%, количество 90-99%, средний процент)
        Это обеспечивает единообразие с режимом сравнения методов.
        """
        results = self.test_method_optimized(method, sample_askupo, eatool_df,
                                            askupo_col, eatool_col)

        stats = self.calculate_statistics(results)

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
            combined = self.combine_columns(row, eatool_cols)
            eatool_combined_names.append(combined)
            eatool_original_values.append(combined)

        # Нормализация для поиска
        eatool_normalized = [self.normalize_string(name) for name in eatool_combined_names]
        choice_dict = {norm: orig for norm, orig in zip(eatool_normalized, eatool_original_values)}

        # Создаём словарь для быстрого поиска строки по комбинированному значению
        eatool_row_dict = {}
        for idx, row in eatool_df.iterrows():
            combined = self.combine_columns(row, eatool_cols)
            eatool_row_dict[combined] = row

        results = []

        for _, row in askupo_df.iterrows():
            # Объединяем значения из выбранных столбцов источника 1
            askupo_combined = self.combine_columns(row, askupo_cols)
            askupo_normalized = self.normalize_string(askupo_combined)

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
                               eatool_df: pd.DataFrame, askupo_col: str, eatool_col: str):
        """Оптимизированное применение метода"""
        
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
        
        eatool_names = eatool_df[eatool_col].tolist()
        eatool_normalized = [self.normalize_string(name) for name in eatool_names]
        choice_dict = {norm: orig for norm, orig in zip(eatool_normalized, eatool_names)}

        # Создаём словарь для быстрого поиска строки по оригинальному имени
        eatool_row_dict = {str(row[eatool_col]): row for _, row in eatool_df.iterrows()}

        status_label.config(text="Обработка записей...")

        results = []
        total = len(askupo_df)
        progress_bar['maximum'] = total

        for idx, row in askupo_df.iterrows():
            askupo_name = str(row[askupo_col])
            askupo_normalized = self.normalize_string(askupo_name)

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
                askupo_combined=askupo_name,
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
        stats = self.calculate_statistics(self.results)
        
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
        stats = self.calculate_statistics(self.results)
        
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
        if not self.methods_comparison:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Сравнение_методов_сопоставления.xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not save_path:
            return
        
        df = pd.DataFrame([
            {
                'Место': i + 1,
                'Метод': stats['method'],
                'Библиотека': stats['library'],
                '100% (точное)': stats['perfect'],
                '90-99% (высокое)': stats['high'],
                '70-89% (среднее)': stats['medium'],
                'Средний %': round(stats['avg_score'], 1),
                'Время (сек)': round(stats['time'], 2)
            }
            for i, stats in enumerate(self.methods_comparison)
        ])
        
        try:
            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Сравнение методов', index=False)
                
                workbook = writer.book
                worksheet = writer.sheets['Сравнение методов']
                
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#7C3AED',
                    'font_color': 'white',
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1
                })
                
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                worksheet.set_column('A:A', 10)
                worksheet.set_column('B:B', 40)
                worksheet.set_column('C:H', 18)
            
            messagebox.showinfo("Успех", f"✅ Файл сохранен:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка экспорта:\n{str(e)}")
    
    def export_full(self):
        if self.results is None:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return
        self.export_excel(self.results, "Полный_отчет_сопоставления.xlsx", include_stats=True)
    
    def export_perfect(self):
        if self.results is None:
            return
        data = self.results[self.results['Процент совпадения'] == 100]
        self.export_excel(data, "Точные_совпадения_100%.xlsx")
    
    def export_problems(self):
        if self.results is None:
            return
        data = self.results[self.results['Процент совпадения'] < 90]
        self.export_excel(data, "Требуют_проверки_менее_90%.xlsx")
    
    def export_no_match(self):
        if self.results is None:
            return
        data = self.results[self.results['Процент совпадения'] == 0]
        self.export_excel(data, "Без_совпадений_0%.xlsx")
    
    def export_excel(self, data: pd.DataFrame, filename: str, include_stats: bool = False):
        """Экспорт в Excel"""
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not save_path:
            return
        
        try:
            data_to_export = data.copy()

            # Заменяем NaN и inf на пустые строки для корректного экспорта
            data_to_export = data_to_export.replace([np.nan, np.inf, -np.inf], "")

            data_to_export.insert(0, '№', range(1, len(data_to_export) + 1))

            with pd.ExcelWriter(save_path, engine='xlsxwriter',
                              engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
                data_to_export.to_excel(writer, sheet_name='Результаты', index=False)
                
                workbook = writer.book
                worksheet = writer.sheets['Результаты']
                
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#7C3AED',
                    'font_color': 'white',
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1
                })
                
                for col_num, value in enumerate(data_to_export.columns.values):
                    worksheet.write(0, col_num, value, header_format)

                # Динамическая ширина столбцов
                worksheet.set_column('A:A', 8)  # Номер строки
                # Столбцы B и далее - автоматическая ширина в зависимости от количества
                for col_num in range(1, len(data_to_export.columns)):
                    col_name = data_to_export.columns[col_num]
                    if 'Источник данных' in str(col_name):
                        worksheet.set_column(col_num, col_num, 45)  # Широкие столбцы для названий
                    elif col_name == 'Процент совпадения':
                        worksheet.set_column(col_num, col_num, 12)  # Узкий для процента
                    elif col_name == 'Метод':
                        worksheet.set_column(col_num, col_num, 35)  # Средний для метода
                    else:
                        worksheet.set_column(col_num, col_num, 20)  # Остальные столбцы из Источника 2
                
                formats = {
                    100: workbook.add_format({'bg_color': '#D1FAE5', 'border': 1}),
                    90: workbook.add_format({'bg_color': '#DBEAFE', 'border': 1}),
                    70: workbook.add_format({'bg_color': '#FEF3C7', 'border': 1}),
                    50: workbook.add_format({'bg_color': '#FED7AA', 'border': 1}),
                    1: workbook.add_format({'bg_color': '#FFE4E1', 'border': 1}),
                    0: workbook.add_format({'bg_color': '#FEE2E2', 'border': 1})
                }
                
                for row_num in range(1, len(data_to_export) + 1):
                    percent = data_to_export.iloc[row_num - 1]['Процент совпадения']
                    
                    if percent == 100:
                        fmt = formats[100]
                    elif percent >= 90:
                        fmt = formats[90]
                    elif percent >= 70:
                        fmt = formats[70]
                    elif percent >= 50:
                        fmt = formats[50]
                    elif percent > 0:
                        fmt = formats[1]
                    else:
                        fmt = formats[0]
                    
                    for col_num in range(len(data_to_export.columns)):
                        worksheet.write(row_num, col_num, 
                                      data_to_export.iloc[row_num - 1, col_num], fmt)
                
                if include_stats and self.results is not None:
                    # Используем ИСПРАВЛЕННУЮ функцию статистики
                    stats = self.calculate_statistics(self.results)
                    
                    stats_data = pd.DataFrame([
                        {'Категория': 'Всего записей', 'Количество': stats['total'], 'Процент': '100%'},
                        {'Категория': '100% (точное совпадение)', 'Количество': stats['perfect'], 'Процент': f"{stats['perfect']/stats['total']*100:.1f}%"},
                        {'Категория': '90-99% (высокое совпадение)', 'Количество': stats['high'], 'Процент': f"{stats['high']/stats['total']*100:.1f}%"},
                        {'Категория': '70-89% (среднее совпадение)', 'Количество': stats['medium'], 'Процент': f"{stats['medium']/stats['total']*100:.1f}%"},
                        {'Категория': '50-69% (низкое совпадение)', 'Количество': stats['low'], 'Процент': f"{stats['low']/stats['total']*100:.1f}%"},
                        {'Категория': '1-49% (очень низкое совпадение)', 'Количество': stats['very_low'], 'Процент': f"{stats['very_low']/stats['total']*100:.1f}%"},
                        {'Категория': '0% (нет совпадения)', 'Количество': stats['none'], 'Процент': f"{stats['none']/stats['total']*100:.1f}%"},
                        {'Категория': '---', 'Количество': '---', 'Процент': '---'},
                        {'Категория': 'Проверка суммы', 'Количество': stats['check_sum'], 'Процент': '✅' if stats['check_sum'] == stats['total'] else '❌ ОШИБКА!'}
                    ])
                    stats_data.to_excel(writer, sheet_name='Статистика', index=False)
            
            messagebox.showinfo("Успех", f"✅ Файл сохранен:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"❌ Ошибка при экспорте:\n{str(e)}")

    def export_full_comparison_to_excel(self, default_filename=None):
        """Экспорт полного сравнения всех методов в Excel

        Args:
            default_filename: Имя файла по умолчанию (опционально)
        """
        if not self.full_comparison_results:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return

        if default_filename is None:
            default_filename = "Полное_сравнение_всех_методов.xlsx"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not save_path:
            return

        try:
            methods_data = self.full_comparison_results['methods_data']
            comparison_stats = self.full_comparison_results['comparison_stats']

            # Очищаем все DataFrame от NaN и inf
            cleaned_methods_data = {}
            for method_name, df in methods_data.items():
                cleaned_df = df.copy()
                cleaned_df = cleaned_df.replace([np.nan, np.inf, -np.inf], "")
                cleaned_methods_data[method_name] = cleaned_df

            with pd.ExcelWriter(save_path, engine='xlsxwriter',
                              engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
                workbook = writer.book

                # Форматы
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#7C3AED',
                    'font_color': 'white',
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1
                })

                formats = {
                    100: workbook.add_format({'bg_color': '#D1FAE5', 'border': 1}),
                    90: workbook.add_format({'bg_color': '#DBEAFE', 'border': 1}),
                    70: workbook.add_format({'bg_color': '#FEF3C7', 'border': 1}),
                    50: workbook.add_format({'bg_color': '#FED7AA', 'border': 1}),
                    1: workbook.add_format({'bg_color': '#FFE4E1', 'border': 1}),
                    0: workbook.add_format({'bg_color': '#FEE2E2', 'border': 1})
                }

                # 1. Лист "Сводка" - сравнительная таблица всех методов
                summary_df = pd.DataFrame([
                    {
                        '🏆 Место': i + 1,
                        'Метод': stats['method'],
                        'Библиотека': stats['library'],
                        'Всего записей': stats['total'],
                        '100% (точное)': stats['perfect'],
                        '90-99% (высокое)': stats['high'],
                        '70-89% (среднее)': stats['medium'],
                        '50-69% (низкое)': stats['low'],
                        '1-49% (очень низкое)': stats['very_low'],
                        '0% (нет)': stats['none'],
                        'Средний %': round(stats['avg_score'], 1),
                        'Время (сек)': round(stats['time'], 2)
                    }
                    for i, stats in enumerate(comparison_stats)
                ])

                summary_df.to_excel(writer, sheet_name='📊 Сводка', index=False)
                worksheet = writer.sheets['📊 Сводка']

                for col_num, value in enumerate(summary_df.columns.values):
                    worksheet.write(0, col_num, value, header_format)

                worksheet.set_column('A:A', 10)
                worksheet.set_column('B:B', 40)
                worksheet.set_column('C:L', 15)

                # 2. Листы для каждого метода
                for method_name, results_df in cleaned_methods_data.items():
                    # Удаляем эмодзи и недопустимые символы сначала
                    sheet_name = method_name

                    # Удаляем эмодзи (могут вызывать проблемы в Excel)
                    sheet_name = ''.join(char for char in sheet_name if ord(char) < 128)

                    # Удаляем недопустимые символы для названия листа Excel
                    invalid_chars = [':', '\\', '/', '?', '*', '[', ']']
                    for char in invalid_chars:
                        sheet_name = sheet_name.replace(char, '_')

                    # Убираем лишние пробелы
                    sheet_name = sheet_name.strip()

                    # Ограничиваем длину названия листа (Excel лимит 31 символ)
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:28] + "..."

                    # Если название пустое после очистки, используем номер
                    if not sheet_name:
                        sheet_name = f"Method_{list(cleaned_methods_data.keys()).index(method_name) + 1}"

                    # Добавляем номер строки
                    export_df = results_df.copy()
                    export_df.insert(0, '№', range(1, len(export_df) + 1))

                    export_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    worksheet = writer.sheets[sheet_name]

                    # Заголовки
                    for col_num, value in enumerate(export_df.columns.values):
                        worksheet.write(0, col_num, value, header_format)

                    # Динамическая ширина столбцов
                    worksheet.set_column('A:A', 8)  # Номер строки
                    for col_num in range(1, len(export_df.columns)):
                        col_name = export_df.columns[col_num]
                        if 'Источник данных' in str(col_name):
                            worksheet.set_column(col_num, col_num, 45)  # Широкие столбцы для названий
                        elif col_name == 'Процент совпадения':
                            worksheet.set_column(col_num, col_num, 12)  # Узкий для процента
                        elif col_name == 'Метод':
                            worksheet.set_column(col_num, col_num, 35)  # Средний для метода
                        else:
                            worksheet.set_column(col_num, col_num, 20)  # Остальные столбцы из Источника 2

                    # Цветовая раскраска по проценту совпадения
                    for row_num in range(1, len(export_df) + 1):
                        percent = export_df.iloc[row_num - 1]['Процент совпадения']

                        if percent == 100:
                            fmt = formats[100]
                        elif percent >= 90:
                            fmt = formats[90]
                        elif percent >= 70:
                            fmt = formats[70]
                        elif percent >= 50:
                            fmt = formats[50]
                        elif percent > 0:
                            fmt = formats[1]
                        else:
                            fmt = formats[0]

                        for col_num in range(len(export_df.columns)):
                            worksheet.write(row_num, col_num,
                                          export_df.iloc[row_num - 1, col_num], fmt)

            messagebox.showinfo("Успех", f"✅ Полное сравнение сохранено!\n\n"
                              f"📁 Файл: {Path(save_path).name}\n"
                              f"📊 Листов: {len(cleaned_methods_data) + 1}\n"
                              f"   • Сводка: 1 лист\n"
                              f"   • Результаты методов: {len(cleaned_methods_data)} листов")

        except Exception as e:
            messagebox.showerror("Ошибка", f"❌ Ошибка при экспорте:\n{str(e)}")


def main():
    root = tk.Tk()
    app = ExpertMatcher(root)
    root.mainloop()


if __name__ == "__main__":
    main()