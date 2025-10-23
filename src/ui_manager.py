"""
UIManager - управление пользовательским интерфейсом

Модуль отвечает за создание всех вкладок и UI компонентов приложения.
Выделен из expert_matcher.py для улучшения модульности.

Создан: 2025-10-22 (Этап 4 рефакторинга)
"""

import tkinter as tk
from tkinter import messagebox, ttk
from typing import TYPE_CHECKING

from .help_content import HelpContent
from .ui_components import TreeviewWithScrollbar, create_title_header

if TYPE_CHECKING:
    from expert_matcher import ExpertMatcher


class UIManager:
    """Менеджер пользовательского интерфейса"""

    def __init__(self, parent: 'ExpertMatcher'):
        """
        Инициализация менеджера UI

        Args:
            parent: Ссылка на экземпляр ExpertMatcher
        """
        self.parent = parent

    def create_widgets(self):
        """Создание интерфейса"""
        from src.constants import AppConstants

        # Заголовок приложения
        title_frame = create_title_header(
            self.parent.root,
            title=f"🔬 Expert Excel Matcher v{AppConstants.VERSION}",
            subtitle=f"⚡ В 100 раз быстрее! • {len(self.parent.methods)} методов • Расширенная нормализация"
        )
        title_frame.pack(fill=tk.X)

        self.parent.notebook = ttk.Notebook(self.parent.root)
        self.parent.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.parent.setup_tab = tk.Frame(self.parent.notebook)
        self.parent.notebook.add(self.parent.setup_tab, text="1. Загрузка и настройка")
        self.create_setup_tab()

        self.parent.comparison_tab = tk.Frame(self.parent.notebook)
        self.parent.notebook.add(self.parent.comparison_tab, text="2. Сравнение методов")
        self.create_comparison_tab()

        self.parent.results_tab = tk.Frame(self.parent.notebook)
        self.parent.notebook.add(self.parent.results_tab, text="3. Результаты")
        self.create_results_tab()

        self.parent.help_tab = tk.Frame(self.parent.notebook)
        self.parent.notebook.add(self.parent.help_tab, text="❓ Справка")
        self.create_help_tab()

    def create_help_tab(self):
        """Вкладка справки"""
        main_frame = tk.Frame(self.parent.help_tab, padx=20, pady=20)
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

        tk.Label(section1, text=HelpContent.get_file_requirements(), font=("Consolas", 9),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        # Раздел 2: Режимы работы
        section2 = tk.LabelFrame(scrollable_frame, text="⚙️ Режимы работы приложения",
                                font=("Arial", 11, "bold"), padx=15, pady=10, bg="white")
        section2.pack(fill=tk.X, pady=(0, 15))

        tk.Label(section2, text=HelpContent.get_modes_description(len(self.parent.methods)), font=("Consolas", 9),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        # Раздел 2.5: Алгоритм выбора лучшего метода
        section2_5 = tk.LabelFrame(scrollable_frame, text="🧮 Алгоритм автоматического выбора метода (v3.0)",
                                font=("Arial", 11, "bold"), padx=15, pady=10, bg="white")
        section2_5.pack(fill=tk.X, pady=(0, 15))

        tk.Label(section2_5, text=HelpContent.get_algorithm_description(), font=("Consolas", 8),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        # Раздел 3: Выбор столбцов
        section3 = tk.LabelFrame(scrollable_frame, text="🎯 Выбор столбцов для сравнения (v3.0)",
                                font=("Arial", 11, "bold"), padx=15, pady=10, bg="white")
        section3.pack(fill=tk.X, pady=(0, 15))

        tk.Label(section3, text=HelpContent.get_columns_selection(), font=("Consolas", 9),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        # Раздел 4: Экспорт результатов
        section4 = tk.LabelFrame(scrollable_frame, text="💾 Экспорт результатов",
                                font=("Arial", 11, "bold"), padx=15, pady=10, bg="white")
        section4.pack(fill=tk.X, pady=(0, 15))

        tk.Label(section4, text=HelpContent.get_export_options(), font=("Consolas", 9),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        # Раздел 5: Библиотеки
        section5 = tk.LabelFrame(scrollable_frame, text="📚 Используемые библиотеки",
                                font=("Arial", 11, "bold"), padx=15, pady=10, bg="white")
        section5.pack(fill=tk.X, pady=(0, 15))

        tk.Label(section5, text=HelpContent.get_libraries_description(), font=("Consolas", 9),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        # Раздел 6: Подробное описание методов
        section6 = tk.LabelFrame(scrollable_frame, text="🔍 Подробное описание всех методов сопоставления",
                                font=("Arial", 11, "bold"), padx=15, pady=10, bg="white")
        section6.pack(fill=tk.X, pady=(0, 15))

        tk.Label(section6, text=HelpContent.get_methods_detailed_description(), font=("Consolas", 8),
                justify=tk.LEFT, anchor="w", bg="white").pack(fill=tk.X)

        # Раздел 7: Расширенная нормализация
        section7 = tk.LabelFrame(scrollable_frame, text="🔬 Расширенная нормализация (v3.0)",
                                 font=("Arial", 11, "bold"), padx=15, pady=10)

        tk.Label(section7, text=HelpContent.get_normalization_description(),
                font=("Consolas", 9), justify=tk.LEFT, anchor="w").pack(anchor=tk.W, fill=tk.X)

        section7.pack(anchor=tk.W, fill=tk.X, pady=(0, 10))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def create_setup_tab(self):
        """Вкладка загрузки файлов с прокруткой"""
        # Создаем Canvas для прокрутки
        canvas = tk.Canvas(self.parent.setup_tab)
        scrollbar = tk.Scrollbar(self.parent.setup_tab, orient="vertical", command=canvas.yview)
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

        description_text = HelpContent.get_app_description(len(self.parent.methods))

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
        self.parent.askupo_label = tk.Label(askupo_frame, text="📂 Файл не выбран", fg="gray", font=("Arial", 9))
        self.parent.askupo_label.pack(anchor=tk.W, padx=20)
        tk.Button(askupo_frame, text="📁 Выбрать Источник данных 1 (целевой)", command=self.parent.select_askupo,
                 bg="#10B981", fg="white", font=("Arial", 9, "bold"),
                 padx=15, pady=5).pack(anchor=tk.W, padx=20, pady=3)

        eatool_frame = tk.Frame(files_frame)
        eatool_frame.pack(fill=tk.X, pady=5)
        tk.Label(eatool_frame, text="2️⃣ Источник данных 2:",
                font=("Arial", 10, "bold")).pack(anchor=tk.W)
        self.parent.eatool_label = tk.Label(eatool_frame, text="📂 Файл не выбран", fg="gray", font=("Arial", 9))
        self.parent.eatool_label.pack(anchor=tk.W, padx=20)
        tk.Button(eatool_frame, text="📁 Выбрать Источник данных 2", command=self.parent.select_eatool,
                 bg="#10B981", fg="white", font=("Arial", 9, "bold"),
                 padx=15, pady=5).pack(anchor=tk.W, padx=20, pady=3)

        settings_frame = tk.LabelFrame(main_frame, text="Настройки обработки",
                                       font=("Arial", 11, "bold"), padx=10, pady=10)
        settings_frame.pack(fill=tk.X, pady=(0, 20))

        mode_frame = tk.Frame(settings_frame)
        mode_frame.pack(fill=tk.X, pady=5)

        tk.Label(mode_frame, text="Режим работы:", font=("Arial", 10, "bold")).pack(anchor=tk.W)

        self.parent.mode_var = tk.StringVar(value="auto")

        tk.Radiobutton(mode_frame,
                      text="🤖 Автоматический - тестирует ВЫБРАННЫЕ методы и выбрав лучший создает результирующий эксель",
                      variable=self.parent.mode_var, value="auto",
                      font=("Arial", 9)).pack(anchor=tk.W, padx=20)
        tk.Radiobutton(mode_frame,
                      text="📊 Сравнение выбранных методов - тестирует на выборке (не более первых 200 записей) и выбирает лучший",
                      variable=self.parent.mode_var, value="compare",
                      font=("Arial", 9)).pack(anchor=tk.W, padx=20)
        tk.Radiobutton(mode_frame,
                      text="🔬 Полное сравнение - выбор одного или нескольких методов для создания результирующего эксель",
                      variable=self.parent.mode_var, value="full_compare",
                      font=("Arial", 9)).pack(anchor=tk.W, padx=20)

        # Фрейм для выбора методов
        self.parent.method_selector_frame = tk.Frame(settings_frame)
        self.parent.method_selector_frame.pack(fill=tk.X, pady=5)

        tk.Label(self.parent.method_selector_frame, text="Выберите методы (для всех режимов):",
                font=("Arial", 9, "bold")).pack(anchor=tk.W, padx=20)

        tk.Label(self.parent.method_selector_frame,
                text="💡 Удерживайте Ctrl для выбора нескольких методов",
                font=("Arial", 8), fg="gray").pack(anchor=tk.W, padx=20)

        # Listbox с прокруткой для выбора методов
        listbox_frame = tk.Frame(self.parent.method_selector_frame)
        listbox_frame.pack(anchor=tk.W, padx=20, pady=5)

        methods_scrollbar = tk.Scrollbar(listbox_frame)
        methods_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.parent.methods_listbox = tk.Listbox(listbox_frame,
                                          selectmode=tk.MULTIPLE,
                                          height=8,
                                          width=80,
                                          yscrollcommand=methods_scrollbar.set,
                                          exportselection=False)
        self.parent.methods_listbox.pack(side=tk.LEFT, fill=tk.BOTH)
        methods_scrollbar.config(command=self.parent.methods_listbox.yview)

        # Заполняем список методами
        for method in self.parent.methods:
            self.parent.methods_listbox.insert(tk.END, method.name)

        # Выбираем первый метод по умолчанию
        if self.parent.methods:
            self.parent.methods_listbox.selection_set(0)

        # Кнопка "Выбрать все методы"
        button_frame = tk.Frame(self.parent.method_selector_frame)
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

        self.parent.askupo_col_listbox = tk.Listbox(listbox_frame1, selectmode=tk.MULTIPLE,
                                             height=5, yscrollcommand=scrollbar1.set,
                                             exportselection=False)
        self.parent.askupo_col_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar1.config(command=self.parent.askupo_col_listbox.yview)
        self.parent.askupo_col_listbox.bind('<<ListboxSelect>>', self.on_askupo_column_select)

        tk.Checkbutton(source1_frame, text="Наследовать остальные столбцы источника 1",
                      variable=self.parent.inherit_askupo_cols_var,
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

        self.parent.eatool_col_listbox = tk.Listbox(listbox_frame2, selectmode=tk.MULTIPLE,
                                             height=5, yscrollcommand=scrollbar2.set,
                                             exportselection=False)
        self.parent.eatool_col_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar2.config(command=self.parent.eatool_col_listbox.yview)
        self.parent.eatool_col_listbox.bind('<<ListboxSelect>>', self.on_eatool_column_select)

        tk.Checkbutton(source2_frame, text="Наследовать остальные столбцы источника 2",
                      variable=self.parent.inherit_eatool_cols_var,
                      font=("Arial", 9)).pack(anchor=tk.W, pady=(5, 0))

        # Чекбокс для режима множественных столбцов
        tk.Checkbutton(columns_frame,
                      text="🔗 Режим сравнения по 2 столбцам одновременно (требует выбора 2 столбцов в каждом источнике)",
                      variable=self.parent.multi_column_mode_var,
                      font=("Arial", 9, "bold"), fg="#7C3AED").pack(anchor=tk.W, pady=(10, 0))

        # Описание режима множественных столбцов
        info_frame = tk.Frame(columns_frame, bg="#EFF6FF", bd=1, relief=tk.SOLID)
        info_frame.pack(anchor=tk.W, padx=40, pady=(5, 0), fill=tk.X)

        info_text = """ℹ️ РЕЖИМ КОНКАТЕНАЦИИ:

📝 Объединяет значения из выбранных столбцов в одну строку для сравнения
   Пример: "Microsoft" + "Office" = "Microsoft Office"

   ✅ Подходит для: vendor+product, категория+подкатегория, любые комбинации"""

        tk.Label(info_frame, text=info_text,
                font=("Consolas", 8), justify=tk.LEFT, anchor="w",
                bg="#EFF6FF", fg="#1E40AF").pack(fill=tk.X, padx=10, pady=5)

        tk.Label(columns_frame,
                text="💡 Подсказка: После выбора файлов, столбцы появятся в списках. Выберите 1-2 столбца для сравнения.",
                font=("Arial", 8), fg="gray", wraplength=700, justify=tk.LEFT).pack(anchor=tk.W, pady=(5, 0))

        # ==== СЕКЦИЯ: Расширенная нормализация ====
        normalization_frame = tk.LabelFrame(main_frame, text="🔬 Расширенная нормализация данных",
                                            font=("Arial", 11, "bold"), padx=15, pady=10, bg="#FFF7ED")
        normalization_frame.pack(fill=tk.X, pady=(10, 0))

        tk.Label(normalization_frame,
                text="⚙️ Опции предобработки текста перед сопоставлением (применяются ко ВСЕМ методам):",
                font=("Arial", 9, "bold"), bg="#FFF7ED", fg="#7C2D12").pack(anchor=tk.W, pady=(0, 10))

        # Создаём две колонки для чекбоксов
        checkbox_container = tk.Frame(normalization_frame, bg="#FFF7ED")
        checkbox_container.pack(fill=tk.X)

        # Левая колонка
        left_col = tk.Frame(checkbox_container, bg="#FFF7ED")
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        tk.Checkbutton(left_col,
                      text="🏢 Удалять юридические формы (ООО, Ltd, Inc, GmbH...)",
                      variable=self.parent.norm_remove_legal_var,
                      font=("Arial", 9), bg="#FFF7ED", anchor="w").pack(anchor=tk.W, pady=2)

        tk.Checkbutton(left_col,
                      text="🔢 Удалять версии (2021, v4.x, R2, SP1, x64, Win10...)",
                      variable=self.parent.norm_remove_versions_var,
                      font=("Arial", 9), bg="#FFF7ED", anchor="w").pack(anchor=tk.W, pady=2)

        tk.Checkbutton(left_col,
                      text="📝 Удалять стоп-слова (и, в, на, the, a, and...)",
                      variable=self.parent.norm_remove_stopwords_var,
                      font=("Arial", 9), bg="#FFF7ED", anchor="w").pack(anchor=tk.W, pady=2)

        # Правая колонка
        right_col = tk.Frame(checkbox_container, bg="#FFF7ED")
        right_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        tk.Checkbutton(right_col,
                      text="🌍 Транслитерация кириллицы → латиница (Фотошоп → Fotoshop)",
                      variable=self.parent.norm_transliterate_var,
                      font=("Arial", 9), bg="#FFF7ED", anchor="w").pack(anchor=tk.W, pady=2)

        tk.Checkbutton(right_col,
                      text="🔤 Удалять пунктуацию (!@#$%^&*...)",
                      variable=self.parent.norm_remove_punctuation_var,
                      font=("Arial", 9), bg="#FFF7ED", anchor="w").pack(anchor=tk.W, pady=2)

        # Подсказка
        hint_text = """💡 РЕКОМЕНДАЦИЯ: Включите все опции для максимальной точности.
Эти преобразования повышают качество сопоставления на 30-50%!

Пример: "ООО 1С Предприятие 8.3 x64" → "predprijatie" (после всех преобразований)"""

        tk.Label(normalization_frame,
                text=hint_text,
                font=("Arial", 8), bg="#FFF7ED", fg="#92400E",
                wraplength=700, justify=tk.LEFT).pack(anchor=tk.W, pady=(10, 0))

        # Кнопка "Применить все опции"
        button_frame = tk.Frame(normalization_frame, bg="#FFF7ED")
        button_frame.pack(anchor=tk.W, pady=(10, 0))

        tk.Button(button_frame, text="✓ Включить все опции (рекомендуется)",
                 command=self.enable_all_normalization,
                 bg="#16A34A", fg="white",
                 font=("Arial", 9, "bold"), padx=15, pady=5).pack(side=tk.LEFT, padx=5)

        tk.Button(button_frame, text="✗ Отключить все опции",
                 command=self.disable_all_normalization,
                 bg="#DC2626", fg="white",
                 font=("Arial", 9, "bold"), padx=15, pady=5).pack(side=tk.LEFT, padx=5)

        self.parent.process_btn = tk.Button(main_frame, text="🚀 Начать обработку",
                 command=self.parent.start_processing, bg="#7C3AED", fg="white",
                 font=("Arial", 13, "bold"), padx=50, pady=12,
                 state=tk.DISABLED)
        self.parent.process_btn.pack(pady=20)

    def create_comparison_tab(self):
        """Вкладка сравнения методов"""
        frame = tk.Frame(self.parent.comparison_tab, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="📊 Сравнение производительности методов",
                font=("Arial", 13, "bold")).pack(pady=10)

        # Создаем Treeview для сравнения методов
        tree_widget = TreeviewWithScrollbar(
            frame,
            columns=("rank", "method", "library", "perfect", "high", "avg_score", "time"),
            headers=[
                ("rank", "🏆", 50),
                ("method", "Метод", 300),
                ("library", "Библиотека", 120),
                ("perfect", "100%", 80),
                ("high", "90-99%", 80),
                ("avg_score", "Средний %", 100),
                ("time", "Время", 90),
            ],
            height=15,
            horizontal_scroll=True
        )
        tree_widget.pack(fill=tk.BOTH, expand=True)
        self.parent.comparison_tree = tree_widget.tree

        btn_frame = tk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=10)

        tk.Button(btn_frame, text="💾 Экспортировать сравнение в Excel",
                 command=self.parent.export_comparison, bg="#3B82F6", fg="white",
                 font=("Arial", 10, "bold"), padx=20, pady=5).pack(side=tk.LEFT, padx=5)

    def create_results_tab(self):
        """Вкладка результатов"""
        frame = tk.Frame(self.parent.results_tab, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        self.parent.result_info_frame = tk.Frame(frame)
        self.parent.result_info_frame.pack(fill=tk.X, pady=(0, 10))

        self.parent.result_stats_frame = tk.Frame(frame)
        self.parent.result_stats_frame.pack(fill=tk.X, pady=(0, 10))

        export_frame = tk.Frame(frame)
        export_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(export_frame, text="💾 Экспорт результатов в Excel:",
                font=("Arial", 11, "bold")).pack(anchor=tk.W)

        btn_container = tk.Frame(export_frame)
        btn_container.pack(fill=tk.X, pady=5)

        export_buttons = [
            ("📊 Полный отчет", self.parent.export_full, "#4F46E5"),
            ("✅ Точные (100%)", self.parent.export_perfect, "#10B981"),
            ("⚠️ Требуют проверки (<90%)", self.parent.export_problems, "#F59E0B"),
            ("❌ Без совпадений (0%)", self.parent.export_no_match, "#EF4444"),
        ]

        for text, command, color in export_buttons:
            tk.Button(btn_container, text=text, command=command, bg=color, fg="white",
                     font=("Arial", 10, "bold"), padx=18, pady=6).pack(side=tk.LEFT, padx=3)

        tk.Label(frame, text="📋 Результаты сопоставления (первые 50 записей):",
                font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))

        # Создаем Treeview для результатов
        tree_widget = TreeviewWithScrollbar(
            frame,
            columns=("num", "askupo", "eatool", "percent"),
            headers=[
                ("num", "№", 50),
                ("askupo", "Источник 1 (сравниваемый столбец)", 350),
                ("eatool", "Источник 2 (сопоставленный столбец)", 350),
                ("percent", "Процент совпадения", 120),
            ],
            height=15
        )
        tree_widget.pack(fill=tk.BOTH, expand=True)
        self.parent.results_tree = tree_widget.tree

    # ========== Вспомогательные методы для UI ==========

    def on_askupo_column_select(self, event):
        """Обработчик выбора столбцов из источника 1"""
        selected_indices = self.parent.askupo_col_listbox.curselection()
        self.parent.selected_askupo_cols = [self.parent.askupo_columns[i] for i in selected_indices]

        # КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Синхронизируем с data_manager!
        self.parent.data_manager.selected_source1_cols = self.parent.selected_askupo_cols

        # Ограничение: максимум 2 столбца
        if len(selected_indices) > 2:
            messagebox.showwarning("Предупреждение",
                                 "Можно выбрать максимум 2 столбца.\n"
                                 "Последний выбор будет отменен.")
            # Отменяем последний выбор
            self.parent.askupo_col_listbox.selection_clear(selected_indices[-1])
            self.parent.selected_askupo_cols = self.parent.selected_askupo_cols[:-1]
            # Синхронизируем изменение с data_manager
            self.parent.data_manager.selected_source1_cols = self.parent.selected_askupo_cols

        # АВТО-РЕЖИМ: Автоматически включаем галку если выбрано 2 столбца в ОБОИХ источниках
        if len(self.parent.selected_askupo_cols) == 2 and len(self.parent.selected_eatool_cols) == 2:
            self.parent.multi_column_mode_var.set(True)
        elif len(self.parent.selected_askupo_cols) == 1 or len(self.parent.selected_eatool_cols) == 1:
            self.parent.multi_column_mode_var.set(False)

    def on_eatool_column_select(self, event):
        """Обработчик выбора столбцов из источника 2"""
        selected_indices = self.parent.eatool_col_listbox.curselection()
        self.parent.selected_eatool_cols = [self.parent.eatool_columns[i] for i in selected_indices]

        # КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Синхронизируем с data_manager!
        self.parent.data_manager.selected_source2_cols = self.parent.selected_eatool_cols

        # Ограничение: максимум 2 столбца
        if len(selected_indices) > 2:
            messagebox.showwarning("Предупреждение",
                                 "Можно выбрать максимум 2 столбца.\n"
                                 "Последний выбор будет отменен.")
            # Отменяем последний выбор
            self.parent.eatool_col_listbox.selection_clear(selected_indices[-1])
            self.parent.selected_eatool_cols = self.parent.selected_eatool_cols[:-1]
            # Синхронизируем изменение с data_manager
            self.parent.data_manager.selected_source2_cols = self.parent.selected_eatool_cols

        # АВТО-РЕЖИМ: Автоматически включаем галку если выбрано 2 столбца в ОБОИХ источниках
        if len(self.parent.selected_askupo_cols) == 2 and len(self.parent.selected_eatool_cols) == 2:
            self.parent.multi_column_mode_var.set(True)
        elif len(self.parent.selected_askupo_cols) == 1 or len(self.parent.selected_eatool_cols) == 1:
            self.parent.multi_column_mode_var.set(False)

    def select_all_methods(self):
        """Выбрать все методы в списке"""
        self.parent.methods_listbox.selection_set(0, tk.END)

    def deselect_all_methods(self):
        """Снять выбор всех методов"""
        self.parent.methods_listbox.selection_clear(0, tk.END)

    def enable_all_normalization(self):
        """Включить все опции расширенной нормализации"""
        self.parent.norm_remove_legal_var.set(True)
        self.parent.norm_remove_versions_var.set(True)
        self.parent.norm_remove_stopwords_var.set(True)
        self.parent.norm_transliterate_var.set(True)
        self.parent.norm_remove_punctuation_var.set(True)
        messagebox.showinfo("✓ Опции нормализации",
                           "Все опции расширенной нормализации включены!\n\n"
                           "Это повысит качество сопоставления на 30-50%.")

    def disable_all_normalization(self):
        """Отключить все опции расширенной нормализации"""
        self.parent.norm_remove_legal_var.set(False)
        self.parent.norm_remove_versions_var.set(False)
        self.parent.norm_remove_stopwords_var.set(False)
        self.parent.norm_transliterate_var.set(False)
        self.parent.norm_remove_punctuation_var.set(False)
        messagebox.showinfo("✗ Опции нормализации",
                           "Все опции расширенной нормализации отключены.\n\n"
                           "Будет использоваться только базовая нормализация (lowercase + trim).")
