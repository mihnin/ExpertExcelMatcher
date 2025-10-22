"""
Переиспользуемые UI компоненты для Expert Excel Matcher

Этот модуль содержит классы и функции для создания стандартных
UI элементов, устраняя дублирование кода в основном приложении.
"""

import tkinter as tk
from tkinter import ttk
from typing import Callable, List, Tuple, Optional


class ScrollableFrame(tk.Frame):
    """
    Фрейм с вертикальной прокруткой

    Использование:
        container = ScrollableFrame(parent)
        container.pack(fill="both", expand=True)

        # Добавляйте виджеты в container.scrollable_frame
        tk.Label(container.scrollable_frame, text="Hello").pack()
    """

    def __init__(self, parent, **kwargs):
        """
        Инициализация прокручиваемого фрейма

        Args:
            parent: Родительский виджет
            **kwargs: Дополнительные параметры для Frame
        """
        super().__init__(parent, **kwargs)

        # Создаем Canvas и Scrollbar
        self.canvas = tk.Canvas(self)
        self.scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, padx=20, pady=20)

        # Связываем фрейм с Canvas
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Размещаем элементы
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Привязываем прокрутку колесиком мыши
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        """Обработчик прокрутки колесиком мыши"""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def destroy(self):
        """Корректное удаление виджета"""
        self.canvas.unbind_all("<MouseWheel>")
        super().destroy()


class TreeviewWithScrollbar(tk.Frame):
    """
    Treeview с вертикальной и горизонтальной прокруткой

    Использование:
        tree_widget = TreeviewWithScrollbar(
            parent,
            columns=("col1", "col2"),
            headers=[("col1", "Column 1", 100), ("col2", "Column 2", 200)],
            height=15
        )
        tree_widget.pack(fill="both", expand=True)

        # Доступ к Treeview через tree_widget.tree
        tree_widget.tree.insert("", "end", values=("value1", "value2"))
    """

    def __init__(self, parent, columns: Tuple[str, ...],
                 headers: List[Tuple[str, str, int]],
                 height: int = 15,
                 horizontal_scroll: bool = False,
                 **kwargs):
        """
        Инициализация Treeview с прокруткой

        Args:
            parent: Родительский виджет
            columns: Кортеж с именами столбцов
            headers: Список кортежей (column_id, header_text, width)
            height: Высота Treeview (количество строк)
            horizontal_scroll: Добавлять ли горизонтальную прокрутку
            **kwargs: Дополнительные параметры для Treeview
        """
        super().__init__(parent)

        # Вертикальная прокрутка
        self.scroll_y = ttk.Scrollbar(self)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # Горизонтальная прокрутка (опционально)
        self.scroll_x = None
        if horizontal_scroll:
            self.scroll_x = ttk.Scrollbar(self, orient=tk.HORIZONTAL)
            self.scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Создаем Treeview
        tree_kwargs = {
            'columns': columns,
            'show': 'headings',
            'yscrollcommand': self.scroll_y.set,
            'height': height
        }

        if horizontal_scroll and self.scroll_x:
            tree_kwargs['xscrollcommand'] = self.scroll_x.set

        tree_kwargs.update(kwargs)

        self.tree = ttk.Treeview(self, **tree_kwargs)

        # Связываем скроллбары
        self.scroll_y.config(command=self.tree.yview)
        if self.scroll_x:
            self.scroll_x.config(command=self.tree.xview)

        # Настраиваем заголовки и ширину столбцов
        for col_id, header_text, width in headers:
            self.tree.heading(col_id, text=header_text)
            # Определяем выравнивание: по центру для числовых, слева для текста
            anchor = tk.W if any(word in header_text.lower() for word in ['метод', 'источник', 'библиотека']) else tk.CENTER
            self.tree.column(col_id, width=width, anchor=anchor)

        self.tree.pack(fill=tk.BOTH, expand=True)


class MethodSelectorListbox(tk.Frame):
    """
    Listbox для выбора методов с прокруткой и кнопками выбора всех/снятия выбора

    Использование:
        selector = MethodSelectorListbox(
            parent,
            methods=["Method 1", "Method 2"],
            on_select_all=callback,
            on_deselect_all=callback
        )
        selector.pack(fill="both", expand=True)

        # Получение выбранных методов
        selected = selector.get_selected()
    """

    def __init__(self, parent, methods: List[str],
                 on_select_all: Optional[Callable] = None,
                 on_deselect_all: Optional[Callable] = None,
                 height: int = 10,
                 **kwargs):
        """
        Инициализация списка выбора методов

        Args:
            parent: Родительский виджет
            methods: Список названий методов
            on_select_all: Callback для кнопки "Выбрать все"
            on_deselect_all: Callback для кнопки "Снять выбор"
            height: Высота Listbox
            **kwargs: Дополнительные параметры для Listbox
        """
        super().__init__(parent)

        # Фрейм для Listbox с прокруткой
        listbox_frame = tk.Frame(self)
        listbox_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox_kwargs = {
            'selectmode': tk.MULTIPLE,
            'yscrollcommand': scrollbar.set,
            'height': height,
            'font': ("Consolas", 9),
            'activestyle': 'none'
        }
        listbox_kwargs.update(kwargs)

        self.listbox = tk.Listbox(listbox_frame, **listbox_kwargs)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)

        # Заполняем методами
        for method in methods:
            self.listbox.insert(tk.END, method)

        # Кнопки выбора (если передаW callbacks)
        if on_select_all or on_deselect_all:
            btn_frame = tk.Frame(self)
            btn_frame.pack(fill=tk.X, pady=(5, 0))

            if on_select_all:
                tk.Button(btn_frame, text="✓ Выбрать все", command=on_select_all,
                         bg="#10B981", fg="white", font=("Arial", 8)).pack(side=tk.LEFT, padx=2)

            if on_deselect_all:
                tk.Button(btn_frame, text="✗ Снять выбор", command=on_deselect_all,
                         bg="#EF4444", fg="white", font=("Arial", 8)).pack(side=tk.LEFT, padx=2)

    def get_selected(self) -> List[int]:
        """Получить индексы выбранных элементов"""
        return list(self.listbox.curselection())

    def select_all(self):
        """Выбрать все элементы"""
        self.listbox.select_set(0, tk.END)

    def deselect_all(self):
        """Снять выбор со всех элементов"""
        self.listbox.selection_clear(0, tk.END)


class FileSelectorWidget(tk.Frame):
    """
    Виджет для выбора файла: метка + кнопка

    Использование:
        file_selector = FileSelectorWidget(
            parent,
            label_text="1️⃣ Источник данных 1:",
            button_text="📁 Выбрать файл",
            on_select=callback
        )
        file_selector.pack(fill="x", pady=5)

        # Обновление метки после выбора
        file_selector.update_file_label("C:/path/to/file.xlsx")
    """

    def __init__(self, parent, label_text: str, button_text: str,
                 on_select: Callable,
                 button_color: str = "#10B981",
                 **kwargs):
        """
        Инициализация виджета выбора файла

        Args:
            parent: Родительский виджет
            label_text: Текст метки (название поля)
            button_text: Текст на кнопке
            on_select: Callback при нажатии кнопки
            button_color: Цвет кнопки (hex)
            **kwargs: Дополнительные параметры для Frame
        """
        super().__init__(parent, **kwargs)

        # Метка с названием поля
        tk.Label(self, text=label_text, font=("Arial", 10, "bold")).pack(anchor=tk.W)

        # Метка с выбранным файлом
        self.file_label = tk.Label(self, text="📂 Файл не выбран", fg="gray", font=("Arial", 9))
        self.file_label.pack(anchor=tk.W, padx=20)

        # Кнопка выбора файла
        tk.Button(
            self,
            text=button_text,
            command=on_select,
            bg=button_color,
            fg="white",
            font=("Arial", 9, "bold"),
            padx=15,
            pady=5
        ).pack(anchor=tk.W, padx=20, pady=3)

    def update_file_label(self, file_path: str):
        """
        Обновить метку с путем к файлу

        Args:
            file_path: Путь к выбранному файлу
        """
        if file_path:
            self.file_label.config(text=f"✅ {file_path}", fg="green")
        else:
            self.file_label.config(text="📂 Файл не выбран", fg="gray")


# ========== ФУНКЦИИ-ХЕЛПЕРЫ ДЛЯ СТИЛИЗАЦИИ ==========

def create_label_frame(parent, title: str, **kwargs) -> tk.LabelFrame:
    """
    Создать стилизованный LabelFrame

    Args:
        parent: Родительский виджет
        title: Заголовок фрейма
        **kwargs: Дополнительные параметры

    Returns:
        LabelFrame с примененным стилем
    """
    defaults = {
        'font': ("Arial", 11, "bold"),
        'padx': 10,
        'pady': 10
    }
    defaults.update(kwargs)

    return tk.LabelFrame(parent, text=title, **defaults)


def create_info_label_frame(parent, title: str, **kwargs) -> tk.LabelFrame:
    """
    Создать информационный LabelFrame (с голубым фоном)

    Args:
        parent: Родительский виджет
        title: Заголовок фрейма
        **kwargs: Дополнительные параметры

    Returns:
        LabelFrame с информационным стилем
    """
    defaults = {
        'font': ("Arial", 11, "bold"),
        'padx': 15,
        'pady': 15,
        'bg': "#F0F9FF"
    }
    defaults.update(kwargs)

    return tk.LabelFrame(parent, text=title, **defaults)


def create_styled_button(parent, text: str, command: Callable,
                        color: str = "#3B82F6", **kwargs) -> tk.Button:
    """
    Создать стилизованную кнопку

    Args:
        parent: Родительский виджет
        text: Текст кнопки
        command: Callback при нажатии
        color: Цвет фона (hex)
        **kwargs: Дополнительные параметры

    Returns:
        Стилизованная кнопка
    """
    defaults = {
        'bg': color,
        'fg': 'white',
        'font': ("Arial", 10, "bold"),
        'padx': 20,
        'pady': 5
    }
    defaults.update(kwargs)

    return tk.Button(parent, text=text, command=command, **defaults)


def create_title_header(parent, title: str, subtitle: str = "",
                       bg_color: str = "#7C3AED") -> tk.Frame:
    """
    Создать заголовок приложения (title + subtitle)

    Args:
        parent: Родительский виджет
        title: Основной заголовок
        subtitle: Подзаголовок (опционально)
        bg_color: Цвет фона (hex)

    Returns:
        Frame с заголовком
    """
    title_frame = tk.Frame(parent, bg=bg_color, pady=15)

    tk.Label(
        title_frame,
        text=title,
        font=("Arial", 18, "bold"),
        fg="white",
        bg=bg_color
    ).pack()

    if subtitle:
        tk.Label(
            title_frame,
            text=subtitle,
            font=("Arial", 10),
            fg="white",
            bg=bg_color
        ).pack()

    return title_frame
