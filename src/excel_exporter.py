"""
Модуль экспорта данных в Excel для Expert Excel Matcher

Этот модуль содержит класс ExcelExporter для экспорта результатов
сопоставления в форматированные Excel-файлы.
"""

import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict, List, Optional
from tkinter import messagebox, filedialog

from src.constants import AppConstants
from src.matching_engine import MatchingEngine


class ExcelExporter:
    """Класс для экспорта данных в Excel с форматированием"""

    def __init__(self, engine: MatchingEngine, results: Optional[pd.DataFrame] = None):
        """
        Инициализация экспортера

        Args:
            engine: Движок сопоставления (для расчета статистики)
            results: DataFrame с результатами (может быть None)
        """
        self.engine = engine
        self.results = results

    # ========== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ (устранение дублирования) ==========

    def _create_header_format(self, workbook):
        """Создание формата для заголовков"""
        return workbook.add_format({
            'bold': True,
            'bg_color': '#7C3AED',
            'font_color': 'white',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

    def _create_color_formats(self, workbook) -> Dict[int, any]:
        """Создание форматов цветовой раскраски по проценту"""
        return {
            100: workbook.add_format({'bg_color': '#D1FAE5', 'border': 1}),  # Зеленый
            90: workbook.add_format({'bg_color': '#DBEAFE', 'border': 1}),   # Синий
            70: workbook.add_format({'bg_color': '#FEF3C7', 'border': 1}),   # Желтый
            50: workbook.add_format({'bg_color': '#FED7AA', 'border': 1}),   # Оранжевый
            1: workbook.add_format({'bg_color': '#FFE4E1', 'border': 1}),    # Розовый
            0: workbook.add_format({'bg_color': '#FEE2E2', 'border': 1})     # Красный
        }

    def _apply_header_format(self, worksheet, columns, header_format):
        """Применение формата к заголовкам"""
        for col_num, value in enumerate(columns):
            worksheet.write(0, col_num, value, header_format)

    def _get_format_by_percent(self, percent: float, formats: Dict[int, any]):
        """Получение формата по проценту совпадения"""
        if percent == 100:
            return formats[100]
        elif percent >= 90:
            return formats[90]
        elif percent >= 70:
            return formats[70]
        elif percent >= 50:
            return formats[50]
        elif percent > 0:
            return formats[1]
        else:
            return formats[0]

    def _apply_color_coding(self, worksheet, df: pd.DataFrame, formats: Dict[int, any]):
        """
        Применение цветовой раскраски по проценту совпадения

        Args:
            worksheet: Лист Excel
            df: DataFrame с данными (должен содержать 'Процент совпадения')
            formats: Словарь форматов
        """
        for row_num in range(1, len(df) + 1):
            percent = df.iloc[row_num - 1][AppConstants.COL_PERCENT]
            fmt = self._get_format_by_percent(percent, formats)

            for col_num in range(len(df.columns)):
                worksheet.write(row_num, col_num,
                              df.iloc[row_num - 1, col_num], fmt)

    def _set_column_widths(self, worksheet, columns: List[str]):
        """
        Установка оптимальной ширины столбцов

        Args:
            worksheet: Лист Excel
            columns: Список названий столбцов
        """
        for col_num, col_name in enumerate(columns):
            if col_num == 0 and col_name == '№':
                worksheet.set_column(col_num, col_num, 8)  # Номер строки
            elif 'Источник данных' in str(col_name) or AppConstants.COL_SOURCE1_PREFIX in str(col_name) or AppConstants.COL_SOURCE2_PREFIX in str(col_name):
                worksheet.set_column(col_num, col_num, 45)  # Широкие столбцы для названий
            elif col_name == AppConstants.COL_PERCENT:
                worksheet.set_column(col_num, col_num, 12)  # Узкий для процента
            elif col_name == AppConstants.COL_METHOD:
                worksheet.set_column(col_num, col_num, 35)  # Средний для метода
            else:
                worksheet.set_column(col_num, col_num, 20)  # Остальные столбцы

    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Очистка DataFrame от NaN и inf для корректного экспорта

        Args:
            df: Исходный DataFrame

        Returns:
            Очищенный DataFrame
        """
        cleaned = df.copy()
        cleaned = cleaned.replace([np.nan, np.inf, -np.inf], "")
        return cleaned

    def _clean_sheet_name(self, name: str) -> str:
        """
        Очистка названия листа Excel от недопустимых символов

        Args:
            name: Исходное название

        Returns:
            Очищенное название (макс. 31 символ)
        """
        # Удаляем эмодзи (символы > 127)
        sheet_name = ''.join(char for char in name if ord(char) < 128)

        # Удаляем недопустимые символы для Excel
        invalid_chars = [':', '\\', '/', '?', '*', '[', ']']
        for char in invalid_chars:
            sheet_name = sheet_name.replace(char, '_')

        # Убираем лишние пробелы
        sheet_name = sheet_name.strip()

        # Ограничиваем длину (Excel лимит 31 символ)
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:28] + "..."

        # Если пустое, возвращаем дефолтное
        if not sheet_name:
            sheet_name = "Sheet1"

        return sheet_name

    def _add_row_numbers(self, df: pd.DataFrame) -> pd.DataFrame:
        """Добавление столбца с номерами строк"""
        result = df.copy()
        result.insert(0, '№', range(1, len(result) + 1))
        return result

    def _create_statistics_sheet(self, writer, workbook):
        """
        Создание листа со статистикой

        Args:
            writer: ExcelWriter
            workbook: Workbook для форматов
        """
        if self.results is None:
            return

        stats = self.engine.calculate_statistics(self.results)

        stats_data = pd.DataFrame([
            {'Категория': 'Всего записей', 'Количество': stats['total'], 'Процент': '100%'},
            {'Категория': '100% (точное совпадение)', 'Количество': stats['perfect'],
             'Процент': f"{stats['perfect']/stats['total']*100:.1f}%"},
            {'Категория': '90-99% (высокое совпадение)', 'Количество': stats['high'],
             'Процент': f"{stats['high']/stats['total']*100:.1f}%"},
            {'Категория': '70-89% (среднее совпадение)', 'Количество': stats['medium'],
             'Процент': f"{stats['medium']/stats['total']*100:.1f}%"},
            {'Категория': '50-69% (низкое совпадение)', 'Количество': stats['low'],
             'Процент': f"{stats['low']/stats['total']*100:.1f}%"},
            {'Категория': '1-49% (очень низкое совпадение)', 'Количество': stats['very_low'],
             'Процент': f"{stats['very_low']/stats['total']*100:.1f}%"},
            {'Категория': '0% (нет совпадения)', 'Количество': stats['none'],
             'Процент': f"{stats['none']/stats['total']*100:.1f}%"},
            {'Категория': '---', 'Количество': '---', 'Процент': '---'},
            {'Категория': 'Проверка суммы', 'Количество': stats['check_sum'],
             'Процент': 'OK' if stats['check_sum'] == stats['total'] else 'ОШИБКА!'}
        ])

        stats_data.to_excel(writer, sheet_name='Статистика', index=False)

    # ========== ОСНОВНЫЕ МЕТОДЫ ЭКСПОРТА ==========

    def export_results(self, data: pd.DataFrame, filename: str,
                      include_stats: bool = False) -> bool:
        """
        Базовая функция экспорта результатов сопоставления

        Args:
            data: DataFrame с результатами
            filename: Имя файла по умолчанию
            include_stats: Добавлять ли лист со статистикой

        Returns:
            True если экспорт успешен, False если отменен или ошибка
        """
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not save_path:
            return False

        try:
            data_to_export = self._clean_dataframe(data)
            data_to_export = self._add_row_numbers(data_to_export)

            with pd.ExcelWriter(save_path, engine='xlsxwriter',
                              engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
                data_to_export.to_excel(writer, sheet_name='Результаты', index=False)

                workbook = writer.book
                worksheet = writer.sheets['Результаты']

                # Применяем форматирование
                header_format = self._create_header_format(workbook)
                self._apply_header_format(worksheet, data_to_export.columns.values, header_format)

                # Устанавливаем ширину столбцов
                self._set_column_widths(worksheet, data_to_export.columns.values)

                # Применяем цветовую раскраску
                formats = self._create_color_formats(workbook)
                self._apply_color_coding(worksheet, data_to_export, formats)

                # Добавляем статистику если нужно
                if include_stats:
                    self._create_statistics_sheet(writer, workbook)

            messagebox.showinfo("Успех", f"✅ Файл сохранен:\n{save_path}")
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"❌ Ошибка при экспорте:\n{str(e)}")
            return False

    def export_comparison(self, methods_comparison: List[Dict],
                         filename: str = "Сравнение_методов_сопоставления.xlsx") -> bool:
        """
        Экспорт сравнения методов

        Args:
            methods_comparison: Список словарей со статистикой методов
            filename: Имя файла по умолчанию

        Returns:
            True если экспорт успешен, False если отменен или ошибка
        """
        if not methods_comparison:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return False

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not save_path:
            return False

        try:
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
                for i, stats in enumerate(methods_comparison)
            ])

            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Сравнение методов', index=False)

                workbook = writer.book
                worksheet = writer.sheets['Сравнение методов']

                # Форматирование заголовков
                header_format = self._create_header_format(workbook)
                self._apply_header_format(worksheet, df.columns.values, header_format)

                # Ширина столбцов
                worksheet.set_column('A:A', 10)
                worksheet.set_column('B:B', 40)
                worksheet.set_column('C:H', 18)

            messagebox.showinfo("Успех", f"✅ Файл сохранен:\n{save_path}")
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка экспорта:\n{str(e)}")
            return False

    def export_full_comparison(self, full_comparison_results: Dict,
                              filename: str = "Полное_сравнение_всех_методов.xlsx") -> bool:
        """
        Экспорт полного сравнения всех методов в многолистовой Excel

        Args:
            full_comparison_results: Словарь с результатами:
                - 'methods_data': Dict[method_name, DataFrame]
                - 'comparison_stats': List[Dict] со статистикой
            filename: Имя файла по умолчанию

        Returns:
            True если экспорт успешен, False если отменен или ошибка
        """
        if not full_comparison_results:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return False

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not save_path:
            return False

        try:
            methods_data = full_comparison_results['methods_data']
            comparison_stats = full_comparison_results['comparison_stats']

            # Очищаем все DataFrame
            cleaned_methods_data = {}
            for method_name, df in methods_data.items():
                cleaned_methods_data[method_name] = self._clean_dataframe(df)

            with pd.ExcelWriter(save_path, engine='xlsxwriter',
                              engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
                workbook = writer.book

                # Создаем форматы
                header_format = self._create_header_format(workbook)
                formats = self._create_color_formats(workbook)

                # 1. ЛИСТ "Сводка" - сравнительная таблица всех методов
                summary_df = pd.DataFrame([
                    {
                        'Место': i + 1,
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

                summary_df.to_excel(writer, sheet_name='Сводка', index=False)
                worksheet = writer.sheets['Сводка']

                self._apply_header_format(worksheet, summary_df.columns.values, header_format)

                worksheet.set_column('A:A', 10)
                worksheet.set_column('B:B', 40)
                worksheet.set_column('C:L', 15)

                # 2. ЛИСТЫ для каждого метода
                for method_name, results_df in cleaned_methods_data.items():
                    sheet_name = self._clean_sheet_name(method_name)

                    # Добавляем номера строк
                    export_df = self._add_row_numbers(results_df)

                    export_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    worksheet = writer.sheets[sheet_name]

                    # Заголовки
                    self._apply_header_format(worksheet, export_df.columns.values, header_format)

                    # Ширина столбцов
                    self._set_column_widths(worksheet, export_df.columns.values)

                    # Цветовая раскраска
                    self._apply_color_coding(worksheet, export_df, formats)

            messagebox.showinfo("Успех", f"✅ Полное сравнение сохранено!\n\n"
                              f"📁 Файл: {Path(save_path).name}\n"
                              f"📊 Листов: {len(cleaned_methods_data) + 1}\n"
                              f"   • Сводка: 1 лист\n"
                              f"   • Результаты методов: {len(cleaned_methods_data)} листов")
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"❌ Ошибка при экспорте:\n{str(e)}")
            return False
