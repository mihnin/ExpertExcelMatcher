"""
Константы приложения Expert Excel Matcher

Этот модуль содержит все константы, используемые в приложении:
- AppConstants: Основные константы приложения
- NormalizationConstants: Константы для нормализации текста
"""


class AppConstants:
    """Константы приложения"""

    # Версия
    VERSION = "3.0.0"
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


class NormalizationConstants:
    """Константы для расширенной нормализации текста"""

    # Стоп-слова (русские + английские)
    RU_STOP = {
        "и", "в", "во", "не", "на", "но", "при", "для", "к", "из", "от",
        "с", "со", "о", "а", "у", "по", "над", "под", "до", "без", "или"
    }
    EN_STOP = {
        "the", "a", "an", "and", "or", "of", "for", "in", "on", "at", "to",
        "from", "with", "by", "without", "into", "out", "over", "under",
        "above", "below"
    }
    STOP_WORDS = RU_STOP | EN_STOP

    # Юридические префиксы (организационно-правовые формы)
    LEGAL_PREFIXES = [
        r'\bООО\b', r'\bАО\b', r'\bЗАО\b', r'\bИП\b', r'\bПАО\b', r'\bГК\b',
        r'\bНКО\b', r'\bНПО\b', r'\bНПП\b', r'\bНПФ\b', r'\bОАО\b',
        r'\bLtd\.?\b', r'\bLimited\b', r'\bInc\.?\b', r'\bLLC\b', r'\bGmbH\b',
        r'\bCorp\.?\b', r'\bCo\.?\b', r'\bSARL\b', r'\bS\.?A\.?\b',
        r'\bPLC\b', r'\bGroup\b', r'\bCompany\b', r'\bКомпания\b',
        r'\bИндивидуальный предприниматель\b',
        r'\bОбщество с ограниченной ответственностью\b'
    ]

    # Паттерны версий (годы, номера версий, архитектуры)
    VERSION_PATTERNS = [
        r'\b(19|20)\d{2}\b',                    # Годы: 2019, 2021, 1995
        r'\b[vV]\.?\d+\.[xX]\b',                # v4.x, V.7.X, v.2.x
        r'\b\d+\.[xX]\b',                       # 8.x, 12.x
        r'\b[vV]\.?\d+(\.\d+)*[a-zA-Z]*\b',     # v.4, v4, v.1.2, v1.2.3a, V2.0
        r'\b\d+\.\d+(\.\d+)*[a-zA-Z]*\b',       # 8.1, 2021.1a, 3.14.15
        r'\bR\d+\b',                            # R2, R12, R2023
        r'\bSP\d+\b',                           # SP1, SP2
        r'\b(x64|x86|64[-\s]?bit|32[-\s]?bit)\b',  # x64, x86, 64-bit, 32 bit
        r'\b(Win|Windows|Linux|Mac|MacOS)\s*\d*\b',  # Win10, Windows 11, MacOS
    ]
