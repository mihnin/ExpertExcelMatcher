# -*- coding: utf-8 -*-
"""
Тестирование новых функций нормализации
"""

import sys
import io
import re
from transliterate import translit

# Установить UTF-8 для вывода в Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Константы для нормализации
class NormalizationConstants:
    RU_STOP = {"и", "в", "во", "не", "на", "но", "при", "для", "к", "из", "от", "с", "со", "о", "а", "у", "по", "над", "под", "до", "без", "или"}
    EN_STOP = {"the", "a", "an", "and", "or", "of", "for", "in", "on", "at", "to", "from", "with", "by", "without", "into", "out", "over", "under", "above", "below"}
    STOP_WORDS = RU_STOP | EN_STOP

    LEGAL_PREFIXES = [
        r'\bООО\b', r'\bАО\b', r'\bЗАО\b', r'\bИП\b', r'\bПАО\b', r'\bГК\b',
        r'\bНКО\b', r'\bНПО\b', r'\bНПП\b', r'\bНПФ\b', r'\bОАО\b',
        r'\bLtd\.?\b', r'\bLimited\b', r'\bInc\.?\b', r'\bLLC\b', r'\bGmbH\b',
        r'\bCorp\.?\b', r'\bCo\.?\b', r'\bSARL\b', r'\bS\.?A\.?\b',
        r'\bPLC\b', r'\bGroup\b', r'\bCompany\b', r'\bКомпания\b',
    ]

    VERSION_PATTERNS = [
        r'\b(19|20)\d{2}\b',
        r'\b[vV]\.?\d+\.[xX]\b',
        r'\b\d+\.[xX]\b',
        r'\b[vV]\.?\d+(\.\d+)*[a-zA-Z]*\b',
        r'\b\d+\.\d+(\.\d+)*[a-zA-Z]*\b',
        r'\bR\d+\b',
        r'\bSP\d+\b',
        r'\b(x64|x86|64[-\s]?bit|32[-\s]?bit)\b',
        r'\b(Win|Windows|Linux|Mac|MacOS)\s*\d*\b',
    ]


def normalize_advanced(s, remove_legal=True, remove_versions=True,
                       remove_stopwords=True, transliterate_text=True,
                       remove_punctuation=True):
    """Расширенная нормализация как в Notebook"""
    if not s:
        return ""

    s = str(s).strip()

    # 1. Удаление юридических форм
    if remove_legal:
        for pattern in NormalizationConstants.LEGAL_PREFIXES:
            s = re.sub(pattern, ' ', s, flags=re.IGNORECASE)

    # 2. Удаление версий
    if remove_versions:
        for pattern in NormalizationConstants.VERSION_PATTERNS:
            s = re.sub(pattern, ' ', s, flags=re.IGNORECASE)

    # 3. Lowercase
    s = s.lower()

    # 4. Удаление пунктуации
    if remove_punctuation:
        s = re.sub(r'[^a-zа-яё0-9\s]', ' ', s)

    # 5. Удаление стоп-слов
    if remove_stopwords:
        words = s.split()
        words = [w for w in words if w and w not in NormalizationConstants.STOP_WORDS]
        s = ' '.join(words)

    # 6. Транслитерация
    if transliterate_text:
        if re.search(r'[а-яё]', s):
            try:
                s = translit(s, 'ru', reversed=True)
            except:
                pass

    # 7. Схлопывание пробелов
    s = re.sub(r'\s+', ' ', s).strip()

    return s


def normalize_basic(s):
    """Базовая нормализация (как было раньше)"""
    if not s:
        return ""
    s = str(s).lower().strip()
    s = re.sub(r'\s+', ' ', s)
    return s


# Тестовые кейсы
test_cases = [
    "ООО 1С Предприятие 8.3 x64",
    "Microsoft Office 2021 Professional",
    "Adobe Photoshop CC 2019",
    "Ltd Norton Antivirus v12.5",
    "AutoCAD 2023 SP1 Windows 10",
    "Oracle Database 19c Enterprise Edition",
    "ООО «Лаборатория Касперского» Kaspersky Endpoint Security 11.3",
    "Google Chrome 120.0.6099.129 64-bit",
    "Visual Studio Code v1.85.2",
    "Яндекс.Браузер 23.11.0.2401",
]

print("=" * 100)
print("ТЕСТИРОВАНИЕ РАСШИРЕННОЙ НОРМАЛИЗАЦИИ")
print("=" * 100)

for test in test_cases:
    basic = normalize_basic(test)
    advanced = normalize_advanced(test)

    print(f"\n{'─' * 100}")
    print(f"📝 ОРИГИНАЛ: {test}")
    print(f"{'─' * 100}")
    print(f"❌ БАЗОВАЯ:  {basic}")
    print(f"✅ РАСШИРЕННАЯ: {advanced}")
    print(f"   Длина: {len(test)} → {len(basic)} (базовая) → {len(advanced)} (расширенная)")

print(f"\n{'=' * 100}")
print("СРАВНЕНИЕ ЭФФЕКТИВНОСТИ")
print("=" * 100)

# Пример сопоставления
source = "ООО 1С Предприятие 8.3 x64"
target = "1C Enterprise"

print(f"\n🎯 ЗАДАЧА: Сопоставить '{source}' и '{target}'")
print("\n" + "─" * 100)

basic_s = normalize_basic(source)
basic_t = normalize_basic(target)
print(f"❌ БАЗОВАЯ НОРМАЛИЗАЦИЯ:")
print(f"   Источник: '{basic_s}'")
print(f"   Цель:     '{basic_t}'")
print(f"   Совпадение слов: {set(basic_s.split()) & set(basic_t.split())}")
print(f"   Результат: {'✓ СОВПАДЁТ' if '1c' in basic_s and '1c' in basic_t else '✗ НЕ СОВПАДЁТ'}")

print("\n" + "─" * 100)

advanced_s = normalize_advanced(source)
advanced_t = normalize_advanced(target)
print(f"✅ РАСШИРЕННАЯ НОРМАЛИЗАЦИЯ:")
print(f"   Источник: '{advanced_s}'")
print(f"   Цель:     '{advanced_t}'")
print(f"   Совпадение слов: {set(advanced_s.split()) & set(advanced_t.split())}")
print(f"   Результат: {'✓ СОВПАДЁТ' if 'enterprise' in advanced_s and 'enterprise' in advanced_t else '✗ НЕ СОВПАДЁТ'}")

print("\n" + "=" * 100)
print("✓ ТЕСТИРОВАНИЕ ЗАВЕРШЕНО")
print("=" * 100)
