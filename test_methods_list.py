# -*- coding: utf-8 -*-
"""
Проверка: какие методы зарегистрированы в приложении
"""

import sys
import io

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Проверяем доступность библиотек
print("=" * 80)
print("ПРОВЕРКА ДОСТУПНОСТИ БИБЛИОТЕК")
print("=" * 80)

try:
    from rapidfuzz import fuzz, process
    RAPIDFUZZ_AVAILABLE = True
    print("✅ RapidFuzz: ДОСТУПЕН")
except ImportError:
    RAPIDFUZZ_AVAILABLE = False
    print("❌ RapidFuzz: НЕ УСТАНОВЛЕН")

try:
    import textdistance
    TEXTDISTANCE_AVAILABLE = True
    print("✅ TextDistance: ДОСТУПЕН")
except ImportError:
    TEXTDISTANCE_AVAILABLE = False
    print("❌ TextDistance: НЕ УСТАНОВЛЕН")

try:
    import jellyfish
    JELLYFISH_AVAILABLE = True
    print("✅ Jellyfish: ДОСТУПЕН")
except ImportError:
    JELLYFISH_AVAILABLE = False
    print("❌ Jellyfish: НЕ УСТАНОВЛЕН")

try:
    from transliterate import translit
    TRANSLITERATE_AVAILABLE = True
    print("✅ Transliterate: ДОСТУПЕН")
except ImportError:
    TRANSLITERATE_AVAILABLE = False
    print("❌ Transliterate: НЕ УСТАНОВЛЕН")

try:
    from FlagEmbedding import BGEM3FlagModel
    BGE_AVAILABLE = True
    print("✅ FlagEmbedding (BGE-M3): ДОСТУПЕН")
except ImportError as e:
    BGE_AVAILABLE = False
    print(f"❌ FlagEmbedding: НЕ УСТАНОВЛЕН")
    print(f"   Ошибка: {e}")

print("\n" + "=" * 80)
print("ПРОВЕРКА: BGE_AVAILABLE в приложении")
print("=" * 80)

if BGE_AVAILABLE:
    print("✅ BGE-M3 метод ДОЛЖЕН быть зарегистрирован в приложении")
    print("\nВ списке методов GUI должна быть строка:")
    print("   '🧠 BGE-M3: Semantic Embeddings (ML)'")
else:
    print("❌ BGE-M3 метод НЕ будет зарегистрирован (библиотека недоступна)")

print("\n" + "=" * 80)
print("РЕКОМЕНДАЦИЯ")
print("=" * 80)
print("\n1. Откройте приложение Expert Excel Matcher")
print("2. Посмотрите в список методов")
print("3. Найдите строку: '🧠 BGE-M3: Semantic Embeddings (ML)'")
print("\nЕсли метода НЕТ в списке:")
print("   → BGE_AVAILABLE = False")
print("   → Проверьте установку: pip install FlagEmbedding")
print("\nЕсли метод ЕСТЬ в списке, но выдаёт 0:")
print("   → Проблема в другом месте (нужна дополнительная диагностика)")
