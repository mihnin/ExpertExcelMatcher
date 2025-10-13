"""
Тест улучшений алгоритма сопоставления
"""
from expert_matcher import MatchingMethod
from rapidfuzz import fuzz

# Создаем метод Partial Ratio (который давал 100% для R)
method = MatchingMethod("Test Partial Ratio", fuzz.partial_ratio, "rapidfuzz",
                       use_process=True, scorer=fuzz.partial_ratio)

# Тестовые данные
query = "R"
choices = ["proficy ifix", "nginx", "sap bis", "r", "studio"]
choice_dict = {choice: choice.upper() for choice in choices}

print(f"Запрос: '{query}' (длина: {len(query)})\n")
print("Результаты сопоставления:")
print("-" * 60)

best_match, best_score = method.find_best_match(query.lower(), choices, choice_dict)

print(f"Лучшее совпадение: '{best_match}'")
print(f"Балл: {best_score:.1f}%")
print()

# Проверим каждый вариант отдельно
print("Детальная проверка всех вариантов:")
print("-" * 60)
for choice in choices:
    # Симулируем процесс
    from rapidfuzz import process
    result = process.extractOne(query.lower(), [choice], scorer=fuzz.partial_ratio)
    if result:
        raw_score = result[1]

        # Применяем штраф
        query_len = len(query)
        choice_len = len(choice)
        length_ratio = min(query_len, choice_len) / max(query_len, choice_len)

        if query_len <= 3 or choice_len <= 3:
            length_penalty = length_ratio ** 2
        else:
            length_penalty = length_ratio ** 0.5

        adjusted_score = raw_score * length_penalty

        print(f"'{choice}' (длина: {choice_len})")
        print(f"  Исходный балл: {raw_score:.1f}%")
        print(f"  Соотношение длин: {length_ratio:.3f}")
        print(f"  Штраф: {length_penalty:.3f}")
        print(f"  Финальный балл: {adjusted_score:.1f}%")
        print(f"  Проходит порог 50%: {'YES' if adjusted_score >= 50 else 'NO'}")
        print()

print("\n" + "="*60)
print("CONCLUSION:")
print("="*60)
if best_score < 50:
    print("OK: Short string 'R' correctly did NOT match long names")
    print("OK: Algorithm works correctly!")
else:
    print(f"ERROR: Found match: '{best_match}' with score {best_score:.1f}%")
    print("ERROR: Need to strengthen penalty")
