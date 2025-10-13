# 🔬 ExpertExcelMatcher

Приложение для автоматического сопоставления названий программного обеспечения между двумя Excel базами данных с использованием алгоритмов нечёткого поиска (fuzzy matching).

---

## ⚡ Шпаргалка команд

| Действие | Команда |
|----------|---------|
| 🏃 **Запустить приложение** | `python expert_matcher.py` |
| 📦 **Установить зависимости** | `pip install -r requirements.txt` |
| 🔨 **Собрать .exe (первый раз)** | `pyinstaller --onefile --windowed --name "ExpertExcelMatcher" expert_matcher.py` |
| 🔨 **Пересобрать .exe (после изменений)** | `pyinstaller ExpertExcelMatcher.spec` |
| 🧪 **Запустить тесты** | `python test_improvements.py` |
| 🔍 **Проверить отчет** | `python check_report.py` |
| 🧹 **Очистить перед пересборкой** | `rmdir /s /q build dist` |

---

## 🚀 Быстрый старт

### 1. Клонировать репозиторий
```bash
git clone <your-repo-url>
cd ExpertExcelMatcher
```

### 2. Установить зависимости
```bash
pip install -r requirements.txt
```

### 3. Запустить приложение
```bash
python expert_matcher.py
```

## 📦 Требования

- Python 3.8 или выше
- Windows 10/11 (для GUI)
- Все зависимости указаны в `requirements.txt`

## 🛠️ Создание исполняемого файла (.exe)

### Первая сборка
```bash
# Установите PyInstaller (один раз)
pip install pyinstaller

# Соберите .exe
pyinstaller --onefile --windowed --name "ExpertExcelMatcher" expert_matcher.py
```

### Последующие сборки (после изменений)
```bash
# Просто используйте готовый .spec файл
pyinstaller ExpertExcelMatcher.spec
```

**Результат:** Готовый файл `dist/ExpertExcelMatcher.exe`

Подробная инструкция: см. `BUILD.md`

## 📚 Документация

- `CLAUDE.md` - техническая документация для разработчиков
- `BUILD.md` - инструкция по сборке .exe
- Встроенная справка - вкладка "❓ Справка" в приложении

## 🧪 Тестирование

```bash
# Тест механизма штрафа за длину
python test_improvements.py

# Проверка точности отчета
python check_report.py
```

## 📋 Входные данные

Приложение работает с Excel файлами (.xlsx/.xls):
- **Файл 1**: База АСКУПО
- **Файл 2**: База EA Tool

**Требование:** Первый столбец должен содержать названия ПО (текст)

## ✨ Возможности

- 🎯 Автоматический выбор лучшего алгоритма сопоставления
- 📊 Сравнение всех доступных методов (17+ алгоритмов)
- ⚡ Оптимизация с RapidFuzz (в 100 раз быстрее)
- 📈 Детальная статистика совпадений
- 💾 Экспорт результатов в цветной Excel
- 🔍 Механизм штрафа за разницу длин (предотвращает ложные совпадения)

## 🔧 Рабочий процесс: Изменение кода и пересборка

```
┌─────────────────────────────────────────────────────────────────┐
│                    WORKFLOW: Внесение изменений                  │
└─────────────────────────────────────────────────────────────────┘

1️⃣  Редактирование кода
    ├─ Откройте: expert_matcher.py
    └─ Внесите изменения

2️⃣  Тестирование через Python (БЫСТРО - 2 сек)
    └─ python expert_matcher.py

3️⃣  Сборка .exe (если всё работает)
    └─ pyinstaller ExpertExcelMatcher.spec  (30 сек)

4️⃣  Тестирование .exe
    └─ dist\ExpertExcelMatcher.exe

5️⃣  Git (опционально)
    ├─ git add expert_matcher.py
    ├─ git commit -m "Описание изменений"
    └─ git push origin main
```

### Сценарий: Вы внесли изменения в код и хотите создать новый .exe

#### Шаг 1: Внесите изменения
```bash
# Откройте файл в редакторе
code expert_matcher.py
# или
notepad expert_matcher.py

# Внесите изменения, сохраните файл
```

#### Шаг 2: Протестируйте изменения
```bash
# Запустите приложение напрямую через Python
python expert_matcher.py

# Проверьте, что всё работает корректно
# Убедитесь, что нет ошибок
```

#### Шаг 3: Пересоберите .exe
```bash
# Используйте готовый .spec файл (БЫСТРО - ~30 сек)
pyinstaller ExpertExcelMatcher.spec

# Готовый файл: dist/ExpertExcelMatcher.exe
```

#### Шаг 4: Протестируйте .exe
```bash
# Перейдите в папку dist
cd dist

# Запустите .exe
ExpertExcelMatcher.exe

# Проверьте, что изменения работают
```

#### Шаг 5: Сохраните в Git (опционально)
```bash
# Вернитесь в корневую папку
cd ..

# Добавьте изменения
git add expert_matcher.py

# Создайте коммит
git commit -m "Описание ваших изменений"

# Отправьте на GitHub
git push origin main
```

### 🎯 Быстрые команды для ежедневной работы

#### Вариант А: Разработка без .exe (БЫСТРО)
```bash
# Редактируйте код
# Сразу тестируйте
python expert_matcher.py
```

#### Вариант Б: Сборка .exe после изменений
```bash
# 1. Тестируем код
python expert_matcher.py

# 2. Если всё работает - собираем .exe
pyinstaller ExpertExcelMatcher.spec

# 3. Тестируем .exe
dist\ExpertExcelMatcher.exe
```

#### Вариант В: Сборка с очисткой (если что-то пошло не так)
```bash
# Удаляем старые артефакты
rmdir /s /q build dist

# Собираем заново
pyinstaller ExpertExcelMatcher.spec
```

### 🐛 Отладка проблем при сборке

#### Проблема: .exe не запускается
**Решение:**
```bash
# Соберите БЕЗ флага --windowed (покажет ошибки в консоли)
pyinstaller --onefile --name "ExpertExcelMatcher" expert_matcher.py

# Запустите из командной строки
dist\ExpertExcelMatcher.exe
```

#### Проблема: Долгая сборка
**Решение:**
```bash
# Используйте .spec файл (кэширование)
pyinstaller ExpertExcelMatcher.spec
# Вместо полной пересборки с нуля
```

#### Проблема: Изменения не попали в .exe
**Решение:**
```bash
# Очистите кэш PyInstaller
pyinstaller --clean ExpertExcelMatcher.spec
```

### 📊 Запуск тестов после изменений

```bash
# Тест механизма штрафа за длину
python test_improvements.py

# Проверка отчета (нужен готовый отчет)
python check_report.py
```

## 🤝 Разработка

См. `CLAUDE.md` для подробной информации об архитектуре и ключевых компонентах.

### Структура проекта
```
ExpertExcelMatcher/
├── expert_matcher.py          # Основной код приложения
├── check_report.py            # Утилита проверки отчетов
├── test_improvements.py       # Тесты алгоритма
├── requirements.txt           # Зависимости Python
├── ExpertExcelMatcher.spec    # Конфигурация PyInstaller
├── README.md                  # Эта инструкция
├── CLAUDE.md                  # Техническая документация
├── BUILD.md                   # Детали сборки .exe
├── .gitignore                 # Исключения Git
├── build/                     # [Игнорируется] Временные файлы сборки
└── dist/                      # [Игнорируется] Готовый .exe файл
```

## 📄 Лицензия

[Укажите вашу лицензию] 
