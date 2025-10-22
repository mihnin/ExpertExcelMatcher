# -*- mode: python ; coding: utf-8 -*-
# Оптимизированная конфигурация PyInstaller для минимального размера .exe
# Создано: 2025-10-22

block_cipher = None

# Список модулей для исключения (уменьшает размер)
# ВАЖНО: Не исключаем модули, которые могут быть нужны setuptools/pkg_resources
excludes = [
    # Matplotlib (не используется)
    'matplotlib',
    'matplotlib.pyplot',

    # Scipy (не используется)
    'scipy',

    # IPython/Jupyter (не используется)
    'IPython',
    'jupyter',
    'notebook',

    # Testing frameworks (не нужны в production)
    'pytest',
    '_pytest',

    # PIL/Pillow (не используется)
    'PIL',
    'Pillow',

    # Другие точно ненужные модули
    'pycparser',
]

# Скрытые импорты (если нужно)
hiddenimports = [
    'openpyxl.cell._writer',
    'xlsxwriter.utility',
    'pkg_resources.extern',
    'pkg_resources.py2_warn',
]

a = Analysis(
    ['expert_matcher.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Добавляем src/ как data files
        ('src', 'src'),
    ],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,  # Без оптимизации байткода (совместимость с numpy + Python 3.13)
)

# Фильтруем ненужные binaries
a.binaries = [x for x in a.binaries if not any(
    excluded in x[0].lower() for excluded in [
        'api-ms-win',  # Windows API DLLs (не всегда нужны)
        'vcruntime140_1',  # Дополнительные runtime (если не используются)
    ]
)]

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ExpertExcelMatcher',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,  # Удаляем отладочные символы
    upx=True,    # Используем UPX сжатие
    upx_exclude=[
        # Не сжимать некоторые критичные DLL
        'vcruntime140.dll',
        'python3.dll',
        'python313.dll',
    ],
    runtime_tmpdir=None,
    console=False,  # GUI приложение
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Можно добавить иконку если есть
)
