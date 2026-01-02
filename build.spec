# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_submodules
import os

block_cipher = None

# Список файлов данных
data_files = [
    ('Матрица.xlsx', '.'), 
    ('Прайс-лист.xlsx', '.'), 
    ('Расчет мощностей METEOR.xlsx', '.'), 
    ('Формуляр для регистрации проектов.xlsm', '.'), 
    ('icon.ico', '.'),  
    ('Lagar.png', '.'),
    ('inst.pdf', '.'),
    # Изображения для подсказок
    ('1.png', '.'),
    ('2.png', '.'),
    ('3.png', '.'),
    # Файлы данных
    ('patterns.json', '.'),
]

# Добавляем дополнительные файлы из папки, если они существуют
additional_files = [
    'correspondence_manager.py',
    'column_selector.py', 
    'meteor_selector.py',
    'pattern_manager.py',
    'pdf_page_selector.py',
    'parsers.py',
    'debug_tools.py',
    'pdf_parser.py',
    'spec_generator.py',
    'interface_builder.py',
]

for file in additional_files:
    if os.path.exists(file):
        data_files.append((file, '.'))

# Собираем все hidden imports
hiddenimports = [
    'pandas', 
    'openpyxl',
    'openpyxl.reader.excel',
    'openpyxl.workbook',
    'openpyxl.worksheet',
    'openpyxl.cell',
    'openpyxl.styles',
    'tkinter',
    'tkinter.ttk',
    'pyperclip',
    'json',
    'numpy',
    'xlrd',
    'chardet',
    're',
    'datetime',
    'tempfile',
    'platform',
    'subprocess',
    'webbrowser',
    'traceback',
    'typing',
    'collections',
    'string',
    'sys',
    'os',
    'queue',
    'threading',
    'io',
    'PIL',
    'PIL.Image',
    'PIL.ImageTk',
    'PIL._imaging',
    'PIL._imagingtk',
    'fitz',
    'pymupdf',
]

# Собираем все подмодули для ключевых библиотек
try:
    # Собираем все подмодули openpyxl
    openpyxl_submodules = collect_submodules('openpyxl')
    hiddenimports.extend(openpyxl_submodules)
    print(f"Добавлены подмодули openpyxl: {len(openpyxl_submodules)}")
except Exception as e:
    print(f"Предупреждение при сборе подмодулей openpyxl: {e}")

try:
    # Данные Pillow (для работы с изображениями)
    pillow_datas = collect_data_files('PIL')
    for dest, src in pillow_datas:
        data_files.append((dest, src))
    print("Добавлены данные Pillow")
except Exception as e:
    print(f"Предупреждение: Не удалось собрать данные Pillow: {e}")

# Добавляем данные PyMuPDF (fitz) вручную
try:
    import fitz
    # Получаем путь к установленному PyMuPDF
    import pymupdf
    fitz_path = os.path.dirname(pymupdf.__file__)
    # Добавляем все файлы из папки fitz
    for root, dirs, files in os.walk(fitz_path):
        for file in files:
            if file.endswith(('.py', '.so', '.dll', '.pyd', '.dat')):
                src_path = os.path.join(root, file)
                rel_path = os.path.relpath(src_path, os.path.dirname(fitz_path))
                data_files.append((src_path, os.path.dirname(rel_path)))
    print(f"Добавлены данные PyMuPDF из: {fitz_path}")
except ImportError as e:
    print(f"Предупреждение: PyMuPDF не установлен или ошибка: {e}")

a = Analysis(
    ['start_v10.2013.py'],  # ← Убедитесь, что это правильное имя файла
    pathex=[os.getcwd()],  
    binaries=[],
    datas=data_files,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 
        'scipy', 
        'pytest',
        'PyQt5',
        'PySide2',
        'wx',
        'test',
        'unittest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='NexTT-1',  
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    runtime_tmpdir=None,
    console=False,  # Измените на True для отладки
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)