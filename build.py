import os
import shutil
from distutils.core import setup
from Cython.Build import cythonize
from PyInstaller.__main__ import run

# 1. Очистка
shutil.rmtree('build', ignore_errors=True)
shutil.rmtree('dist', ignore_errors=True)
shutil.rmtree('obfuscated', ignore_errors=True)
os.makedirs('obfuscated', exist_ok=True)

# 2. Переименование (если нужно)
if os.path.exists("start_v6.7.py") and not os.path.exists("start_v6_7.py"):
    os.rename("start_v6.7.py", "start_v6_7.py")

# 3. Обфускация через Cython
print("🔒 Obfuscating with Cython...")
setup(
    script_args=["build_ext", "--inplace"],
    ext_modules=cythonize(
        "start_v6_7.py",
        compiler_directives={'language_level': "3"}
    )
)

# 4. Копируем результат в папку obfuscated
pyd_file = "start_v6_7.cp312-win_amd64.pyd"  # Имя сгенерированного .pyd
if os.path.exists(pyd_file):
    shutil.move(pyd_file, "obfuscated/start_v6_7.pyd")

# 5. Создаем минимальный загрузочный скрипт
with open("obfuscated/__main__.py", "w") as f:
    f.write("from start_v6_7 import *\n")

# 6. Копируем ресурсы
for file in ['Матрица.xlsx', 'Lagar.png', 'icon.ico', 'favicon.ico', 
             'Прайс-лист.xlsx', 'Формуляр для регистрации проектов.xlsm']:
    if os.path.exists(file):
        shutil.copy(file, 'obfuscated')

# 7. Сборка EXE
print("⚙️ Building EXE...")
run([
    '--onefile',
    '--windowed',
    '--icon=favicon.ico',
    '--add-data=Матрица.xlsx;.',
    '--add-data=Lagar.png;.',
    '--add-data=icon.ico;.',
    '--add-data=favicon.ico;.',
    '--add-data=Прайс-лист.xlsx;.',
    '--add-data=Формуляр для регистрации проектов.xlsm;.',
    '--name=RadiatorSelector',
    'obfuscated/__main__.py'  # Точка входа
])

print("✅ Готово! EXE-файл находится в папке 'dist'")