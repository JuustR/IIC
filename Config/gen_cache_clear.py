from win32com.client import gencache
import shutil
import os

# Удалить сгенерированные файлы
gen_py_path = os.path.join(os.environ["LOCALAPPDATA"], "Temp", "gen_py")
if os.path.exists(gen_py_path):
    shutil.rmtree(gen_py_path)

# Пересоздать кэш
gencache.Rebuild()