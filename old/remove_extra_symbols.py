import os
from os import listdir
from os.path import isfile, join
import re

path = r'Z:\home\!ПМиИТ\!ПМиИТ\!УЧЕБНЫЙ ПОРТАЛ\!2024 год набора\Бакалавры 01.03.02\2 профиль ПАДиИИ\Рабочие программы\Аннотации'

files = [f for f in listdir(path) if isfile(join(path, f))]

for f in files:
    file_path = join(path, f)
    new_file_path = re.sub(r'\+', r'', file_path)
    os.rename(file_path, new_file_path)