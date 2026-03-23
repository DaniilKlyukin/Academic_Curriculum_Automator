import os
from os import listdir
from os.path import isfile, join
from pathlib import Path
from fpdf import FPDF

path = r'Z:\home\!ПМиИТ\!ПМиИТ\!УЧЕБНЫЙ ПОРТАЛ\!2024 год набора\Магистры 01.04.04\Рабочие программы\Сканы'

files = [f for f in listdir(path) if isfile(join(path, f)) and (f.endswith('.jpg') or f.endswith('.jpeg')) and '~' not in f]

pdfs_path = join(path, 'PDF')
if not os.path.exists(pdfs_path):
    os.makedirs(pdfs_path)

triples = [files[i:i+3] for i in range(0, len(files), 3)]

for triple in triples:

    pdf_name = Path(triple[0]).stem[:-1] # убираем расширение jpg и цифру после предмета

    pdf = FPDF()
    # imagelist is the list with all image filenames
    for image_name in triple:
        pdf.add_page()
        image_path = join(path, image_name)
        pdf.image(image_path, 0, 0, 210, 297)

    pdf_path = join(pdfs_path, f"{pdf_name}.pdf")
    pdf.output(pdf_path, "F")