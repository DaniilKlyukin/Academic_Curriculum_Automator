import os
from os import listdir
from os.path import isfile, join, exists
import comtypes.client
from pypdf import PdfReader, PdfWriter
from pathlib import Path
from typing import List, Optional
from contextlib import contextmanager
import logging


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def is_supported_extension(path: str, extensions: List[str]) -> bool:
    return any(path.lower().endswith((ext.lower()) for ext in extensions))


@contextmanager
def word_application():
    word = None
    try:
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        yield word
    finally:
        if word is not None:
            word.Quit()


class AnnotationProcessor:

    WD_FORMAT_PDF = 17

    def __init__(self, annotation_page: int = 3, extensions: Optional[List[str]] = None):

        """
        Инициализация процессора.

        :param annotation_page: Номер страницы для извлечения (начиная с 0)
        :param extensions: Список поддерживаемых расширений файлов
        """

        self.extensions = extensions or [".doc", ".docx"]
        self.annotation_page = annotation_page

    def _validate_directories(self, input_folder: str, output_folder: str) -> None:
        if not exists(input_folder):
            raise FileNotFoundError(f"Входная директория не найдена: {input_folder}")

        os.makedirs(output_folder, exist_ok=True)

    def _process_single_file(self, word_app, input_file: str, output_folder: str) -> None:
        file_name = Path(input_file).stem
        pdf_path = join(output_folder, f'{file_name}.pdf')

        try:
            doc = word_app.Documents.Open(input_file)
            doc.SaveAs(pdf_path, FileFormat=self.WD_FORMAT_PDF)
            doc.Close()

            self.

    def _extract_page_from_pdf(self, pdf_path: str) -> None:

        try:
            with open(pdf_path, "rb") as f:
                pdf_reader = PdfReader(f)

                if self.annotation_page >= len(pdf_reader.pages):
                    raise ValueError(
                        f"PDF содержит только {len(pdf_reader.pages)} страниц. "
                        f"Не могу извлечь страницу {self.annotation_page + 1}"
                    )

                pdf_writer = PdfWriter()
                pdf_writer.add_page(pdf_reader.pages[self.annotation_page])

                with open(pdf_path, "wb") as out:
                    pdf_writer.write(out)

        except Exception as e:
            logger


    def ExtractAnnotation(self, input_folder: str, output_folder: str):

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        files = [f for f in listdir(input_folder) if
                 isfile(join(input_folder, f)) and (isSupportedExtension(f, self.extensions)) and '~' not in f]

        for f in files:
            name = Path(f).stem
            pdf_path = join(output_folder, '{0}.pdf'.format(name))

            word_path = join(input_folder, f)
            word = comtypes.client.CreateObject('Word.Application')
            doc = word.Documents.Open(word_path)

            doc.SaveAs(pdf_path, FileFormat=self.__wdFormatPDF)
            doc.Close()
            word.Quit()

            pdf_reader = PdfReader(pdf_path)
            pdf_writer = PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[self.annotation_page - 1])

            with open(pdf_path, "wb") as out:
                pdf_writer.write(out)

            pdf_writer.close()

            print(f'Готово: {f}')


processor = AnnotationProcessor(2, [".doc", ".docx"])
processor.ExtractAnnotation(r'Z:\home\!ПМиИТ\!ПМиИТ\!УЧЕБНЫЙ ПОРТАЛ\!2025 год набора\Магистры 01.04.04_new\РП\РП по другим кафедрам',
                                r'C:\Users\admin\Desktop\123')