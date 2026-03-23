from os import listdir, makedirs, remove
from os.path import isfile, join, exists
import comtypes.client
from pypdf import PdfReader, PdfWriter
from pathlib import Path
from typing import List, Optional
from contextlib import contextmanager
import logging

logger = logging.getLogger(__name__)


def is_supported_extension(path: str, extensions: List[str]) -> bool:
    return any(path.lower().endswith(ext.lower()) for ext in extensions)


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


def _validate_directories(input_folder: str, output_folder: str) -> None:
    if not exists(input_folder):
        raise FileNotFoundError(f"Входная директория не найдена: {input_folder}")

    makedirs(output_folder, exist_ok=True)


class AnnotationExtractor:
    WD_FORMAT_PDF = 17
    TEMP_WD_SYMBOL = '~'

    def __init__(self, annotation_page: int = 3, extensions: Optional[List[str]] = None):

        """
        Инициализация процессора.

        :param annotation_page: Номер страницы для извлечения (начиная с 1)
        :param extensions: Список поддерживаемых расширений файлов
        """

        self.extensions = extensions or [".doc", ".docx"]
        self.annotation_page = annotation_page

    def _process_single_file(self, word_app, input_file: str, output_folder: str) -> None:
        file_name = Path(input_file).stem
        pdf_path = join(output_folder, f'{file_name}.pdf')

        try:
            doc = word_app.Documents.Open(input_file)
            doc.SaveAs(pdf_path, FileFormat=self.WD_FORMAT_PDF)
            doc.Close()

            self._extract_page_from_pdf(pdf_path)

            logger.info(f'Успешно обработан: {input_file}')
        except Exception as e:
            logger.error(f"Ошибка при обработке файла {input_file}: {str(e)}")

            if exists(pdf_path):
                remove(pdf_path)
            raise

    def _extract_page_from_pdf(self, pdf_path: str) -> None:

        try:
            with open(pdf_path, "rb") as f:
                pdf_reader = PdfReader(f)

                if self.annotation_page > len(pdf_reader.pages):
                    raise ValueError(
                        f"PDF содержит только {len(pdf_reader.pages)} страниц. "
                        f"Не могу извлечь страницу {self.annotation_page}"
                    )

                pdf_writer = PdfWriter()
                pdf_writer.add_page(pdf_reader.pages[self.annotation_page - 1])

                with open(pdf_path, "wb") as out:
                    pdf_writer.write(out)

        except Exception as e:
            logger.error(f"Ошибка при обработке PDF {pdf_path}: {str(e)}")
            raise

    def extract_annotations(self, input_folder: str, output_folder: str) -> None:

        """
        Основной метод для извлечения аннотаций из всех файлов в директории.

        :param input_folder: Путь к входной директории с Word файлами
        :param output_folder: Путь к выходной директории для PDF
        """

        _validate_directories(input_folder, output_folder)

        files = [
            join(input_folder, f) for f in listdir(input_folder)
            if isfile(join(input_folder, f))
               and is_supported_extension(f, self.extensions)
               and self.TEMP_WD_SYMBOL not in f
        ]

        if not files:
            logger.warning(f"Не найдено файлов для обработки в {input_folder}")
            return

        logger.info(f"Начало обработки {len(files)} файлов...")

        with word_application() as word_app:
            for file_path in files:
                self._process_single_file(word_app, file_path, output_folder)

        logger.info("Обработка завершена.")


if __name__ == "__main__":
    try:
        logging.basicConfig(level=logging.INFO)

        processor = AnnotationExtractor(
            annotation_page=3,
            extensions=[".doc", ".docx"]
        )

        input_dir = input("Введите путь до папки с рабочими программами:\n")
        output_dir = input("Введите путь до папки, в которую будут записаны аннотации:\n")

        processor.extract_annotations(input_dir, output_dir)
    except Exception as e:
        logger.error(f"Критическая ошибка: {str(e)}", exc_info=True)
        exit(1)
