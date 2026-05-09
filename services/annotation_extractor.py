import os
import logging
from pathlib import Path
from contextlib import contextmanager
import comtypes.client

logger = logging.getLogger(__name__)


class WordConstants:
    wdExportFormatPDF = 17
    wdExportRangeFromTo = 3


@contextmanager
def word_application():
    logging.getLogger("comtypes").setLevel(logging.WARNING)
    word = None
    try:
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        yield word
    finally:
        if word is not None:
            try:
                word.Quit()
            except:
                pass


class AnnotationExtractor:
    def __init__(self, annotation_page: int = 3):
        self.annotation_page = annotation_page
        self.extensions = [".docx", ".doc"]

    def extract_annotations(self, input_folder: str, output_folder: str):
        os.makedirs(output_folder, exist_ok=True)

        files = []
        for root, _, filenames in os.walk(input_folder):
            for f in filenames:
                if any(f.lower().endswith(ext) for ext in self.extensions) and not f.startswith('~$'):
                    files.append(os.path.join(root, f))

        if not files:
            print("Файлы для обработки не найдены.")
            return

        total = len(files)
        print(f"\n{'№':<9} | {'Статус':<8} | {'Файл'}")
        print("-" * 80)

        with word_application() as word_app:
            for i, file_path in enumerate(files, 1):
                file_name = os.path.basename(file_path)
                output_pdf = os.path.join(output_folder, f"{Path(file_name).stem}.pdf")
                status = "OK"
                try:
                    doc = word_app.Documents.Open(file_path, ReadOnly=True)
                    doc.ExportAsFixedFormat(output_pdf, 17, Range=3, From=self.annotation_page, To=self.annotation_page)
                    doc.Close(False)
                except Exception as e:
                    status = "ERR"
                    logger.error(f"{file_name}: {e}")

                print(f"[{i:03}/{total:03}] | {status:<8} | {file_name[:60]}")