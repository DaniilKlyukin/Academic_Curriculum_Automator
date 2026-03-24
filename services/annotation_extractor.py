import os
from os import listdir, makedirs, remove
from os.path import isfile, join, exists
from pathlib import Path
from contextlib import contextmanager
import comtypes.client
from pypdf import PdfReader, PdfWriter


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

class AnnotationExtractor:
    WD_FORMAT_PDF = 17

    def __init__(self, annotation_page: int = 3, extensions=None):
        self.extensions = extensions or [".doc", ".docx"]
        self.annotation_page = annotation_page

    def _process_single(self, word_app, input_file: str, output_folder: str):
        pdf_path = join(output_folder, f'{Path(input_file).stem}.pdf')
        try:
            doc = word_app.Documents.Open(input_file)
            doc.SaveAs(pdf_path, FileFormat=self.WD_FORMAT_PDF)
            doc.Close()
            self._extract_page(pdf_path)
        except Exception:
            if exists(pdf_path):
                remove(pdf_path)
            raise

    def _extract_page(self, pdf_path: str):
        with open(pdf_path, "rb") as f:
            pdf_reader = PdfReader(f)
            if self.annotation_page > len(pdf_reader.pages):
                raise ValueError("Недостаточно страниц")
            pdf_writer = PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[self.annotation_page - 1])
        with open(pdf_path, "wb") as out:
            pdf_writer.write(out)

    def extract_annotations(self, input_folder: str, output_folder: str):
        makedirs(output_folder, exist_ok=True)
        files = []
        for root, _, filenames in os.walk(input_folder):
            for f in filenames:
                if any(f.lower().endswith(e) for e in self.extensions) and '~' not in f:
                    files.append(join(root, f))

        with word_application() as word_app:
            for file_path in files:
                self._process_single(word_app, file_path, output_folder)