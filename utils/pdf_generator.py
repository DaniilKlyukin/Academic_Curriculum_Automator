import win32com.client
import os
from pathlib import Path

class PDFGenerator:
    """Конвертация .docx в .pdf через установленный MS Word."""

    @staticmethod
    def convert(docx_path: str, pdf_path: str = None):
        if not pdf_path:
            pdf_path = str(Path(docx_path).with_suffix('.pdf'))

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(os.path.abspath(docx_path))
            # wdFormatPDF = 17
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
            doc.Close()
            return True
        except Exception as e:
            print(f"Ошибка PDF: {e}")
            return False
        finally:
            word.Quit()