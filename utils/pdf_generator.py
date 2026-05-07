import win32com.client
import os
from pathlib import Path
import logging

# Настройка логирования, чтобы видеть ошибки в консоли
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)


class PDFGenerator:
    def __init__(self):
        self.word = None
        self.ppt = None
        self.success_count = 0
        self.fail_count = 0

    def _get_word(self):
        if self.word is None:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
        return self.word

    def _get_ppt(self):
        if self.ppt is None:
            self.ppt = win32com.client.Dispatch("PowerPoint.Application")
        return self.ppt

    def convert_docx(self, docx_path):
        docx_path = os.path.abspath(docx_path)
        pdf_path = str(Path(docx_path).with_suffix('.pdf'))
        try:
            word = self._get_word()
            doc = word.Documents.Open(docx_path)
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
            doc.Close()
            return True
        except Exception as e:
            logger.error(f"Ошибка конвертации Word {docx_path}: {e}")
            return False

    def convert_pptx(self, pptx_path):
        pptx_path = os.path.abspath(pptx_path)
        pdf_path = str(Path(pptx_path).with_suffix('.pdf'))
        try:
            ppt = self._get_ppt()
            pres = ppt.Presentations.Open(pptx_path, WithWindow=False)
            pres.SaveAs(pdf_path, FileFormat=32)  # 32 = ppSaveAsPDF
            pres.Close()
            return True
        except Exception as e:
            logger.error(f"Ошибка конвертации PowerPoint {pptx_path}: {e}")
            return False

    def process_folder(self, folder_path):
        """Рекурсивный обход папок и конвертация всех найденных файлов."""
        self.success_count = 0
        self.fail_count = 0

        if not os.path.isdir(folder_path):
            logger.error(f"Путь не найден: {folder_path}")
            return

        print(f"--- Сканирование папки: {folder_path} ---")

        for root, _, filenames in os.walk(folder_path):
            for filename in filenames:
                if filename.startswith('~$'):
                    continue

                full_path = os.path.join(root, filename)
                ext = filename.lower()
                result = False

                if ext.endswith('.docx') or ext.endswith('.doc'):
                    print(f"[PROCESS] Word: {filename}")
                    result = self.convert_docx(full_path)

                elif ext.endswith('.pptx') or ext.endswith('.ppt'):
                    print(f"[PROCESS] PPT:  {filename}")
                    result = self.convert_pptx(full_path)

                else:
                    continue

                if result:
                    self.success_count += 1
                else:
                    self.fail_count += 1

    def quit(self):
        """Корректное закрытие приложений."""
        if self.word:
            try:
                self.word.Quit()
            except:
                pass
        if self.ppt:
            try:
                self.ppt.Quit()
            except:
                pass
        print(f"\n[Завершено] Успешно: {self.success_count}, Ошибок: {self.fail_count}")