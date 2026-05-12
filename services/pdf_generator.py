import win32com.client
import os
import logging
from pathlib import Path
from typing import Optional, Any

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)


class PDFGenerator:
    """
    Сервис для автоматической конвертации документов Microsoft Word и презентаций PowerPoint в формат PDF.
    Использует COM-интерфейсы для взаимодействия с установленными приложениями Office.
    """

    def __init__(self) -> None:
        self.word: Optional[Any] = None
        self.ppt: Optional[Any] = None
        self.success_count: int = 0
        self.fail_count: int = 0

    def _get_word(self) -> Any:
        """
        Инициализирует и возвращает экземпляр приложения Word (Lazy initialization).
        """
        if self.word is None:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
        return self.word

    def _get_ppt(self) -> Any:
        """
        Инициализирует и возвращает экземпляр приложения PowerPoint (Lazy initialization).
        """
        if self.ppt is None:
            self.ppt = win32com.client.Dispatch("PowerPoint.Application")
        return self.ppt

    def convert_docx(self, docx_path: str) -> bool:
        """
        Конвертирует документ Word (.doc, .docx) в PDF.
        """
        abs_docx_path: str = os.path.abspath(docx_path)
        pdf_path: str = str(Path(abs_docx_path).with_suffix('.pdf'))
        try:
            word: Any = self._get_word()
            doc: Any = word.Documents.Open(abs_docx_path)
            doc.SaveAs(pdf_path, FileFormat=17)
            doc.Close()
            return True
        except Exception as e:
            logger.error(f"Ошибка конвертации Word {docx_path}: {e}")
            return False

    def convert_pptx(self, pptx_path: str) -> bool:
        """
        Конвертирует презентацию PowerPoint (.ppt, .pptx) в PDF.
        """
        abs_pptx_path: str = os.path.abspath(pptx_path)
        pdf_path: str = str(Path(abs_pptx_path).with_suffix('.pdf'))
        try:
            ppt: Any = self._get_ppt()
            pres: Any = ppt.Presentations.Open(abs_pptx_path, WithWindow=False)
            pres.SaveAs(pdf_path, FileFormat=32)
            pres.Close()
            return True
        except Exception as e:
            logger.error(f"Ошибка конвертации PowerPoint {pptx_path}: {e}")
            return False

    def process_folder(self, folder_path: str) -> None:
        """
        Рекурсивно обходит папку и конвертирует все найденные документы Word и PowerPoint в PDF.
        """
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

                full_path: str = os.path.join(root, filename)
                ext: str = filename.lower()
                result: bool = False

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

    def quit(self) -> None:
        """
        Завершает работу запущенных приложений Microsoft Office.
        """
        if self.word:
            try:
                self.word.Quit()
            except Exception:
                pass
        if self.ppt:
            try:
                self.ppt.Quit()
            except Exception:
                pass
        print(f"\n[Завершено] Успешно: {self.success_count}, Ошибок: {self.fail_count}")