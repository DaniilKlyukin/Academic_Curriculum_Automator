from ScanFinder import ScanFinder
from ScanInsertRule import ScanInsertRule
from WordDocumentEditor import WordDocumentEditor

from typing import List, Optional
from pathlib import Path
from comtypes.client import CreateObject
import logging

logger = logging.getLogger(__name__)


class ScanInsertionManager:
    """Управляет процессом вставки сканов в документы"""

    def __init__(self, scan_finder: ScanFinder, insertion_rules: List[ScanInsertRule]):
        self.scan_finder = scan_finder
        self.insertion_rules = insertion_rules

    def process_documents(self, word_files: List[str]):
        """Обрабатывает все документы"""
        word_app = CreateObject("Word.Application")
        word_app.Visible = False

        try:
            for word_file in word_files:
                self._process_single_document(word_app, word_file)
        finally:
            word_app.Quit()

    def _process_single_document(self, word_app, word_path: str):
        """Обрабатывает один документ"""
        try:
            scans = self.scan_finder.find_scans_for_program(Path(word_path).name)
            if not scans:
                logger.warning(f"Не найдены сканы для документа {word_path}")
                return

            with WordDocumentEditor(word_app, word_path) as editor:
                for rule in self.insertion_rules:
                    try:
                        page_num = self._resolve_page_number(editor, rule)
                        if page_num:
                            editor.insert_image(scans[rule.scan_index], page_num)
                    except Exception as e:
                        logger.error(f"Ошибка при обработке правила {rule} для {word_path}: {str(e)}")

        except Exception as e:
            logger.error(f"Ошибка при обработке документа {word_path}: {str(e)}")

    def _resolve_page_number(self, editor: WordDocumentEditor, rule: ScanInsertRule) -> Optional[int]:
        """Определяет номер страницы для вставки"""
        return rule.page_resolver(editor)