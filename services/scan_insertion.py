import logging
from pathlib import Path
from typing import List, Union, Tuple
from core.docx_editor import DocxEditor
from services.scan_finder import ScanFinder

logger = logging.getLogger(__name__)


class ScanInsertionManager:
    """
    Управляющий класс для автоматизации вставки сканов в документы.
    Связывает логику поиска файлов (ScanFinder) и редактирования документов (DocxEditor).
    """

    def __init__(self, scan_finder: ScanFinder) -> None:
        self.scan_finder: ScanFinder = scan_finder

    def process_documents(self, doc_paths: List[Union[str, Path]]) -> None:
        """
        Обрабатывает список документов: ищет подходящие сканы по имени файла и
        интегрирует их в структуру DOCX на заданные позиции.
        """
        total: int = len(doc_paths)
        print(f"\n{'№':<7} | {'Статус':<8} | {'Скор':<5} | {'Файл'}")
        print("-" * 80)

        for i, path in enumerate(doc_paths, 1):
            file_path: Path = Path(path)
            file_name: str = file_path.name

            scans_data: Tuple[List[str], float] = self.scan_finder.find_scans_for_program(file_name)
            scans: List[str] = scans_data[0]
            score: float = scans_data[1]

            status: str = "SKIP"
            if scans and len(scans) == 3:
                try:
                    with DocxEditor(str(file_path)) as editor:
                        editor.add_scan_to_page(1, scans[0])
                        editor.add_scan_to_page(2, scans[1])

                        if not editor.insert_image_after_text("Лист согласования", scans[2]):
                            editor.add_scan_to_page(30, scans[2], floating=False)
                    status = "OK"
                except Exception as e:
                    status = "ERR"
                    logger.error(f"Ошибка при обработке {file_name}: {e}")

            print(f"[{i:0>3}/{total:0>3}] | {status:<8} | {score:<5} | {file_name[:50]}")