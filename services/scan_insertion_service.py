import logging
from pathlib import Path
from core.docx_editor import DocxEditor

logger = logging.getLogger(__name__)


class ScanInsertionManager:
    def __init__(self, scan_finder):
        self.scan_finder = scan_finder

    def process_documents(self, doc_paths: list):
        total = len(doc_paths)
        print(f"\n{'№':<7} | {'Статус':<8} | {'Скор':<5} | {'Файл'}")
        print("-" * 80)

        for i, path in enumerate(doc_paths, 1):
            file_name = Path(path).name
            scans, score = self.scan_finder.find_scans_for_program(file_name)

            status = "SKIP"
            if scans and len(scans) == 3:
                try:
                    with DocxEditor(path) as editor:
                        editor.add_scan_to_page(1, scans[0])
                        editor.add_scan_to_page(2, scans[1])
                        if not editor.insert_image_after_text("Лист согласования", scans[2]):
                            editor.add_scan_to_page(30, scans[2], floating=False)
                    status = "OK"
                except Exception as e:
                    status = "ERR"
                    logger.error(f"{file_name}: {e}")

            print(f"[{i:0>3}/{total:0>3}] | {status:<8} | {score:<5} | {file_name[:50]}")