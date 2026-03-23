import logging
from pathlib import Path
from core.docx_editor import DocxEditor
from utils.scan_finder import ScanFinder

# Настройка логгера для вывода в консоль
logger = logging.getLogger(__name__)


class ScanInsertionManager:
    """
    Класс-менеджер для управления процессом вставки сканов в документы Word.
    Использует ScanFinder для поиска нужных файлов и DocxEditor для модификации .docx.
    """

    def __init__(self, scan_finder: ScanFinder):
        self.scan_finder = scan_finder

    def process_documents(self, doc_paths: list):
        """
        Пакетная обработка списка путей к документам.
        """
        total = len(doc_paths)
        logger.info(f"[*] Начало обработки {total} файлов...")

        for i, path in enumerate(doc_paths, 1):
            try:
                self._process_single(path)
                if i % 5 == 0:
                    logger.info(f"[*] Обработано {i}/{total} документов...")
            except Exception as e:
                logger.error(f"[!] Критическая ошибка при обработке {path}: {e}")

    def _process_single(self, doc_path: str):
        """
        Логика обработки одного конкретного документа.
        """
        file_name = Path(doc_path).name
        # Ищем 3 скана через Finder (1 - Титульник, 2 - Аннотация, 3 - Лист согласования)
        scans = self.scan_finder.find_scans_for_program(file_name)

        if not scans or len(scans) < 3:
            logger.warning(f"[-] Пропуск: Не найден комплект из 3-х сканов для файла: {file_name}")
            return

        try:
            # Открываем редактор (контекстный менеджер сам сохранит файл в конце)
            with DocxEditor(doc_path) as editor:

                # --- ЭТАП 1: Титульный лист ---
                # Накладываем скан 1 поверх самой первой страницы (координаты 0,0)
                editor.add_image_at_beginning(scans[0])
                logger.debug(f"[{file_name}] Скан титульника наложен.")

                # --- ЭТАП 2: Аннотация (обычно 2 или 3 страница) ---
                # Ищем маркер 'АННОТАЦИЯ' и накладываем скан 2 поверх этой страницы
                # По умолчанию ширина 210мм и высота 297мм (А4)
                annotation_found = editor.insert_image_after_text(
                    search_text="АННОТАЦИЯ",
                    image_path=scans[1]
                )

                if not annotation_found:
                    # Если текст 'АННОТАЦИЯ' не найден, пробуем наложить на 2-ю страницу (условно 5-й параграф)
                    if len(editor.doc.paragraphs) > 5:
                        editor.add_floating_scan(editor.doc.paragraphs[5], scans[1])
                        logger.info(
                            f"[!] Маркер 'АННОТАЦИЯ' не найден в {file_name}, наложено по умолчанию на 2-ю стр.")
                    else:
                        logger.warning(f"[!] Не удалось вставить аннотацию в {file_name} (документ слишком короткий).")

                # --- ЭТАП 3: Лист согласования ---
                # Ищем маркер 'Лист согласования' и накладываем скан 3 поверх
                success_sheet = editor.insert_image_after_text(
                    search_text="Лист согласования",
                    image_path=scans[2]
                )

                if not success_sheet:
                    # Если маркер не найден, создаем новый абзац в конце и накладываем скан на последнюю страницу
                    editor.add_image_at_end(scans[2])
                    logger.info(f"[!] Маркер согласования не найден в {file_name}, наложено на последнюю страницу.")

            logger.info(f"[+] УСПЕХ: {file_name}")

        except PermissionError:
            logger.error(f"[!] ОШИБКА: Файл {file_name} открыт в Word. Закройте его и повторите.")
        except Exception as e:
            logger.error(f"[!] ОШИБКА в {file_name}: {str(e)}")