import os
import logging
from pathlib import Path
from contextlib import contextmanager
from typing import List, Generator, Optional, Any, Union

import comtypes.client
from comtypes import COMError

logger = logging.getLogger(__name__)


class WordConstants:
    """Константы Microsoft Word для экспорта."""
    wdExportFormatPDF: int = 17
    wdExportRangeFromTo: int = 3
    wdDoNotSaveChanges: int = 0


@contextmanager
def word_application() -> Generator[Any, None, None]:
    """
    Контекстный менеджер для инициализации и корректного завершения приложения Word.
    """
    logging.getLogger("comtypes").setLevel(logging.WARNING)

    word: Optional[Any] = None
    try:
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        yield word
    finally:
        if word is not None:
            try:
                word.Quit(WordConstants.wdDoNotSaveChanges)
            except Exception as e:
                logger.debug(f"Ошибка при закрытии Word: {e}")


class AnnotationExtractor:
    """Класс для извлечения конкретных страниц из документов Word в PDF."""

    def __init__(self, annotation_page: int = 3) -> None:
        self.annotation_page: int = annotation_page
        self.extensions: List[str] = [".docx", ".doc"]

    def extract_annotations(self, input_folder: Union[str, Path], output_folder: Union[str, Path]) -> None:
        """
        Проходит по папке, находит документы Word и экспортирует указанную страницу в PDF.
        """
        input_path = Path(input_folder)
        output_path = Path(output_folder)
        output_path.mkdir(parents=True, exist_ok=True)

        files: List[Path] = []
        for root, _, filenames in os.walk(str(input_path)):
            for f in filenames:
                file_ext = Path(f).suffix.lower()
                if file_ext in self.extensions and not f.startswith('~$'):
                    files.append(Path(root) / f)

        if not files:
            print("Файлы для обработки не найдены.")
            return

        total: int = len(files)
        print(f"\n{'№':<9} | {'Статус':<8} | {'Файл'}")
        print("-" * 80)

        with word_application() as word_app:
            for i, file_path in enumerate(files, 1):
                file_name: str = file_path.name
                output_pdf: Path = output_path / f"{file_path.stem}.pdf"
                status: str = "OK"

                try:
                    doc = word_app.Documents.Open(str(file_path.absolute()), ReadOnly=True)

                    doc.ExportAsFixedFormat(
                        OutputFileName=str(output_pdf.absolute()),
                        ExportFormat=WordConstants.wdExportFormatPDF,
                        Range=WordConstants.wdExportRangeFromTo,
                        From=self.annotation_page,
                        To=self.annotation_page
                    )

                    doc.Close(WordConstants.wdDoNotSaveChanges)
                except COMError as ce:
                    status = "ERR"
                    logger.error(f"COM ошибка при обработке {file_name}: {ce}")
                except Exception as e:
                    status = "ERR"
                    logger.error(f"Непредвиденная ошибка при обработке {file_name}: {e}")

                print(f"[{i:03}/{total:03}] | {status:<8} | {file_name[:60]}")


def main():
    print("\n" + "="*50)
    print("=== ИЗВЛЕЧЕНИЕ АННОТАЦИЙ (СТРАНИЦА 3) ===")
    print("="*50)

    input_dir = input("Путь к папке с РП (.docx): ").strip().strip('"')
    output_dir = input("Куда сохранять PDF-аннотации: ").strip().strip('"')

    if not os.path.isdir(input_dir):
        print(f"Ошибка: Путь не найден: {input_dir}")
        return

    extractor = AnnotationExtractor(annotation_page=3)

    try:
        extractor.extract_annotations(input_dir, output_dir)
        print("\n" + "="*50)
        print(f"ГОТОВО! Результаты в: {output_dir}")
    except Exception as e:
        print(f"\nКритическая ошибка: {e}")

    input("\nНажмите Enter, чтобы выйти...")

if __name__ == "__main__":
    main()