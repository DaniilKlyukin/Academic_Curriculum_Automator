import os
import logging
from pathlib import Path
from typing import List, Optional
import comtypes.client
import comtypes


class DocToPdfConverter:
    """
    Улучшенный класс для конвертации .doc/.docx в PDF с правильной инициализацией COM.
    """

    WORD_FORMAT_PDF = 17

    def __init__(self, root_dir: str, output_dir: Optional[str] = None, max_workers: int = 1):
        self.root_dir = Path(root_dir)
        self.output_dir = Path(output_dir) if output_dir else self.root_dir / "pdf_output"
        self.max_workers = max_workers  # Word COM не потокобезопасен
        self.logger = self._setup_logger()
        self.word_app = None

    @staticmethod
    def _setup_logger() -> logging.Logger:
        logger = logging.getLogger("DocToPdfConverter")
        logger.setLevel(logging.INFO)
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            "%(asctime)s - %(levelname)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        return logger

    def _initialize_word(self):
        """Правильная инициализация COM и Word."""
        try:
            comtypes.CoInitialize()  # Инициализация COM
            self.word_app = comtypes.client.CreateObject("Word.Application")
            self.word_app.Visible = False
        except Exception as e:
            self.logger.error(f"Word initialization failed: {str(e)}")
            raise

    def _close_word(self):
        """Корректное закрытие Word и COM."""
        try:
            if self.word_app:
                self.word_app.Quit()
                self.word_app = None
        finally:
            comtypes.CoUninitialize()  # Важно!

    def _convert_single_file(self, doc_path: Path) -> Optional[bool]:
        """Конвертация одного файла с обработкой ошибок."""
        try:
            pdf_path = self.output_dir / f"{doc_path.stem}.pdf"
            pdf_path.parent.mkdir(parents=True, exist_ok=True)

            if pdf_path.exists():
                self.logger.info(f"Skipped (exists): {pdf_path}")
                return None

            doc = self.word_app.Documents.Open(str(doc_path))
            doc.SaveAs(str(pdf_path), FileFormat=self.WORD_FORMAT_PDF)
            doc.Close(SaveChanges=False)
            return True

        except Exception as e:
            self.logger.error(f"Failed {doc_path}: {str(e)}")
            return False
        finally:
            if 'doc' in locals():
                try:
                    doc.Close(SaveChanges=False)
                except:
                    pass

    def find_doc_files(self) -> List[Path]:
        """Поиск файлов с обработкой исключений."""
        try:
            return [
                Path(root) / file
                for root, _, files in os.walk(self.root_dir)
                for file in files
                if file.lower().endswith(('.doc', '.docx'))
            ]
        except Exception as e:
            self.logger.error(f"File search error: {str(e)}")
            return []

    def run(self):
        """Безопасный основной метод."""
        if not self.root_dir.exists():
            raise FileNotFoundError(f"Directory not found: {self.root_dir}")

        self.output_dir.mkdir(parents=True, exist_ok=True)
        doc_files = self.find_doc_files()

        if not doc_files:
            self.logger.warning("No documents found")
            return

        self._initialize_word()

        try:
            success = failed = skipped = 0
            for file in doc_files:
                result = self._convert_single_file(file)
                if result is True:
                    success += 1
                elif result is False:
                    failed += 1
                else:
                    skipped += 1

            self.logger.info(
                f"Results: {success} converted, {skipped} skipped, {failed} failed"
            )
        finally:
            self._close_word()


def main():
    """Интерфейс командной строки."""

    input_folder = input("Введите путь до папки с РП:\n")
    output_folder = input("Введите путь до папки куда будут сохранены PDF:\n")

    converter = DocToPdfConverter(input_folder, output_folder, 1)
    try:
        converter.run()
    except Exception as e:
        converter.logger.error(f"Fatal error: {str(e)}", exc_info=True)
        exit(1)


if __name__ == "__main__":
    main()
