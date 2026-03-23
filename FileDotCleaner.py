import os
from pathlib import Path
from typing import List, Optional
import logging
from concurrent.futures import ThreadPoolExecutor


class FileDotCleaner:
    """
    Класс для рекурсивного убирания точек из файлов в директории,
    удаляя точки в названиях (кроме расширения).

    Особенности:
    - Рекурсивная обработка поддиректорий
    - Потокобезопасные операции
    - Подробное логирование
    - Проверка существования файлов
    - Обработка коллизий имен
    """

    def __init__(self, root_dir: str, max_workers: int = 4):
        self.root_dir = Path(root_dir)
        self.max_workers = max_workers
        self.logger = self._setup_logger()

    @staticmethod
    def _setup_logger() -> logging.Logger:
        """Настройка логгера с форматированием."""
        logger = logging.getLogger("DocRenamer")
        logger.setLevel(logging.INFO)

        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            "%(asctime)s - %(levelname)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)

        return logger

    def _sanitize_filename(self, filename: str) -> str:
        """
        Удаляет точки из имени файла, сохраняя расширение.
        Пример:
            "Resume.v2.final.pdf" -> "Resumev2final.pdf"
        """
        name, ext = os.path.splitext(filename)
        # Удаляем все точки, кроме последней (которая перед расширением)
        sanitized = name.replace('.', ' ')
        return f"{sanitized}{ext}"

    def _process_file(self, file_path: Path) -> Optional[bool]:
        """Обрабатывает один файл, возвращает успешность операции."""
        try:
            new_name = self._sanitize_filename(file_path.name)
            if new_name == file_path.name:
                return None  # Имя не изменилось

            new_path = file_path.with_name(new_name)

            # Обработка коллизий имен
            counter = 1
            while new_path.exists():
                stem = new_path.stem
                new_path = new_path.with_name(f"{stem}_{counter}{new_path.suffix}")
                counter += 1

            file_path.rename(new_path)
            self.logger.info(f"Renamed: {file_path.name} -> {new_path.name}")
            return True

        except Exception as e:
            self.logger.error(f"Error processing {file_path}: {str(e)}")
            return False

    def find_doc_files(self) -> List[Path]:
        doc_files = []
        for root, _, files in os.walk(self.root_dir):
            for file in files:
                doc_files.append(Path(root) / file)
        return doc_files

    def run(self) -> None:
        """Основной метод выполнения переименования."""
        if not self.root_dir.exists():
            raise FileNotFoundError(f"Directory not found: {self.root_dir}")

        self.logger.info(f"Starting processing in: {self.root_dir}")
        doc_files = self.find_doc_files()

        if not doc_files:
            self.logger.warning("No files found")
            return

        self.logger.info(f"Found {len(doc_files)} files to process")

        # Многопоточная обработка
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            results = list(executor.map(self._process_file, doc_files))

        success_count = sum(1 for r in results if r is True)
        self.logger.info(
            f"Processing complete. Success: {success_count}, "
            f"Skipped: {results.count(None)}, Failed: {results.count(False)}"
        )


def main():
    """Интерфейс командной строки."""

    input_folder = input("Введите путь до папки:\n")

    renamer = FileDotCleaner(input_folder, 1)
    try:
        renamer.run()
    except Exception as e:
        renamer.logger.error(f"Fatal error: {str(e)}", exc_info=True)
        exit(1)


if __name__ == "__main__":
    main()
