import os
import re
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from typing import List


class FilenameCleaner:
    """
    Класс для массовой очистки имен файлов: удаление спецсимволов, нормализация пробелов
    и исправление префиксов в именах файлов в указанной директории.
    """

    def __init__(self, root_dir: str, max_workers: int = 4) -> None:
        self.root_dir: Path = Path(root_dir)
        self.max_workers: int = max_workers

    def _sanitize(self, filename: str) -> str:
        """
        Очищает строку имени файла от лишних символов (+, .), ведущих знаков и лишних пробелов.
        """
        name, ext = os.path.splitext(filename)
        new_name: str = re.sub(r'^[\+\-\s]+', '', name).replace('+', '').replace('.', ' ')
        return f"{re.sub(r'\s+', ' ', new_name).strip()}{ext}"

    def _process_file(self, file_path: Path) -> None:
        """
        Переименовывает файл после санитарной обработки имени, обрабатывая коллизии имен.
        """
        try:
            new_name: str = self._sanitize(file_path.name)
            if new_name == file_path.name:
                return

            new_path: Path = file_path.with_name(new_name)

            counter: int = 1
            while new_path.exists():
                stem: str = Path(new_name).stem
                new_path = file_path.with_name(f"{stem}_{counter}{new_path.suffix}")
                counter += 1

            file_path.rename(new_path)
        except Exception:
            pass

    def run(self) -> None:
        """
        Запускает рекурсивный обход директории и многопоточную очистку имен файлов.
        """
        files: List[Path] = [
            Path(root) / f
            for root, _, fs in os.walk(self.root_dir)
            for f in fs
        ]

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            executor.map(self._process_file, files)