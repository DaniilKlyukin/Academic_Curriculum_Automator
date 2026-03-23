import os
import re
import logging
from pathlib import Path
from typing import List, Optional
from concurrent.futures import ThreadPoolExecutor

logger = logging.getLogger(__name__)


class FilenameCleaner:
    """Универсальный класс для очистки имен файлов от мусорных символов."""

    def __init__(self, root_dir: str, max_workers: int = 4):
        self.root_dir = Path(root_dir)
        self.max_workers = max_workers

    def _sanitize(self, filename: str) -> str:
        name, ext = os.path.splitext(filename)

        # 1. Удаляем +, - и пробелы в начале
        new_name = re.sub(r'^[\+\-\s]+', '', name)

        # 2. Удаляем все плюсы в любой части имени (из remove_extra_symbols)
        new_name = new_name.replace('+', '')

        # 3. Заменяем точки на пробелы (кроме расширения)
        new_name = new_name.replace('.', ' ')

        # Убираем лишние пробелы, если они возникли
        new_name = re.sub(r'\s+', ' ', new_name).strip()

        return f"{new_name}{ext}"

    def _process_file(self, file_path: Path) -> Optional[bool]:
        try:
            new_name = self._sanitize(file_path.name)
            if new_name == file_path.name:
                return None

            new_path = file_path.with_name(new_name)

            # Обработка дублей
            counter = 1
            while new_path.exists():
                new_path = file_path.with_name(f"{Path(new_name).stem}_{counter}{new_path.suffix}")
                counter += 1

            file_path.rename(new_path)
            logger.info(f"Переименовано: {file_path.name} -> {new_path.name}")
            return True
        except Exception as e:
            logger.error(f"Ошибка при переименовании {file_path}: {e}")
            return False

    def run(self):
        files = [Path(root) / f for root, _, fs in os.walk(self.root_dir) for f in fs]
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            executor.map(self._process_file, files)