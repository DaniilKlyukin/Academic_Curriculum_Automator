import os
from typing import Tuple, Union


class FileCleaner:
    """
    Утилитарный класс для безопасного удаления файлов и очистки директорий.
    """

    @staticmethod
    def delete(file_path: str) -> bool:
        """
        Удаляет файл по указанному пути, если он существует.
        """
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
            return False
        except Exception:
            return False

    @staticmethod
    def cleanup_folder(
        folder_path: str,
        extensions: Union[Tuple[str, ...], str] = ('.pdf', '.jpg', '.jpeg', '.png')
    ) -> int:
        """
        Рекурсивно удаляет файлы с указанными расширениями в целевой папке.
        """
        count: int = 0
        for root, _, files in os.walk(folder_path):
            for filename in files:
                if filename.lower().endswith(extensions):
                    if FileCleaner.delete(os.path.join(root, filename)):
                        count += 1
        return count