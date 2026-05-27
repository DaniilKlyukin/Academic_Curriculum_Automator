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


def main():
    print("=== Удаление PDF и изображений (JPG, PNG) ===")
    path = input("Введите путь к папке: ").strip().strip('"')

    if not os.path.isdir(path):
        print("Путь не найден.")
        return

    target_extensions = ('.pdf', '.jpg', '.jpeg', '.png')
    files_to_delete = []

    for root, _, files in os.walk(path):
        for filename in files:
            if filename.lower().endswith(target_extensions):
                files_to_delete.append(os.path.join(root, filename))

    if not files_to_delete:
        print("Целевые файлы не найдены.")
        return

    print(f"Найдено файлов: {len(files_to_delete)}")
    confirm = input("Удалить все найденные файлы? (y/n): ").lower()

    if confirm != 'y':
        print("Отмена.")
        return

    cleaner = FileCleaner()
    success_count = 0

    for file_path in files_to_delete:
        if cleaner.delete(file_path):
            print(f"[OK] {file_path}")
            success_count += 1
        else:
            print(f"[FAIL] {file_path}")

    print(f"\nУспешно удалено: {success_count} из {len(files_to_delete)}")

if __name__ == "__main__":
    main()