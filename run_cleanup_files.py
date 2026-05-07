import os
from utils.file_cleaner import FileCleaner

logging.basicConfig(
    filename='app_errors.log',
    filemode='w',
    level=logging.ERROR,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding='utf-8'
)

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