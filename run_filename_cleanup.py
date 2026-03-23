import logging
from utils.filename_cleaner import FilenameCleaner

logging.basicConfig(level=logging.INFO, format="%(message)s")


def main():
    print("=== Массовая очистка имен файлов ===")
    path = input("Введите путь к папке: ").strip().strip('"')

    cleaner = FilenameCleaner(path)
    print("Начинаю переименование...")
    cleaner.run()
    print("\n[Готово] Имена файлов приведены в порядок.")


if __name__ == "__main__":
    main()