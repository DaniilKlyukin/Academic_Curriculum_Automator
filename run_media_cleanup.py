import logging
from utils.media_cleaner import WordImageCleanerDocx

logging.basicConfig(level=logging.INFO, format="%(message)s")


def main():
    print("=== Очистка DOCX от тяжелых медиа-объектов ===")
    path = input("Введите путь к папке с .docx: ").strip().strip('"')

    cleaner = WordImageCleanerDocx(path)
    print("Анализ и очистка файлов (это может занять время)...")
    cleaner.process_all()
    print("\n[Готово] Вес файлов оптимизирован.")


if __name__ == "__main__":
    main()