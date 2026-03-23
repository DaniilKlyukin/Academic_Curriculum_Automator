import logging
import os
from utils.scan_finder import ScanFinder
from services.scan_insertion_service import ScanInsertionManager

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

def main():
    print("=== Вставка сканов подписей в документы ===")
    words_dir = input("Путь к папке с РП (.docx): ").strip().strip('"')
    scans_dir = input("Путь к папке со сканами (.jpg): ").strip().strip('"')

    if not os.path.isdir(words_dir) or not os.path.isdir(scans_dir):
        print("Ошибка: Пути не найдены.")
        return

    scan_finder = ScanFinder(scans_dir)
    manager = ScanInsertionManager(scan_finder)

    docx_files = [
        os.path.join(words_dir, f) for f in os.listdir(words_dir)
        if f.lower().endswith('.docx') and not f.startswith('~$')
    ]

    if not docx_files:
        print("Файлы .docx не найдены.")
        return

    print(f"Начинаю обработку {len(docx_files)} файлов...")
    manager.process_documents(docx_files)
    print("\n[Готово] Сканы вставлены.")

if __name__ == "__main__":
    main()