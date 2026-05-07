import os
import logging
from utils.scan_finder import ScanFinder
from services.scan_insertion_service import ScanInsertionManager

logging.basicConfig(
    filename='app_errors.log',
    filemode='w',
    level=logging.ERROR,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding='utf-8'
)

def main():
    print("=== АВТОМАТИЧЕСКАЯ ВСТАВКА СКАНОВ ===")

    doc_dir = input("Папка с Word (.docx): ").strip().strip('"')
    img_dir = input("Папка со сканами (.jpg): ").strip().strip('"')

    if not os.path.isdir(doc_dir) or not os.path.isdir(img_dir):
        print("Ошибка: Путь не существует.")
        return

    finder = ScanFinder(img_dir)
    manager = ScanInsertionManager(finder)

    files = [
        os.path.join(doc_dir, f) for f in os.listdir(doc_dir)
        if f.lower().endswith('.docx') and not f.startswith('~$')
    ]

    if not files:
        print("В папке нет файлов .docx")
        return

    print(f"\nНайдено документов: {len(files)}")
    print("Начинаю работу...")

    manager.process_documents(files)

    print("\n" + "=" * 30)
    print("ГОТОВО!")
    print("Если были ошибки, проверьте файл insertion_log.log")
    input("Нажмите Enter для выхода...")


if __name__ == "__main__":
    main()