import os
import logging
from services.doc_converter import convert_doc_to_docx
from services.media_cleaner import WordImageCleanerDocx
from services.approval_processor import process_docx, generate_years
from services.signature_processor import process_docx_signatures
from services.filename_cleaner import FilenameCleaner
from services.file_cleaner import FileCleaner

logging.basicConfig(
    filename='app_errors.log',
    filemode='w',
    level=logging.ERROR,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding='utf-8'
)

def main():
    folder_path = input("Введите путь к папке: ").strip().strip('"')
    if not os.path.isdir(folder_path):
        print("Ошибка: Путь не найден.")
        return

    try:
        start_y = int(input("Введите начальный год (для согласования): "))
        end_y = int(input("Введите конечный год (для согласования): "))
    except ValueError:
        print("Ошибка: Введите корректные числа для годов.")
        return

    old_fio = input("Введите ФИО, которое нужно заменить (подписи): ").strip()
    new_fio = input("Введите новое ФИО: ").strip()

    old_title = input("Введите старую должность для замены (напр. Заведующий кафедрой): ").strip()
    new_title = input("Введите новую должность (напр. И.о. заведующего кафедрой): ").strip()

    print("\n[1/6] Очистка папки от PDF и изображений...")
    deleted_count = FileCleaner.cleanup_folder(folder_path)
    print(f"Удалено файлов: {deleted_count}")

    print("\n[2/6] Конвертация .doc в .docx...")
    convert_doc_to_docx(folder_path)

    print("\n[3/6] Оптимизация медиа-объектов...")
    cleaner = WordImageCleanerDocx(folder_path)
    cleaner.process_all()

    docx_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.docx') and not f.startswith('~$')]

    print("\n[4/6] Листы согласования...")
    years_list = generate_years(start_y, end_y)
    for i, filename in enumerate(docx_files, 1):
        success, _ = process_docx(os.path.join(folder_path, filename), years_list)
        print(f" [{i}/{len(docx_files)}] Обработка: {filename[:40]}...", end='\r')

    print("\n\n[5/6] Зоны подписей...")
    for i, filename in enumerate(docx_files, 1):
        process_docx_signatures(os.path.join(folder_path, filename), old_fio, new_fio, old_title, new_title)
        print(f" [{i}/{len(docx_files)}] Обновление: {filename[:40]}...", end='\r')

    print("\n[6/6] Очистка имен файлов...")
    fn_cleaner = FilenameCleaner(folder_path)
    fn_cleaner.run()

    print("\n=== Все операции успешно завершены ===")


if __name__ == "__main__":
    main()