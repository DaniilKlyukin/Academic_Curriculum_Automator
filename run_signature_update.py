import logging
import os
from services.signature_processor import process_docx_signatures

logging.basicConfig(
    filename='app_errors.log',
    filemode='w',
    level=logging.ERROR,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding='utf-8'
)

def main():
    print("=== Утилита замены ФИО и Должности в зонах подписей — Рекурсивный поиск ===")

    dir_path = input("Введите путь к папке с файлами .docx: ").strip().strip('"')
    if not os.path.isdir(dir_path):
        print("Ошибка: Путь не найден.")
        return

    old_fio = input("Введите ФИО, которое нужно заменить: ").strip()
    new_fio = input("Введите новое ФИО: ").strip()

    old_title = input("Введите старую должность (или оставьте пустым): ").strip()
    new_title = input("Введите новую должность (или оставьте пустым): ").strip()

    if not old_fio or not new_fio:
        print("Ошибка: ФИО не могут быть пустыми.")
        return

    docx_files = [os.path.join(root, f) for root, _, fs in os.walk(dir_path)
                  for f in fs if f.lower().endswith('.docx') and not f.startswith('~$')]

    total = len(docx_files)
    print(f"\n{'№':<9} | {'Статус':<8} | {'Файл'}")
    print("-" * 80)

    for i, path in enumerate(docx_files, 1):
        rel = os.path.relpath(path, dir_path)
        success, msg = process_docx_signatures(path, old_fio, new_fio, old_title, new_title)

        status = "OK" if success else "SKIP"

        print(f"[{i:03}/{total:03}] | {status:<8} | {os.path.basename(path)[:60]}")


if __name__ == "__main__":
    main()