import os
import logging
from services.approval_processor import process_docx, generate_years

logging.basicConfig(
    filename='app_errors.log',
    filemode='w',
    level=logging.ERROR,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding='utf-8'
)


def main():
    print("=== Обновление учебных лет в Листах согласования ===")
    folder_path = input("Введите путь до папки с .docx: ").strip().strip('"')

    if not os.path.isdir(folder_path):
        print("Ошибка: Путь не найден.")
        return

    try:
        start_y = int(input("Введите начальный год: "))
        end_y = int(input("Введите конечный год: "))
    except ValueError:
        print("Ошибка: Введите числа.")
        return

    years_list = generate_years(start_y, end_y)

    files = [f for f in os.listdir(folder_path) if f.lower().endswith('.docx') and not f.startswith('~$')]
    total = len(files)

    print(f"\n{'№':<9} | {'Статус':<8} | {'Файл'}")
    print("-" * 80)

    for i, filename in enumerate(files, 1):
        file_path = os.path.join(folder_path, filename)
        success, message = process_docx(file_path, years_list)

        status = "OK" if success else "SKIP"

        print(f"[{i:03}/{total:03}] | {status:<8} | {os.path.basename(file_path)[:60]}")


if __name__ == "__main__":
    main()