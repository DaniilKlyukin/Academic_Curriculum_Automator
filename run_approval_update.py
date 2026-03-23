import os
import logging
from services.approval_processor import process_docx, generate_years

logging.basicConfig(level=logging.INFO, format="%(message)s")


def main():
    print("=== Обновление учебных лет в Листах согласования ===")
    folder_path = input("Введите путь до папки с .docx: ").strip().strip('"')

    if not os.path.isdir(folder_path):
        print("Ошибка: Путь не найден.")
        return

    try:
        start_y = int(input("Введите начальный год (напр. 2024): "))
        end_y = int(input("Введите конечный год (напр. 2028): "))
    except ValueError:
        print("Ошибка: Введите числа.")
        return

    years_list = generate_years(start_y, end_y)
    processed = 0
    errors = 0

    files = [f for f in os.listdir(folder_path) if f.lower().endswith('.docx') and not f.startswith('~$')]

    for filename in files:
        file_path = os.path.join(folder_path, filename)
        success, message = process_docx(file_path, years_list)

        if success:
            print(f"[OK] {filename}")
            processed += 1
        else:
            print(f"[SKIP] {filename} ({message})")
            errors += 1

    print(f"\nИтог: Обновлено: {processed}, Пропущено: {errors}")


if __name__ == "__main__":
    main()