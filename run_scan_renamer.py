import os
import logging
from services.scan_renamer import ScanRenamer

logging.basicConfig(
    filename='app_errors.log',
    filemode='w',
    level=logging.ERROR,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding='utf-8'
)


def main():
    print("=== УМНОЕ ПЕРЕИМЕНОВАНИЕ ===")
    path = input("Путь к папке со сканами (.jpg): ").strip().strip('"')

    if not os.path.isdir(path):
        print("Ошибка: Путь не найден.")
        return

    print("\nЗагрузка ИИ-моделей")
    logic = ScanRenamer()

    files = [f for f in os.listdir(path) if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
    total = len(files)

    print(f"\nАнализ {total} файлов. Пожалуйста, подождите...")
    print(f"{'№':<9} | {'Статус':<8} | {'Аббревиатура':<15} | {'OCR Текст (начало)'}")
    print("-" * 110)

    current_discipline = "Unknown"

    for i, filename in enumerate(files, 1):
        full_path = os.path.join(path, filename)
        text = logic.get_text(full_path)
        page_type = logic.identify_page_type(text)

        status = "SKIP"
        abbr = "---"
        preview = text[:45].strip() + "..."

        if page_type:
            if page_type == 1:
                extracted = logic.extract_discipline(text)
                if extracted:
                    current_discipline = extracted

            abbr = logic.make_abbreviation(current_discipline)
            new_name = f"{abbr}{page_type}.jpg"
            new_path = os.path.join(path, new_name)

            c = 1
            while os.path.exists(new_path) and os.path.abspath(full_path) != os.path.abspath(new_path):
                new_name = f"{abbr}_{c}_{page_type}.jpg"
                new_path = os.path.join(path, new_name)
                c += 1

            try:
                os.rename(full_path, new_path)
                status = f"PAGE {page_type}"
            except Exception as e:
                status = "ERR"
                logging.error(f"Error {filename}: {e}")

        print(f"[{i:03}/{total:03}] | {status:<8} | {abbr:<15} | {preview}")

    print("\n" + "=" * 40)
    print("ГОТОВО!")
    input("Нажмите Enter для выхода...")


if __name__ == "__main__":
    main()