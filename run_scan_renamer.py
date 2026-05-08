import os
import logging
import sys
from utils.scan_renamer import ScanRenamer


def check_dependencies():
    """Проверка наличия Tesseract перед стартом."""
    import shutil
    from pytesseract import pytesseract
    if not shutil.which("tesseract") and not os.path.exists(pytesseract.tesseract_cmd):
        print("\n[!] ОШИБКА: Tesseract OCR не найден.")
        print("Скачайте и установите: https://github.com/UB-Mannheim/tesseract/wiki")
        return False
    return True


def main():
    if not check_dependencies():
        input("\nНажмите Enter для выхода...")
        return

    print("=== УМНОЕ ПЕРЕИМЕНОВАНИЕ (OCR + ПРЕДОБРАБОТКА) ===")
    path = input("Путь к папке со сканами: ").strip().strip('"')

    if not os.path.isdir(path):
        return

    logic = ScanRenamer()
    files = [f for f in os.listdir(path) if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
    total = len(files)

    print(f"\nАнализ {total} файлов...\n")
    print(f"{'№':<9} | {'Статус':<8} | {'Аббревиатура':<15} | {'Результат OCR (фрагмент)'}")
    print("-" * 100)

    current_discipline = "Unknown"

    for i, filename in enumerate(files, 1):
        full_path = os.path.join(path, filename)
        text = logic.get_text(full_path)
        page_type = logic.identify_page_type(text)

        status = "SKIP"
        abbr = "---"
        ocr_preview = text.replace('\n', ' ')[:40].strip()

        if page_type:
            if page_type == 1:
                extracted = logic.extract_discipline(text)
                if extracted: current_discipline = extracted

            abbr = logic.make_abbreviation(current_discipline)
            new_name = f"{abbr}{page_type}.jpg"
            new_path = os.path.join(path, new_name)

            # Если такой файл уже есть, добавляем индекс
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

        print(f"[{i:03}/{total:03}] | {status:<8} | {abbr:<15} | {ocr_preview}...")

    print("\n[Готово]")
    input("Нажмите Enter...")


if __name__ == "__main__":
    main()