import os
import logging
from utils.pdf_generator import PDFGenerator

logging.basicConfig(level=logging.INFO, format="%(message)s")


def main():
    print("=== Конвертация .docx в .pdf (через MS Word) ===")
    path = input("Введите путь к папке с .docx: ").strip().strip('"')

    if not os.path.isdir(path):
        print("Путь не найден.")
        return

    files = [f for f in os.listdir(path) if f.lower().endswith('.docx') and not f.startswith('~$')]

    if not files:
        print("Файлы .docx не найдены.")
        return

    print(f"Найдено файлов: {len(files)}. Начинаю конвертацию...")

    generator = PDFGenerator()
    success_count = 0

    for filename in files:
        docx_path = os.path.join(path, filename)
        # Создаем PDF в той же папке
        if generator.convert(docx_path):
            print(f"[OK] {filename} -> PDF")
            success_count += 1
        else:
            print(f"[FAIL] {filename}")

    print(f"\n[Готово] Успешно сконвертировано: {success_count} из {len(files)}")


if __name__ == "__main__":
    main()