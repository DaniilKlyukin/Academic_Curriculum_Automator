import logging
import os
import sys
from services.annotation_extractor import AnnotationExtractor

logging.basicConfig(
    filename='app_errors.log',
    filemode='w',
    level=logging.ERROR,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding='utf-8'
)

def main():
    print("\n" + "="*50)
    print("=== ИЗВЛЕЧЕНИЕ АННОТАЦИЙ (СТРАНИЦА 3) ===")
    print("="*50)

    input_dir = input("Путь к папке с РП (.docx): ").strip().strip('"')
    output_dir = input("Куда сохранять PDF-аннотации: ").strip().strip('"')

    if not os.path.isdir(input_dir):
        print(f"Ошибка: Путь не найден: {input_dir}")
        return

    # Извлекаем 3-ю страницу
    extractor = AnnotationExtractor(annotation_page=3)

    try:
        extractor.extract_annotations(input_dir, output_dir)
        print("\n" + "="*50)
        print(f"ГОТОВО! Результаты в: {output_dir}")
        if os.path.getsize('annotation_errors.log') > 0:
            print("В процессе были ошибки. См. annotation_errors.log")
    except Exception as e:
        print(f"\nКритическая ошибка: {e}")

    input("\nНажмите Enter, чтобы выйти...")

if __name__ == "__main__":
    main()