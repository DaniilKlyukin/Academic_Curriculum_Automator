import logging
import os
from services.annotation_extractor import AnnotationExtractor

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)


def main():
    print("=== Извлечение аннотаций (3-я страница) в PDF ===")
    input_dir = input("Путь к папке с РП (docx): ").strip().strip('"')
    output_dir = input("Путь к папке для сохранения аннотаций: ").strip().strip('"')

    if not os.path.exists(input_dir):
        print("Ошибка: Путь не найден.")
        return

    # По умолчанию извлекаем 3-ю страницу
    extractor = AnnotationExtractor(annotation_page=3)

    try:
        extractor.extract_annotations(input_dir, output_dir)
        print(f"\n[Готово] Аннотации сохранены в: {output_dir}")
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}")


if __name__ == "__main__":
    main()