import sys
import logging

from services.prepare_pipeline import main as run_prepare_pipeline
from services.annotation_extractor import main as run_annotation_extractor
from services.approval_processor import main as run_approval_update
from services.file_cleaner import main as run_cleanup_files
from services.pdf_generator import main as run_convert_to_pdf
from services.doc_converter import main as run_doc_converter
from services.filename_cleaner import main as run_filename_cleanup
from services.media_cleaner import main as run_media_cleanup
from services.scan_insertion import main as run_scan_insertion
from services.scan_renamer import main as run_scan_renamer
from services.signature_processor import main as run_signature_update
from services.structure_exporter import main as run_structure_exporter
from services.generate_assessment.generate_assessment_materials import main as generate_assessment_materials

logging.basicConfig(
    filename='app_errors.log',
    filemode='w',
    level=logging.ERROR,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding='utf-8'
)


def execute_service(service_func) -> None:
    """Безопасный вызов функции сервиса."""
    try:
        service_func()
    except KeyboardInterrupt:
        print("\n\n[Процесс прерван пользователем]")
    except Exception as err:
        print(f"\n[Критическая ошибка выполнения]: {err}")
        logging.error(f"Ошибка при вызове {service_func.__name__}: {err}", exc_info=True)


def main() -> None:
    menu_items = {
        "1": ("Комплексная подготовка (Пайплайн)", run_prepare_pipeline),
        "2": ("Извлечение аннотаций (Страница 3)", run_annotation_extractor),
        "3": ("Обновление учебных лет в Листах согласования", run_approval_update),
        "4": ("Удаление PDF и изображений (JPG, PNG)", run_cleanup_files),
        "5": ("Рекурсивная конвертация DOCX/PPTX в PDF", run_convert_to_pdf),
        "6": ("Конвертация .doc в .docx", run_doc_converter),
        "7": ("Массовая очистка имен файлов", run_filename_cleanup),
        "8": ("Очистка DOCX от тяжелых медиа-объектов", run_media_cleanup),
        "9": ("Автоматическая вставка сканов", run_scan_insertion),
        "10": ("Умное переименование сканов", run_scan_renamer),
        "11": ("Замена ФИО и должности в зонах подписей", run_signature_update),
        "12": ("Генератор структуры папок для ИИ", run_structure_exporter),
        "13": ("Создание оценочных материалов на основе плана", generate_assessment_materials),
    }

    while True:
        print("\n" + "="*60)
        print("=== ACADEMIC CURRICULUM AUTOMATOR ===")
        print("="*60)
        for key, (name, _) in menu_items.items():
            print(f"{key:>2}. {name}")
        print("-" * 60)
        print(" 0. Выход")
        print("="*60)

        choice = input("Выберите номер операции: ").strip()

        if choice == "0":
            print("\nЗавершение работы программы.")
            break
        elif choice in menu_items:
            execute_service(menu_items[choice][1])
        else:
            print("\nОшибка: Некорректный выбор. Пожалуйста, введите число от 0 до 14.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрограмма принудительно завершена.")
        sys.exit(0)