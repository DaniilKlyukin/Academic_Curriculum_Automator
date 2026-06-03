import sys
import logging

from services.prepare_pipeline import main as run_prepare_pipeline
from services.annotation_extractor import main as run_annotation_extractor
from services.approval_processor import main as run_approval_update
from services.file_cleaner import main as run_cleanup_files
from services.pdf_generator import main as run_convert_to_pdf
from services.doc_converter import main as run_doc_converter
from services.filename_standardizer import main as run_filename_standardizer
from services.media_cleaner import main as run_media_cleanup
from services.scan_insertion import main as run_scan_insertion
from services.filename_scans_standardizer import main as run_scan_standardizer
from services.signature_processor import main as run_signature_update
from services.structure_exporter import main as run_structure_exporter
from services.generate_assessment.generate_assessment_materials import main as generate_assessment_materials
from services.image_service import ImageToPDFService
from services.rp_generator.rp_personnel_extractor import main as run_rp_personnel_extractor
from services.rp_generator.rp_academic_parser import main as run_rp_academic_parser
from services.rp_generator.rp_competency_mapper import main as run_rp_competency_mapper
from services.rp_generator.rp_ai_data_generator import main as run_rp_ai_data_generator
from services.rp_generator.rp_generator import main as run_rp_generator

logging.basicConfig(
    filename='app_errors.log',
    filemode='w',
    level=logging.ERROR,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding='utf-8'
)


def run_image_to_pdf() -> None:
    """Интерфейс для объединения набора изображений в PDF."""
    print("\n=== ОБЪЕДИНЕНИЕ ИЗОБРАЖЕНИЙ В PDF ===")
    input_dir = input("Введите путь к папке с изображениями: ").strip().strip('"')
    if not input_dir:
        print("Путь не указан.")
        return
    output_dir = input(
        "Введите путь для сохранения PDF (оставьте пустым для сохранения в подпапку PDF_Output): ").strip().strip('"')
    output_path = output_dir if output_dir else None

    service = ImageToPDFService()
    try:
        service.generate_pdfs(input_dir, output_path)
        print("Обработка графических файлов завершена.")
    except Exception as err:
        print(f"Ошибка при объединении изображений: {err}")


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
        # Группа 1: Комплексные пайплайны и генерация материалов
        "1": ("Запуск комплексного пайплайна предварительной подготовки документов", run_prepare_pipeline),
        "2": ("Генерация фондов оценочных материалов (ФОМ) по учебному плану", generate_assessment_materials),

        # Группа 2: Интеграция графических материалов и сканов
        "3": ("Автоматическая вставка подготовленных сканов в документы Word", run_scan_insertion),
        "4": ("Объединение отдельных графических файлов сканов в PDF-документы", run_image_to_pdf),

        # Группа 3: Замена реквизитов и текстовых зон
        "5": ("Массовая замена ФИО и должностей в зонах подписей", run_signature_update),
        "6": ("Обновление периодов (учебных лет) в листах согласования", run_approval_update),

        # Группа 4: Переименование и стандартизация имен файлов
        "7": ("Массовое приведение имен файлов РПД к единому стандарту", run_filename_standardizer),
        "8": ("Распознавание текста и интеллектуальное переименование сканов", run_scan_standardizer),

        # Группа 5: Конвертация форматов
        "9": ("Рекурсивная пакетная конвертация файлов DOCX/PPTX в PDF", run_convert_to_pdf),
        "10": ("Пакетное конвертирование устаревших файлов .doc в .docx", run_doc_converter),

        # Группа 6: Извлечение, очистка и оптимизация
        "11": ("Извлечение страниц аннотаций (страница 3) в PDF-файлы", run_annotation_extractor),
        "12": ("Оптимизация объема документов (удаление неиспользуемых медиа-объектов)", run_media_cleanup),
        "13": ("Очистка директории от временных PDF и графических файлов", run_cleanup_files),

        # Группа 7: Служебные утилиты
        "14": ("Экспорт структуры папок проекта для передачи в ИИ", run_structure_exporter),

        "15": ("[РП Модуль] Извлечение кадров из старых РП", run_rp_personnel_extractor),
        "16": ("[РП Модуль] Парсинг учебных нагрузок и часов (План)", run_rp_academic_parser),
        "17": ("[РП Модуль] Разбор связей предметов и компетенций", run_rp_competency_mapper),
        "18": ("[РП Модуль] ИИ Генерация недостающих данных РП (Gemini)", run_rp_ai_data_generator),
        "19": ("[РП Модуль] Генерация полных рабочих программ (РПД)", run_rp_generator),
    }

    while True:
        print("\n" + "=" * 60)
        print("=== ACADEMIC CURRICULUM AUTOMATOR (ACA) ===")
        print("=" * 60)
        for key, (name, _) in menu_items.items():
            print(f"{key:>2}. {name}")
        print("-" * 60)
        print(" 0. Выход")
        print("=" * 60)

        choice = input("Выберите номер операции: ").strip()

        if choice == "0":
            print("\nЗавершение работы программы.")
            break
        elif choice in menu_items:
            execute_service(menu_items[choice][1])
        else:
            print("\nОшибка: Некорректный выбор. Пожалуйста, введите число от 0 до 19.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрограмма принудительно завершена.")
        sys.exit(0)