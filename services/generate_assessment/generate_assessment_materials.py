import logging
from pathlib import Path
from services.generate_assessment.generate_assessment_materials_1 import CompetencyReportGenerator as TableGenerator
from services.generate_assessment.generate_assessment_materials_2 import CompetencyReportGenerator as Section2Generator
from services.generate_assessment.generate_assessment_materials_3 import CompetencyReportGenerator as Section3Generator

logger = logging.getLogger(__name__)


def main():
    print("=== Мастер-панель комплексной генерации оценочных материалов ===")

    user_excel_path: str = input("Шаг 1. Введите путь к исходному файлу Excel (например, plan.xlsx): ").strip()
    user_folder_path: str = input(
        "Шаг 2. Введите путь к папке для сохранения документа Word (например, C:\\Reports): ").strip()

    if not user_excel_path:
        user_excel_path = "plan.xlsx"
        print(f"Используется путь по умолчанию для Excel: {user_excel_path}")

    if not user_folder_path:
        user_folder_path = "."
        print(f"Используется текущая рабочая папка по умолчанию: {Path(user_folder_path).absolute()}")

    excel_path = Path(user_excel_path)
    folder_path = Path(user_folder_path)

    # Путь к итоговому файлу, который будет создан на Шаге 1 и последовательно наполнен на Шагах 2 и 3
    final_docx_path = folder_path / "Оценочные материалы.docx"

    print("\n" + "=" * 70)
    print("ЗАПУСК ШАГА 1: Генерация Титульного листа, Содержания и Раздела 1 (Таблица)")
    print("=" * 70)

    table_generator = TableGenerator(
        excel_path=excel_path,
        word_file_path=final_docx_path
    )
    table_generator.generate()

    # Проверка, что базовый файл успешно создан перед переходом ко второму шагу
    if not final_docx_path.exists():
        logger.error(f"Базовый файл не найден по пути: {final_docx_path}. Дальнейшая сборка остановлена.")
        print(f"\n[Ошибка] Шаг 1 не завершился созданием файла. Шаги 2 и 3 отменены.")
        return

    print("\n" + "=" * 70)
    print("ЗАПУСК ШАГА 2: Настройка стилей страниц и генерация Раздела 2 (Оценочные листы)")
    print("=" * 70)

    section2_generator = Section2Generator(
        word_path=str(final_docx_path.absolute())
    )
    section2_generator.generate()

    print("\n" + "=" * 70)
    print("ЗАПУСК ШАГА 3: Генерация Раздела 3 (Варианты диагностической работы)")
    print("=" * 70)

    section3_generator = Section3Generator(
        word_path=str(final_docx_path.absolute())
    )
    section3_generator.generate()

    print("\n" + "=" * 70)
    print("ОБРАБОТКА ПОЛНОСТЬЮ ЗАВЕРШЕНА")
    print(f"Итоговый документ собран в: '{final_docx_path.name}'")
    print(f"Путь к файлу: {final_docx_path.parent}")
    print("=" * 70)


if __name__ == "__main__":
    main()