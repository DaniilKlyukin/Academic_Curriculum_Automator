import os
import logging
from pathlib import Path
from services.generate_assessment.generate_assessment_materials_1 import CompetencyReportGenerator as TableGenerator
from services.generate_assessment.generate_assessment_materials_2 import CompetencyReportGenerator as Section2Generator
from services.generate_assessment.generate_assessment_materials_3 import CompetencyReportGenerator as Section3Generator
from services.generate_assessment.generate_assessment_materials_4 import CompetencyReportGenerator as Section4Generator

logger = logging.getLogger(__name__)


def main():
    print("=== Мастер-панель комплексной генерации оценочных материалов ===")

    user_excel_path: str = input("Шаг 1. Введите путь к исходному файлу Excel (например, plan.xlsx): ").strip()
    user_folder_path: str = input("Шаг 2. Введите путь к папке для сохранения документа Word (например, C:\\Reports): ").strip()

    if not user_excel_path:
        user_excel_path = "plan.xlsx"
        print(f"Используется путь по умолчанию для Excel: {user_excel_path}")

    if not user_folder_path:
        user_folder_path = "."
        print(f"Используется текущая рабочая папка по умолчанию: {Path(user_folder_path).absolute()}")

    # Шаг 3. Настройка режима интеграции тестов (ИИ)
    print("\nШаг 3. Настройка режима интеграции тестов (ИИ):")
    print("  1 — Полный ИИ (все тесты генерируются ИИ, файлы РП не требуются)")
    print("  2 — Смешанный (тесты берутся из РП; при отсутствии — генерируются ИИ)")
    print("  3 — Без ИИ (оригинальное поведение: тесты только из РП, иначе заглушки)")
    ai_mode_input = input("Выберите режим [По умолчанию: 3]: ").strip()
    ai_mode = int(ai_mode_input) if ai_mode_input in ["1", "2", "3"] else 3

    # Запрашиваем путь к РП только если выбран смешанный (2) или оригинальный (3) режим
    user_rp_folder = ""
    if ai_mode in [2, 3]:
        user_rp_folder = input("Шаг 4. Введите путь к папке с файлами Рабочих Программ (РП): ").strip()
        if not user_rp_folder:
            user_rp_folder = "."
            print(f"Используется текущая рабочая папка по умолчанию для РП: {Path(user_rp_folder).absolute()}")

    # Настройка API-ключа и лимитов только если выбран режим 1 или 2
    api_key = ""
    rpm_limit = 15
    if ai_mode in [1, 2]:
        api_key = os.environ.get("GEMINI_API_KEY", "").strip()
        if not api_key:
            api_key = input("Введите ваш API-ключ Gemini (или оставьте пустым при наличии системной переменной): ").strip()
            if not api_key:
                api_key = os.environ.get("GEMINI_API_KEY", "").strip()

        rpm_input = input("Введите лимит запросов в минуту (RPM) [По умолчанию: 15]: ").strip()
        if rpm_input.isdigit():
            rpm_limit = int(rpm_input)

    excel_path = Path(user_excel_path)
    folder_path = Path(user_folder_path)

    # Путь к итоговому файлу, который будет создан на Шаге 1 и последовательно наполнен на Шагах 2 и 3
    final_docx_path = folder_path / "Оценочные материалы образовательной программы.docx"

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
    print("ЗАПУСК ШАГА 4: Поиск и интеграция реальных тестов из файлов РП / ИИ")
    print("=" * 70)

    section4_generator = Section4Generator(
        word_path=str(final_docx_path.absolute()),
        rp_folder_path=user_rp_folder,
        ai_mode=ai_mode,
        api_key=api_key,
        rpm_limit=rpm_limit
    )
    section4_generator.generate()

    print("\n" + "=" * 70)
    print("КОМПЛЕКСНАЯ ОБРАБОТКА ПОЛНОСТЬЮ ЗАВЕРШЕНА")
    print(f"Итоговый документ собран в: '{final_docx_path.name}'")
    print(f"Путь к файлу: {final_docx_path.parent}")
    print("=" * 70)


if __name__ == "__main__":
    main()