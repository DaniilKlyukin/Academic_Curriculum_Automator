import os
import re
import copy
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def clean_name_for_match(name: str) -> str:
    """Очищает строку от кавычек, пробелов и регистра для точного сравнения имен."""
    s = name.lower()
    s = "".join(s.split())  # Удаляем все пробелы и невидимые символы
    s = s.replace("«", "").replace("»", "").replace('"', '').replace("'", "")
    return s


def get_rp_title(doc: Document) -> str:
    """Извлекает официальное название дисциплины с титульного листа РП."""
    found_marker = False
    for p in doc.paragraphs:
        txt = " ".join(p.text.split()).strip().lower()
        if "рабочая программа" in txt or "программа дисциплины" in txt:
            found_marker = True
            continue
        if found_marker:
            title = " ".join(p.text.split()).strip()
            if title:
                return title
    # Резервный поиск на случай нестандартной верстки титула
    for p in doc.paragraphs[:12]:
        txt = " ".join(p.text.split()).strip()
        if txt and len(txt) > 5 and not txt.isupper() and "министерство" not in txt.lower():
            return txt
    return ""


def find_test_elements(doc: Document) -> tuple[list[Any], bool]:
    """Находит XML-элементы теста.

    Если оригинальная таблица ключей в файле РП не найдена —
    полностью бракует блок и возвращает пустой список.
    """
    body_elements = list(doc.element.body)
    start_idx = -1

    # Поиск начала теста
    for i, elem in enumerate(body_elements):
        if elem.tag.endswith('p'):
            text = "".join(elem.itertext()).strip().lower()
            if "примерныйвариант" in text or "вариантитоговоготеста" in text or "итоговыйтест" in text or "варианттеста" in text:
                start_idx = i
                break

    if start_idx == -1:
        for i, elem in enumerate(body_elements):
            text = "".join(elem.itertext()).strip().lower()
            if "тест" in text and len(text) < 100:
                start_idx = i
                break

    if start_idx == -1:
        return [], False

    test_elements = []
    has_keys_table = False

    # Поиск содержимого теста до стоп-сигналов или ключей
    for i in range(start_idx, len(body_elements)):
        elem = body_elements[i]
        text = "".join(elem.itertext()).strip().lower()

        # Если встретили ключевые слова другого раздела после начала теста — останавливаемся
        if i > start_idx:
            if "критерии" in text or "наименование:" in text or "представлениевфос:" in text:
                break

            # Нашли ключи теста - завершаем сбор элементов
            if "ключи" in text or "ответы" in text:
                test_elements.append(elem)
                # Ищем таблицу ключей непосредственно следом (в пределах 4 элементов)
                for j in range(i + 1, min(i + 5, len(body_elements))):
                    next_elem = body_elements[j]
                    if next_elem.tag.endswith('tbl'):
                        test_elements.append(next_elem)
                        has_keys_table = True
                        break
                break

        test_elements.append(elem)

    # Жесткое правило: если в РП нет таблицы ключей теста — полностью игнорируем этот блок
    if not has_keys_table:
        return [], False

    return test_elements, True


def insert_elements_after_paragraph(anchor_paragraph, elements: list):
    """Вставляет копии XML-элементов (параграфы и таблицы) строго друг за другом после anchor_paragraph."""
    current_element = anchor_paragraph._element
    for elem in elements:
        new_elem = copy.deepcopy(elem)
        current_element.addnext(new_elem)
        current_element = new_elem


class CompetencyReportGenerator:
    """Класс управления интеграцией тестов из Рабочих Программ."""

    def __init__(self, word_path: str, rp_folder_path: str):
        self.word_path = word_path
        self.rp_folder_path = rp_folder_path

    def generate(self):
        logger.info(f"Загрузка итогового документа: {self.word_path}")
        if not os.path.exists(self.word_path):
            logger.error(f"Файл не найден: {self.word_path}")
            return

        try:
            doc = Document(self.word_path)
        except Exception as e:
            logger.error(f"Не удалось открыть документ: {e}")
            return

        rp_folder = Path(self.rp_folder_path)
        if not rp_folder.exists():
            logger.error(f"Папка с РП не найдена по пути: {self.rp_folder_path}")
            return

        # 1. Сканируем папку с РП и строим карту соответствия
        logger.info("Сканирование папки с Рабочими Программами...")
        rp_map: Dict[str, Path] = {}
        for docx_path in rp_folder.glob("*.docx"):
            if docx_path.name.startswith("~$"):
                continue
            try:
                rp_doc = Document(docx_path)
                title = get_rp_title(rp_doc)
                if title:
                    cleaned_title = clean_name_for_match(title)
                    rp_map[cleaned_title] = docx_path
                    logger.info(f"  -> Найдена РП: '{title}' ({docx_path.name})")
            except Exception as e:
                logger.warning(f"Не удалось прочитать файл РП {docx_path.name}: {e}")

        logger.info(f"Всего успешно проиндексировано РП: {len(rp_map)}")

        # 2. Ищем заглушки в итоговом файле и заменяем их тестами из РП
        logger.info("Интеграция тестов в структуру документа...")
        subject_pattern = re.compile(r'^(Дисциплина|Практика)\s+«(.*?)»')

        paragraphs = list(doc.paragraphs)
        updated_count = 0

        for i, p in enumerate(paragraphs):
            match = subject_pattern.match(p.text.strip())
            if match:
                subj_type = match.group(1)
                subj_name = match.group(2)
                cleaned_subj_name = clean_name_for_match(subj_name)

                # Ищем подходящую РП
                rp_path = rp_map.get(cleaned_subj_name)
                if not rp_path:
                    logger.info(f"  [-] Тест для {subj_type.lower()} «{subj_name}» не найден в РП (оставлен шаблон)")
                    continue

                try:
                    rp_doc = Document(rp_path)
                    test_elements, has_keys_table = find_test_elements(rp_doc)

                    # Если оригинальной таблицы ключей не найдено — полностью пропускаем интеграцию этого теста
                    if not test_elements or not has_keys_table:
                        logger.info(
                            f"  [-] В РП '{rp_path.name}' не найдена корректная таблица ключей. Интеграция отменена, оставлен шаблон.")
                        continue

                    # Находим параграф "Проведение работы...", идущий следом за заголовком предмета
                    p_conduct = None
                    for k in range(i + 1, min(i + 5, len(paragraphs))):
                        if "проведение работы" in paragraphs[k].text.lower():
                            p_conduct = paragraphs[k]
                            break

                    if not p_conduct:
                        continue

                    # Находим и удаляем все наши сгенерированные заглушки (пустые строки, вопрос '1.' и пустую таблицу)
                    main_elements = list(doc.element.body)
                    p_idx = main_elements.index(p_conduct._element)

                    elements_to_delete = []
                    for k in range(p_idx + 1, len(main_elements)):
                        elem = main_elements[k]
                        if elem.tag.endswith('tbl'):
                            elements_to_delete.append(elem)
                            break  # Удаляем до первой встреченной пустой таблицы ключей включительно
                        else:
                            elements_to_delete.append(elem)

                    for elem in elements_to_delete:
                        doc.element.body.remove(elem)

                    # Вставляем реальные XML-элементы теста из РП сразу после параграфа проведения работы
                    insert_elements_after_paragraph(p_conduct, test_elements)
                    logger.info(f"  [+] Успешно интегрирован тест для {subj_type.lower()} «{subj_name}»")
                    updated_count += 1

                except Exception as e:
                    logger.error(f"  [!] Ошибка при переносе теста для '{subj_name}': {e}")

        # Сохраняем обновленный документ
        try:
            doc.save(self.word_path)
            print(f"\n[Успешно] Интеграция тестов завершена. Обновлено предметов: {updated_count}")
        except Exception as e:
            logger.error(f"Не удалось сохранить обновленный документ Word: {e}")


def main():
    print("=== Панель управления интеграцией тестов из РП ===")

    user_word_path: str = input(
        "Шаг 1. Введите путь к итоговому файлу Word (например, Оценочные материалы.docx): ").strip()
    user_rp_folder: str = input("Шаг 2. Введите путь к папке с Рабочими Программами (РП): ").strip()

    if not user_word_path:
        user_word_path = "Оценочные материалы.docx"

    if not user_rp_folder:
        user_rp_folder = "."

    print("\nЗапуск процесса интеграции...")
    generator = CompetencyReportGenerator(
        word_path=user_word_path,
        rp_folder_path=user_rp_folder
    )
    generator.generate()


if __name__ == "__main__":
    main()