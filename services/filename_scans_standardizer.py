import os
import re
import logging
import io
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Any
from openpyxl import load_workbook
from PIL import Image
import fitz
import easyocr
import numpy as np

# Настройка логирования
logger = logging.getLogger(__name__)

# Ленивая инициализация EasyOCR, чтобы не тратить ресурсы до начала работы
_easyocr_reader: Optional[easyocr.Reader] = None


def get_ocr_reader() -> easyocr.Reader:
    """Инициализирует и возвращает глобальный объект распознавания EasyOCR."""
    global _easyocr_reader
    if _easyocr_reader is None:
        logger.info("Инициализация библиотеки EasyOCR (при первом запуске будут загружены языковые модели)...")
        # Инициализируем распознавание русского и английского текста на процессоре (GPU=False для совместимости)
        _easyocr_reader = easyocr.Reader(['ru', 'en'], gpu=False)
    return _easyocr_reader


def clean_name_for_match(name: str) -> str:
    """Очищает строку от кавычек, пробелов, регистра и всех типов дефисов/тире

    для максимально надежного сравнения текста.
    """
    if not name:
        return ""
    s = name.lower()
    s = "".join(s.split())
    # Удаляем кавычки и знаки препинания
    s = s.replace("«", "").replace("»", "").replace('"', '').replace("'", "")
    # Принудительно удаляем любые типы дефисов и тире
    s = s.replace("-", "").replace("—", "").replace("–", "")
    return s


def dice_similarity(s1: str, s2: str) -> float:
    """Вычисляет взаимное сходство Сёренсена-Диса между двумя строками (от 0.0 до 1.0)."""
    if not s1 or not s2:
        return 0.0
    if s1 == s2:
        return 1.0

    if len(s1) < 2 or len(s2) < 2:
        return 1.0 if s1 in s2 or s2 in s1 else 0.0

    b1 = set(s1[i:i + 2] for i in range(len(s1) - 1))
    b2 = set(s2[i:i + 2] for i in range(len(s2) - 1))

    overlap = len(b1 & b2)
    return 2.0 * overlap / (len(b1) + len(b2))


def abbreviate_word(word: str) -> str:
    """Сокращает слово до первой согласной после первой гласной."""
    vowels = set("аеёиоуыэюя")
    word_clean = "".join([c for c in word if c.isalpha()])
    if not word_clean:
        return ""

    first_vowel_idx = -1
    for idx, char in enumerate(word_clean):
        if char.lower() in vowels:
            first_vowel_idx = idx
            break

    if first_vowel_idx == -1:
        return word_clean

    first_consonant_after_vowel_idx = -1
    for idx in range(first_vowel_idx + 1, len(word_clean)):
        if word_clean[idx].lower() not in vowels:
            first_consonant_after_vowel_idx = idx
            break

    if first_consonant_after_vowel_idx == -1:
        return word_clean

    return word_clean[:first_consonant_after_vowel_idx + 1]


def abbreviate_discipline(name: str) -> str:
    """Сокращает название дисциплины (например, Объектно-ориентированное программирование -> ОбОрПрог)."""
    normalized_name = name.replace("-", " ")
    words = normalized_name.split()

    stop_words = {"на", "по", "в", "с", "для", "под", "о", "об", "за", "из", "от", "до", "без"}

    abbr_words = []
    for w in words:
        if w.lower() in stop_words:
            continue

        abbr = abbreviate_word(w)
        if not abbr:
            continue

        if w.lower() == "и":
            abbr_words.append("и")
        else:
            abbr_words.append(abbr.capitalize())

    return "".join(abbr_words)


def load_disciplines_from_excel(excel_path: Path) -> List[str]:
    """Считывает список названий всех дисциплин из вкладки 'План' Excel."""
    wb = load_workbook(str(excel_path.absolute()), data_only=True)
    disciplines = []

    plan_sheet_name = "План"
    if plan_sheet_name not in wb.sheetnames:
        for name in wb.sheetnames:
            if "план" in name.lower() and "свод" not in name.lower():
                plan_sheet_name = name
                break

    if plan_sheet_name in wb.sheetnames:
        sheet = wb[plan_sheet_name]
        for row in range(1, sheet.max_row + 1):
            code_val = str(sheet.cell(row=row, column=2).value or "").strip()
            name_val = str(sheet.cell(row=row, column=3).value or "").strip()
            if code_val and name_val:
                if re.match(r"^[БФТД]\d+", code_val):
                    if name_val not in disciplines:
                        disciplines.append(name_val)
    return disciplines


def extract_text_from_file(file_path: Path) -> str:
    """Распознает и извлекает русский текст из PDF-файла или картинки-скана в памяти."""
    ext = file_path.suffix.lower()
    text = ""

    try:
        if ext == ".pdf":
            # Открываем PDF в памяти через PyMuPDF (супербыстро)
            doc = fitz.open(file_path)

            # 1. Пробуем извлечь векторный текст напрямую
            for page in doc:
                text += page.get_text() or ""

                # 2. Если векторного текста нет, рендерим первую страницу в картинку и шлем в EasyOCR
                if not text.strip() and len(doc) > 0:
                    page = doc[0]
                    pix = page.get_pixmap(dpi=150)
                    img_bytes = pix.tobytes("png")

                    reader = get_ocr_reader()
                    results = reader.readtext(img_bytes)
                    text = "\n".join([res[1] for res in results])  # Склеиваем по строкам

        elif ext in (".jpg", ".jpeg", ".png"):
            # Читаем картинку через EasyOCR напрямую
            img = Image.open(file_path).convert('RGB')
            img_array = np.array(img)

            reader = get_ocr_reader()
            results = reader.readtext(img_array)
            text = "\n".join([res[1] for res in results])  # Склеиваем по строкам


    except Exception as e:
        logger.warning(f"Ошибка оптического распознавания файла {file_path.name}: {e}")

    return text


def classify_page_text(text: str) -> Optional[int]:
    """Определяет тип страницы скана по ключевым маркерам текста."""
    t = text.lower()

    # Страница 3: Лист согласования
    if any(x in t for x in ["лист согласования", "согласована на ведение", "учебный год", "подпись и дата"]):
        return 3

    # Страница 2: Составители
    if any(x in t for x in ["составитель", "разработчик", "составители", "разработчики", "ф.и.о. (полностью)"]):
        return 2

    # Страница 1: Титульный лист
    if any(x in t for x in ["рабочая программа", "утверждаю", "направление подготовки", "направленность"]):
        return 1

    return None


def match_discipline_from_text(text: str, plan_disciplines: List[str]) -> Optional[str]:
    """Находит наиболее подходящую дисциплину, сравнивая её построчно

    с распознанным текстом скана с помощью коэффициента Сёренсена-Диса.
    """
    # Выводим распознанный текст построчно в лог для отладки
    raw_lines = [line.strip() for line in text.split('\n') if line.strip()]
    logger.info("Распознанный текст на странице:")
    for line in raw_lines:
        logger.info(f"  | {line}")

    # Очищаем строки текста от пробелов и дефисов
    lines_clean = [clean_name_for_match(line) for line in raw_lines if len(line.strip()) > 3]
    if not lines_clean:
        return None

    best_match = None
    max_similarity = 0.0

    for disc in plan_disciplines:
        disc_clean = clean_name_for_match(disc)
        if not disc_clean:
            continue

        # Сравниваем эталонное название дисциплины с каждой строкой скана по отдельности
        for line in lines_clean:
            sim = dice_similarity(disc_clean, line)
            if sim > max_similarity:
                max_similarity = sim
                best_match = disc

    if best_match:
        logger.info(f"  -> Лучший кандидат: '{best_match}' (построчное сходство: {max_similarity:.1%})")

    # Увеличиваем порог уверенности до 70% (для построчного сравнения это очень надежно)
    if max_similarity >= 0.70:
        return best_match

    return None


class ScanRenamer:
    """Класс группировки и переименования сканов РП по тройкам."""

    def __init__(self, scans_dir: str, excel_path: str):
        self.scans_dir = Path(scans_dir)
        self.excel_path = Path(excel_path)

    def run(self):
        if not self.scans_dir.exists() or not self.excel_path.exists():
            logger.error("Проверьте правильность путей к папке со сканами и файлу Excel.")
            return

        logger.info("Загрузка списка дисциплин из Excel...")
        plan_disciplines = load_disciplines_from_excel(self.excel_path)
        logger.info(f"Загружено дисциплин: {len(plan_disciplines)}")

        # Собираем все поддерживаемые файлы сканов и сортируем их по имени
        valid_extensions = {".jpg", ".jpeg", ".png", ".pdf"}
        files = sorted([
            f for f in self.scans_dir.iterdir()
            if f.suffix.lower() in valid_extensions and not f.name.startswith("~$")
        ])

        logger.info(f"Найдено файлов для анализа: {len(files)}")

        current_discipline_abbr = None
        updated_count = 0

        for file_path in files:
            logger.info(f"Анализ файла '{file_path.name}'...")
            text = extract_text_from_file(file_path)
            page_type = classify_page_text(text)

            if page_type == 1:
                # Нашли новый титульный лист
                matched_disc = match_discipline_from_text(text, plan_disciplines)
                if matched_disc:
                    current_discipline_abbr = abbreviate_discipline(matched_disc)
                    logger.info(f"  -> Распознана дисциплина: '{matched_disc}' ({current_discipline_abbr})")
                else:
                    current_discipline_abbr = "НеизвестнаяДисциплина"
                    logger.warning("  -> [!] Не удалось надежно сопоставить дисциплину по тексту титульного листа.")

                new_name = f"{current_discipline_abbr}1{file_path.suffix}"

            elif page_type == 2:
                # Лист составителей
                if not current_discipline_abbr:
                    current_discipline_abbr = "ПропущенТитул"
                new_name = f"{current_discipline_abbr}2{file_path.suffix}"

            elif page_type == 3:
                # Лист согласования
                if not current_discipline_abbr:
                    current_discipline_abbr = "ПропущенТитул"
                new_name = f"{current_discipline_abbr}3{file_path.suffix}"

            else:
                logger.warning("  -> [!] Не удалось классифицировать тип страницы.")
                continue

            # Переименование с обходом конфликтов имен
            new_path = file_path.with_name(new_name)
            counter = 1
            while new_path.exists():
                stem = Path(new_name).stem
                new_path = file_path.with_name(f"{stem}_{counter}{file_path.suffix}")
                counter += 1

            file_path.rename(new_path)
            logger.info(f"  [+] Успешно переименован: '{file_path.name}' -> '{new_path.name}'")
            updated_count += 1

        print(f"\n[Готово] Обработка сканов завершена. Успешно переименовано файлов: {updated_count}")


def main():
    print("=== Интеллектуальное переименование сканов РП ===")
    user_scans_dir = input("Шаг 1. Введите путь к папке со сканами РП (картинки/pdf): ").strip().strip('"')
    user_excel_path = input("Шаг 2. Введите путь к файлу учебного плана Excel (plan.xlsx): ").strip().strip('"')

    if not user_scans_dir:
        user_scans_dir = "."
    if not user_excel_path:
        user_excel_path = "plan.xlsx"

    renamer = ScanRenamer(user_scans_dir, user_excel_path)
    renamer.run()


if __name__ == "__main__":
    main()