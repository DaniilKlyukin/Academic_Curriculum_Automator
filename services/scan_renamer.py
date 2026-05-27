import os
import re
import logging
from typing import List, Optional, Union
import easyocr


logger = logging.getLogger(__name__)


class ScanRenamer:
    """
    Класс для распознавания текста на сканах и классификации страниц.
    Использует EasyOCR для извлечения текста и регулярные выражения для поиска метаданных.
    """

    def __init__(self) -> None:
        self.reader: easyocr.Reader = easyocr.Reader(['ru'], gpu=False)

        self.keywords_p1: List[str] = [
            "рабочая", "программа", "дисциплины", "минобрнауки",
            "утверждаю", "технический", "университет"
        ]
        self.keywords_p2: List[str] = [
            "составитель", "протокол", "председатель", "согласовано", "руководитель"
        ]
        self.keywords_p3: List[str] = [
            "лист", "согласования", "учебный", "год", "план"
        ]

    def get_text(self, img_path: str) -> str:
        """
        Выполняет оптическое распознавание символов (OCR) на изображении.
        """
        try:
            results: List[str] = self.reader.readtext(img_path, detail=0)
            return " ".join(results).lower()
        except Exception as e:
            logger.error(f"Ошибка OCR для файла {img_path}: {e}")
            return f"error: {e}"

    def identify_page_type(self, text: str) -> Optional[int]:
        """
        Определяет порядковый номер страницы (1, 2 или 3) на основе вхождения ключевых слов.
        """
        if not text or "error" in text:
            return None

        score3: int = sum(1 for kw in self.keywords_p3 if kw in text)
        score2: int = sum(1 for kw in self.keywords_p2 if kw in text)
        score1: int = sum(1 for kw in self.keywords_p1 if kw in text)

        if score3 >= 1:
            return 3
        if score2 >= 2:
            return 2
        if score1 >= 1:
            return 1
        return None

    def extract_discipline(self, text: str) -> Optional[str]:
        """
        Пытается извлечь название дисциплины из текста с помощью регулярных выражений.
        """
        clean_text: str = re.sub(r'\s+', ' ', text)
        pattern: str = r'(?:дисциплины|модуля)\s+["«]?\s*([а-яё\s\-]{5,120})["»]?\s+(?:направление|направленность|уровень)'
        match = re.search(pattern, clean_text, re.IGNORECASE)

        if match:
            return match.group(1).strip()
        return None

    def make_abbreviation(self, text: Optional[str]) -> str:
        """
        Создает короткую аббревиатуру из названия дисциплины для именования файлов.
        """
        if not text or text == "Unknown":
            return "UNKNOWN"

        words: List[str] = re.findall(r'[а-яё]{3,}', text.lower())

        stop_words = {"программа", "рабочая", "дисциплины", "модуля", "бакалавриат", "очная"}

        significant: List[str] = [w for w in words if w not in stop_words]

        if not significant:
            return "DOC"

        parts: List[str] = [w[:4].capitalize() for w in significant]
        return "".join(parts)


def main():
    print("=== УМНОЕ ПЕРЕИМЕНОВАНИЕ ===")
    path = input("Путь к папке со сканами (.jpg): ").strip().strip('"')

    if not os.path.isdir(path):
        print("Ошибка: Путь не найден.")
        return

    print("\nЗагрузка ИИ-моделей")
    logic = ScanRenamer()

    files = [f for f in os.listdir(path) if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
    total = len(files)

    print(f"\nАнализ {total} файлов. Пожалуйста, подождите...")
    print(f"{'№':<9} | {'Статус':<8} | {'Аббревиатура':<15} | {'OCR Текст (начало)'}")
    print("-" * 110)

    current_discipline = "Unknown"

    for i, filename in enumerate(files, 1):
        full_path = os.path.join(path, filename)
        text = logic.get_text(full_path)
        page_type = logic.identify_page_type(text)

        status = "SKIP"
        abbr = "---"
        preview = text[:45].strip() + "..."

        if page_type:
            if page_type == 1:
                extracted = logic.extract_discipline(text)
                if extracted:
                    current_discipline = extracted

            abbr = logic.make_abbreviation(current_discipline)
            new_name = f"{abbr}{page_type}.jpg"
            new_path = os.path.join(path, new_name)

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
                logging.error(f"Error {filename}: {e}")

        print(f"[{i:03}/{total:03}] | {status:<8} | {abbr:<15} | {preview}")

    print("\n" + "=" * 40)
    print("ГОТОВО!")
    input("Нажмите Enter для выхода...")


if __name__ == "__main__":
    main()