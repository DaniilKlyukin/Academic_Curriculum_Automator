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