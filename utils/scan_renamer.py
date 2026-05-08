import pytesseract
from PIL import Image, ImageOps, ImageEnhance
import re
import os
import shutil


class ScanRenamer:
    def __init__(self):
        # 1. Пытаемся найти Tesseract в системе автоматически
        self._setup_tesseract()

        self.keywords_p1 = ["рабочая", "программа", "дисциплины", "минобрнауки", "утверждаю"]
        self.keywords_p2 = ["составитель", "протокол", "председатель", "согласовано"]
        self.keywords_p3 = ["лист", "согласования", "учебный", "год"]

    def _setup_tesseract(self):
        """Ищет Tesseract в стандартных путях Windows."""
        standard_paths = [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Users\\' + os.getlogin() + r'\AppData\Local\Tesseract-OCR\tesseract.exe'
        ]
        if shutil.which("tesseract"):
            return  # Уже в PATH
        for path in standard_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                return

    def _preprocess_image(self, img):
        """Улучшает картинку для OCR (ч/б + контраст)."""
        img = img.convert('L')  # В оттенки серого
        img = ImageOps.autocontrast(img)
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(2.0)  # Увеличиваем контраст
        return img

    def get_text(self, img_path):
        try:
            img = Image.open(img_path)
            # Берем чуть больше половины страницы
            w, h = img.size
            img_crop = img.crop((0, 0, w, int(h * 0.6)))
            img_clean = self._preprocess_image(img_crop)

            # Конфигурация для лучшего распознавания кириллицы
            text = pytesseract.image_to_string(img_clean, lang='rus')
            return text.lower()
        except Exception as e:
            return f"error: {e}"

    def identify_page_type(self, text):
        # Проверка на пустой текст
        if not text or "error" in text: return None

        # Считаем количество попаданий ключевых слов
        score3 = sum(1 for kw in self.keywords_p3 if kw in text)
        score2 = sum(1 for kw in self.keywords_p2 if kw in text)
        score1 = sum(1 for kw in self.keywords_p1 if kw in text)

        if score3 >= 1: return 3
        if score2 >= 1: return 2
        if score1 >= 1: return 1
        return None

    def extract_discipline(self, text):
        # Очищаем текст от лишних символов
        clean_text = re.sub(r'[^а-яё\s]', ' ', text)
        match = re.search(r'(?:дисциплины|модуля)\s+([а-яё\s]{5,100})\s+(?:направление|направленность)', clean_text,
                          re.IGNORECASE | re.DOTALL)
        if match:
            return match.group(1).strip()
        return None

    def make_abbreviation(self, text):
        if not text or text == "Unknown": return "UNKNOWN"
        words = re.findall(r'[а-яё]{3,}', text.lower())
        stop_words = {"программа", "рабочая", "дисциплины", "модуля"}
        significant = [w for w in words if w not in stop_words]

        if not significant: return "PRED"

        parts = [w[:4].capitalize() for w in significant]
        return "".join(parts)