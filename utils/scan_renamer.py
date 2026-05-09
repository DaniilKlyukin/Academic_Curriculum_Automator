import easyocr
import re


class ScanRenamer:
    def __init__(self):
        self.reader = easyocr.Reader(['ru'], gpu=False)

        self.keywords_p1 = ["рабочая", "программа", "дисциплины", "минобрнауки", "утверждаю", "технический",
                            "университет"]
        self.keywords_p2 = ["составитель", "протокол", "председатель", "согласовано", "руководитель"]
        self.keywords_p3 = ["лист", "согласования", "учебный", "год", "план"]

    def get_text(self, img_path):
        try:
            results = self.reader.readtext(img_path, detail=0)
            return " ".join(results).lower()
        except Exception as e:
            return f"error: {e}"

    def identify_page_type(self, text):
        if not text or "error" in text: return None

        score3 = sum(1 for kw in self.keywords_p3 if kw in text)
        score2 = sum(1 for kw in self.keywords_p2 if kw in text)
        score1 = sum(1 for kw in self.keywords_p1 if kw in text)

        if score3 >= 1: return 3
        if score2 >= 2: return 2
        if score1 >= 1: return 1
        return None

    def extract_discipline(self, text):
        clean_text = re.sub(r'\s+', ' ', text)
        match = re.search(
            r'(?:дисциплины|модуля)\s+["«]?\s*([а-яё\s\-]{5,120})["»]?\s+(?:направление|направленность|уровень)',
            clean_text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return None

    def make_abbreviation(self, text):
        if not text or text == "Unknown": return "UNKNOWN"
        words = re.findall(r'[а-яё]{3,}', text.lower())
        stop_words = {"программа", "рабочая", "дисциплины", "модуля", "бакалавриат", "очная"}
        significant = [w for w in words if w not in stop_words]

        if not significant: return "DOC"

        parts = [w[:4].capitalize() for w in significant]
        return "".join(parts)