import os
import re
import difflib
from pathlib import Path
from collections import defaultdict


class ScanFinder:
    def __init__(self, scans_dir, threshold=0.6):
        self.scans_dir = scans_dir
        self.threshold = threshold
        self.groups = self._index_scans()

    def _normalize(self, text: str) -> str:
        """Очистка: всё в нижний регистр, только буквы и цифры."""
        if not text: return ""
        stem = Path(text).stem
        base = re.sub(r'^РП\s+', '', stem, flags=re.IGNORECASE)
        return re.sub(r'[^a-zа-я0-9]', '', base.lower())

    def _index_scans(self):
        """Сканирует папку и группирует файлы по именам (без 1,2,3 в конце)."""
        groups = defaultdict(dict)
        for root, _, filenames in os.walk(self.scans_dir):
            for f in filenames:
                if f.lower().endswith(('.jpg', '.jpeg', '.png')):
                    match = re.search(r'([123])\.(?:jpg|jpeg|png)$', f.lower())
                    if match:
                        idx = match.group(1)
                        raw_base = re.sub(r'[123]\.(?:jpg|jpeg|png)$', '', f, flags=re.IGNORECASE)
                        norm_base = self._normalize(raw_base)
                        groups[norm_base][idx] = os.path.join(root, f)
        return groups

    def find_scans_for_program(self, doc_name: str):
        norm_doc = self._normalize(doc_name)
        best_match = None
        max_score = 0

        for norm_base, files in self.groups.items():
            if not norm_base: continue

            if norm_base in norm_doc:
                score = len(norm_base) / len(norm_doc) + 1.0
            else:
                score = difflib.SequenceMatcher(None, norm_base, norm_doc).ratio()

            if score > max_score and score >= self.threshold:
                max_score = score
                best_match = files

        if best_match and len(best_match) >= 3:
            return [best_match['1'], best_match['2'], best_match['3']], round(max_score, 2)

        return [], 0