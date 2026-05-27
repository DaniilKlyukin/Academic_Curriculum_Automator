import os
import re
from pathlib import Path
from collections import defaultdict
from typing import Dict, List, Tuple, DefaultDict, Union, Optional
from rapidfuzz import fuzz


class ScanFinder:
    LATIN_TO_CYRILLIC_HOMOGLYPHS = {
        'A': 'А', 'B': 'В', 'C': 'С', 'E': 'Е', 'H': 'Н', 'K': 'К', 'M': 'М',
        'O': 'О', 'P': 'Р', 'T': 'Т', 'X': 'Х', 'Y': 'У',
        'a': 'а', 'c': 'с', 'e': 'е', 'o': 'о', 'p': 'р', 'x': 'х', 'y': 'у',
    }
    _homoglyph_translation = str.maketrans(LATIN_TO_CYRILLIC_HOMOGLYPHS)

    def __init__(self, scans_dir: Union[str, Path], threshold: float = 70.0) -> None:
        """
        threshold от 0 до 100.
        70-80 — оптимально для опечаток и схожих названий.
        """
        self.scans_dir: Union[str, Path] = scans_dir
        self.threshold: float = threshold
        self.noise_pattern: re.Pattern = re.compile(
            r'^(рп|рпд|б1|рабпрог|программа|дисциплины)\s+',
            re.IGNORECASE
        )
        self.groups: DefaultDict[str, Dict[str, str]] = self._index_scans()

    def _clean_extensions(self, filename: str) -> str:
        """Чистит двойные расширения типа .jpeg.jpeg"""
        name: str = filename.lower()
        while True:
            new_name = re.sub(r'\.(jpg|jpeg|png|pdf)$', '', name)
            if new_name == name: break
            name = new_name
        return name

    def _normalize(self, text: str, strip_noise: bool = False) -> str:
        """
        Базовая очистка для индексации.
        Приводит схожие латинские символы к кириллице.
        """
        if not text:
            return ""

        translated: str = text.translate(self._homoglyph_translation)

        base: str = self._clean_extensions(translated.lower())

        if strip_noise:
            base = self.noise_pattern.sub('', base)

        return re.sub(r'[^a-zа-я0-9\s]', '', base).strip()

    def _index_scans(self) -> DefaultDict[str, Dict[str, str]]:
        groups: DefaultDict[str, Dict[str, str]] = defaultdict(dict)
        for root, _, filenames in os.walk(str(self.scans_dir)):
            for f in filenames:
                f_lower = f.lower()
                match = re.search(r'([123])\.(?:jpg|jpeg|png|pdf|.+)+$', f_lower)
                if match:
                    idx: str = match.group(1)
                    raw_base: str = f_lower[:match.start()]
                    norm_base: str = self._normalize(raw_base, strip_noise=False)
                    if norm_base:
                        groups[norm_base][idx] = os.path.join(root, f)
        return groups

    def find_scans_for_program(self, doc_name: str) -> Tuple[List[str], float]:
        norm_doc_raw: str = self._normalize(doc_name, strip_noise=False)
        norm_doc_stripped: str = self._normalize(doc_name, strip_noise=True)

        if not norm_doc_raw:
            return [], 0.0

        best_match_key: Optional[str] = None
        max_score: float = 0.0

        doc_tokens = set(norm_doc_raw.split())

        for norm_base in self.groups.keys():
            norm_base_stripped = self.noise_pattern.sub('', norm_base).strip()

            ignore_words = {'рп', 'рпд', 'б1', 'б2', 'б3', 'о'}
            base_tokens = [t for t in norm_base.split() if t]

            meaningful_base_tokens = [
                t for t in base_tokens
                if len(t) >= 2 and t not in ignore_words
            ]

            if meaningful_base_tokens and all(t in doc_tokens for t in meaningful_base_tokens):
                score = 100.0
            else:
                score_raw = fuzz.WRatio(norm_base, norm_doc_raw)
                score_stripped = fuzz.WRatio(norm_base_stripped, norm_doc_stripped)

                norm_base_compressed = norm_base.replace(" ", "")
                norm_doc_compressed = norm_doc_raw.replace(" ", "")
                score_compressed = fuzz.ratio(norm_base_compressed, norm_doc_compressed)

                score = max(score_raw, score_stripped, score_compressed)

            if score > max_score and score >= self.threshold:
                max_score = score
                best_match_key = norm_base

        if best_match_key:
            files = self.groups[best_match_key]
            if len(files) >= 3:
                return [files.get('1', ''), files.get('2', ''), files.get('3', '')], round(max_score, 2)

        return [], 0.0