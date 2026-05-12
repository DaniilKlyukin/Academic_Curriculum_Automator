import os
import re
from pathlib import Path
from collections import defaultdict
from typing import Dict, List, Tuple, DefaultDict, Union, Optional
from rapidfuzz import fuzz, process


class ScanFinder:
    def __init__(self, scans_dir: Union[str, Path], threshold: float = 70.0) -> None:
        """
        threshold от 0 до 100.
        70-80 — оптимально для опечаток.
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

    def _normalize(self, text: str) -> str:
        """Базовая очистка для индексации"""
        if not text: return ""
        base: str = self._clean_extensions(text)
        base = self.noise_pattern.sub('', base)
        return re.sub(r'[^a-zа-я0-9\s]', '', base.lower()).strip()

    def _index_scans(self) -> DefaultDict[str, Dict[str, str]]:
        groups: DefaultDict[str, Dict[str, str]] = defaultdict(dict)
        for root, _, filenames in os.walk(str(self.scans_dir)):
            for f in filenames:
                f_lower = f.lower()
                match = re.search(r'([123])\.(?:jpg|jpeg|png|pdf|.+)+$', f_lower)
                if match:
                    idx: str = match.group(1)
                    raw_base: str = f_lower[:match.start()]
                    norm_base: str = self._normalize(raw_base)
                    if norm_base:
                        groups[norm_base][idx] = os.path.join(root, f)
        return groups

    def find_scans_for_program(self, doc_name: str) -> Tuple[List[str], float]:
        norm_doc: str = self._normalize(doc_name)
        if not norm_doc:
            return [], 0.0

        best_match_key: Optional[str] = None
        max_score: float = 0.0

        for norm_base in self.groups.keys():

            score = fuzz.WRatio(norm_base, norm_doc)

            if score > max_score and score >= self.threshold:
                max_score = score
                best_match_key = norm_base

        if best_match_key:
            files = self.groups[best_match_key]
            if len(files) >= 3:
                return [files.get('1', ''), files.get('2', ''), files.get('3', '')], round(max_score, 2)

        return [], 0.0