import os
import re
import difflib
from typing import List, Tuple, Optional
from pathlib import Path
from collections import defaultdict


class ScanFinder:
    def __init__(self, scans_folder: str, extensions: List[str] = None, threshold: float = 0.75):
        self.scans_folder = scans_folder
        self.extensions = extensions or ['.jpg', '.jpeg', '.png', '.pdf']
        self.threshold = threshold

    def _normalize(self, text: str) -> str:
        """Очищает строку: РП Б1.О.01 Математика -> б1о01математика"""
        if not text: return ""
        stem = Path(text).stem
        base = re.sub(r'^РП\s+', '', stem, flags=re.IGNORECASE)
        return re.sub(r'[^a-zа-я0-9]', '', base.lower())

    def find_scans_for_program(self, program_name: str) -> Optional[Tuple[str, str, str]]:
        norm_doc = self._normalize(program_name)
        groups = defaultdict(dict)

        for root, _, filenames in os.walk(self.scans_folder):
            for f in filenames:
                if any(f.lower().endswith(ext) for ext in self.extensions):
                    match = re.search(r'([123])\.(?:png|jpg|jpeg|pdf)$', f.lower())
                    if not match:
                        continue

                    idx = match.group(1)
                    raw_base = re.sub(r'[123]\.(?:png|jpg|jpeg|pdf)$', '', f, flags=re.IGNORECASE)
                    norm_base = self._normalize(raw_base)
                    groups[norm_base][idx] = os.path.join(root, f)

        scored_groups = []
        for norm_base, files in groups.items():
            if len(files) < 3:
                continue

            similarity = difflib.SequenceMatcher(None, norm_doc, norm_base).ratio()
            if similarity >= self.threshold:
                scored_groups.append((similarity, files))

        if not scored_groups:
            return None

        scored_groups.sort(key=lambda x: x[0], reverse=True)
        best_group_files = scored_groups[0][1]
        return (best_group_files['1'], best_group_files['2'], best_group_files['3'])