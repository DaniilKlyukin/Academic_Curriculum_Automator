import os
import re
import difflib
from typing import List, Tuple, Optional
from pathlib import Path


class ScanFinder:
    def __init__(self, scans_folder: str, extensions: List[str] = None, threshold: float = 0.8):
        self.scans_folder = scans_folder
        self.extensions = extensions or ['.jpg', '.jpeg', '.png', '.pdf']
        self.threshold = threshold

    def _normalize(self, text: str) -> str:
        """Очистка строки от мусора, пробелов и знаков препинания."""
        if not text:
            return ""
        text = Path(text).stem
        text = re.sub(r'^РП\s*', '', text, flags=re.IGNORECASE)
        return re.sub(r'[^a-zа-я0-9]', '', text.lower())

    def find_scans_for_program(self, program_name: str) -> Optional[Tuple[str, str, str]]:
        norm_program = self._normalize(program_name)
        if not norm_program:
            return None

        candidates = {'1': [], '2': [], '3': []}

        all_files = [f for f in os.listdir(self.scans_folder)
                     if any(f.lower().endswith(ext) for ext in self.extensions)]

        for f in all_files:
            match = re.search(r'([123])\.(?:png|jpg|jpeg|pdf)$', f.lower())
            if not match:
                continue

            index = match.group(1)
            scan_base_raw = re.sub(r'[123]\.(?:png|jpg|jpeg|pdf)$', '', f, flags=re.IGNORECASE)
            norm_scan = self._normalize(scan_base_raw)

            similarity = difflib.SequenceMatcher(None, norm_program, norm_scan).ratio()

            if similarity >= self.threshold:
                candidates[index].append((similarity, os.path.join(self.scans_folder, f)))

        result = []
        for i in ['1', '2', '3']:
            if not candidates[i]:
                return None
            best_match = sorted(candidates[i], key=lambda x: x[0], reverse=True)[0]
            result.append(best_match[1])

        return tuple(result)