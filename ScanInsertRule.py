from dataclasses import dataclass
from typing import Callable


@dataclass
class ScanInsertRule:
    """Правило вставки скана"""
    scan_index: int  # Индекс скана в найденном списке
    page_resolver: Callable  # Функция определения номера страницы для вставки
