import os
import re
from typing import List, Tuple, Optional
from pathlib import Path


class ScanFinder:
    def __init__(self, scans_folder: str, image_extensions: List[str] = None):
        """
        Инициализация класса для поиска сканов

        :param scans_folder: Папка со сканами
        :param image_extensions: Допустимые расширения изображений (по умолчанию ['.jpg', '.jpeg', '.png'])
        """
        self.scans_folder = scans_folder
        self.image_extensions = image_extensions or ['.jpg', '.jpeg', '.png']

    def find_scans_for_program(self, program_name: str) -> Optional[Tuple[str, str, str]]:
        """
        Находит три скана для указанной рабочей программы

        :param program_name: Название рабочей программы (например "РП Б1.О.07 РП АиСД 01.04.04.doc")
        :return: Кортеж из трех путей к сканам или None, если не найдены
        """
        # Извлекаем ключевую часть названия (например "АиСД" из "РП Б1.О.07 РП АиСД 01.04.04.doc")
        base_name = self._extract_base_name(program_name)
        if not base_name:
            return None

        # Ищем все подходящие файлы
        matching_files = self._find_matching_files(base_name)

        # Сортируем по номеру в названии и проверяем, что есть ровно 3 файла
        return self._sort_scans(matching_files)

    def _extract_base_name(self, program_name: str) -> Optional[str]:
        """
        Извлекает базовое название из имени файла рабочей программы
        (например "АиСД" из "РП Б1.О.07 РП АиСД 01.04.04.doc")
        """
        # Удаляем расширение файла
        name_without_ext = Path(program_name).stem

        # Ищем часть между "РП" и датой/номером
        match = re.search(r'РП\s+(\S+)(?=\s|$)', name_without_ext)

        if not match:
            return None

        base_name = match[0].split(' ')[1].strip()

        return base_name if base_name else None

    def _find_matching_files(self, program_name: str) -> List[str]:
        """
        Находит все файлы сканов, соответствующие базовому названию
        """
        matching_files = []

        for filename in os.listdir(self.scans_folder):
            # Проверяем расширение файла
            if not any(filename.lower().endswith(ext) for ext in self.image_extensions):
                continue

            program_img_name, file_extension = os.path.splitext(filename.strip().lower())
            program_img_name = program_img_name[:-1]

            # Проверяем точное совпадение с базовым названием
            if program_name.lower() == program_img_name:
                matching_files.append(os.path.join(self.scans_folder, filename))

        return matching_files

    def _sort_scans(self, files: List[str]) -> Optional[Tuple[str, str, str]]:
        """
        Проверяет, что найдено ровно 3 скана и сортирует их по номеру
        """
        if len(files) != 3:
            return None

        # Извлекаем номера из имен файлов и сортируем
        def get_file_number(file_path):
            filename = Path(file_path).stem
            match = re.search(r'(\d+)', filename)
            return int(match.group(1)) if match else 0

        sorted_files = sorted(files, key=get_file_number)

        return tuple(sorted_files)
