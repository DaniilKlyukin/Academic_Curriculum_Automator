import os
import re
import json
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional

from openpyxl import load_workbook

# Настройка логирования
logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
logger = logging.getLogger(__name__)


def parse_semesters_string(sem_str: str) -> List[int]:
    """Интеллектуально разбирает строку семестров контроля (например, '12345', '78', '1-5', '1, 2')."""
    if not sem_str:
        return []

    sem_str = sem_str.strip()
    semesters = set()

    # 1. Если строка состоит только из цифр без разделителей
    if sem_str.isdigit():
        for char in sem_str:
            val = int(char)
            if 1 <= val <= 12:
                semesters.add(val)
        return sorted(list(semesters))

    # 2. Обработка диапазонов вида "1-5"
    range_match = re.match(r"^(\d+)\s*[-—–]\s*(\d+)$", sem_str)
    if range_match:
        start = int(range_match.group(1))
        end = int(range_match.group(2))
        if start <= end and end <= 12:
            return list(range(start, end + 1))

    # 3. Разделение по разделителям
    parts = re.split(r'[,\s;]+', sem_str)
    for part in parts:
        part = part.strip()
        if not part:
            continue

        if part.isdigit():
            if len(part) > 1 and int(part) > 12:
                for char in part:
                    val = int(char)
                    if 1 <= val <= 12:
                        semesters.add(val)
            else:
                val = int(part)
                if 1 <= val <= 12:
                    semesters.add(val)

    return sorted(list(semesters))


class AcademicPlanParser:
    """Парсер учебной нагрузки, метаданных титульного листа и структуры плана из Excel."""

    def __init__(self, excel_path: Path):
        self.excel_path = excel_path

    def _find_plan_sheet(self, wb) -> str:
        """Ищет подходящий лист детального учебного плана, строго исключая сводные листы."""
        # 1. Сначала ищем строгое совпадение с "План"
        if "План" in wb.sheetnames:
            return "План"

        # 2. Ищем лист, содержащий слово "план", но НЕ содержащий слово "свод"
        for name in wb.sheetnames:
            name_lower = name.lower()
            if "план" in name_lower and "свод" not in name_lower:
                return name

        # 3. Резервный поиск любого листа со словом "план"
        for name in wb.sheetnames:
            if "план" in name.lower():
                return name

        raise ValueError(f"Детальный лист учебного плана ('План') не найден в книге. Доступные листы: {wb.sheetnames}")

    def _parse_department_directory(self, wb) -> Dict[int, str]:
        """Строит справочник 'Код кафедры -> Название кафедры' на основе листа 'ПланСвод'."""
        dept_map = {}
        if "ПланСвод" not in wb.sheetnames:
            return dept_map

        sheet = wb["ПланСвод"]
        col_code = 26  # Дефолтная колонка Кода в ПланСвод
        col_name = 27  # Дефолтная колонка Наименования в ПланСвод

        # Динамический поиск колонок Кода и Наименования на случай смещения
        for col in range(1, sheet.max_column + 1):
            for row in range(1, 4):
                val = str(sheet.cell(row=row, column=col).value or "").lower().strip()
                if "закреплен" in val or "кафедр" in val:
                    sub_val = str(sheet.cell(row=3, column=col).value or "").lower().strip()
                    if "код" in sub_val:
                        col_code = col
                    elif "наимен" in sub_val:
                        col_name = col

        # Наполнение справочника
        for row in range(4, sheet.max_row + 1):
            code_val = sheet.cell(row=row, column=col_code).value
            name_val = sheet.cell(row=row, column=col_name).value
            try:
                code_int = int(float(code_val)) if code_val is not None else None
            except (ValueError, TypeError):
                code_int = None

            if code_int is not None and name_val:
                dept_map[code_int] = str(name_val).strip()

        logger.info(f"Справочник кафедр успешно построен из 'ПланСвод' (всего записей: {len(dept_map)})")
        return dept_map

    def _map_semester_columns(self, sheet) -> Dict[int, Dict[str, int]]:
        """Автоматически определяет номера колонок нагрузки (Лек, Лаб, Пр, СР) для каждого семестра."""
        semester_map = {}
        for col in range(1, sheet.max_column + 1):
            cell_val = str(sheet.cell(row=2, column=col).value or "").strip()
            if "семестр" in cell_val.lower():
                m = re.search(r"семестр\s*(\d+)", cell_val, re.IGNORECASE)
                if m:
                    sem_num = int(m.group(1))
                    semester_map[sem_num] = {
                        "Lek": col,
                        "Lab": col + 1,
                        "Pr": col + 2,
                        "CP": col + 3
                    }
        return semester_map

    def _find_department_column(self, sheet) -> int:
        """Автоматически находит индекс колонки кода закрепленной кафедры."""
        for col in range(1, sheet.max_column + 1):
            for row in range(1, 4):
                cell_val = str(sheet.cell(row=row, column=col).value or "").lower().strip()
                if "закреплен" in cell_val:
                    return col
        return 50

    def _extract_field_by_keyword(self, sheet, regex: re.Pattern, max_cols_offset: int = 15) -> str:
        """Помехоустойчиво ищет значение параметра на листе по ключевому регулярному выражению."""
        for row in range(1, min(sheet.max_row + 1, 150)):  # Ограничение сканирования титула первыми 150 строками
            for col in range(1, sheet.max_column + 1):
                cell_val = str(sheet.cell(row=row, column=col).value or "").strip()
                if not cell_val:
                    continue
                if regex.search(cell_val):
                    # Если значение указано в этой же ячейке после двоеточия
                    if ":" in cell_val:
                        parts = cell_val.split(":", 1)
                        val = parts[1].strip()
                        if val and len(val) > 1:
                            return val
                    # Иначе ищем значение в ячейках справа
                    for offset in range(1, max_cols_offset):
                        if col + offset <= sheet.max_column:
                            test_val = str(sheet.cell(row=row, column=col + offset).value or "").strip()
                            if test_val and len(test_val) > 1:
                                return test_val
        return ""

    def _extract_title_metadata(self, wb) -> Dict[str, str]:
        """Парсит лист 'Титул' и собирает метаданные учебного плана для генерации заголовков РП."""
        metadata = {
            "qualification": "",
            "education_form": "",
            "study_duration": "",
            "start_year": "",
            "academic_year": "",
            "fgos_standard": "",
            "faculty": "",
            "department": "",
            "direction_code": "",
            "direction_name": "",
            "profile": ""
        }

        if "Титул" not in wb.sheetnames:
            logger.warning("Лист 'Титул' отсутствует в книге. Пропуск сбора общих метаданных.")
            return metadata

        sheet = wb["Титул"]

        # Паттерны поиска метаданных
        pat_qualification = re.compile(r"квалификация", re.IGNORECASE)
        pat_education_form = re.compile(r"форма\s+обучения", re.IGNORECASE)
        pat_study_duration = re.compile(r"срок\s+(?:получения|обучения)", re.IGNORECASE)
        pat_start_year = re.compile(r"год\s+начала\s+подготовки", re.IGNORECASE)
        pat_academic_year = re.compile(r"учебный\s+год", re.IGNORECASE)
        pat_fgos = re.compile(r"(?:стандарт|фгос)\b", re.IGNORECASE)
        pat_faculty = re.compile(r"факультет\b|институт\b", re.IGNORECASE)
        pat_department = re.compile(r"кафедра\b", re.IGNORECASE)
        pat_profile = re.compile(r"профиль\b|направленность\b", re.IGNORECASE)
        pat_direction = re.compile(r"направление\b|специальность\b", re.IGNORECASE)

        # Сбор данных с листа "Титул"
        metadata["qualification"] = self._extract_field_by_keyword(sheet, pat_qualification)
        metadata["education_form"] = self._extract_field_by_keyword(sheet, pat_education_form)
        metadata["study_duration"] = self._extract_field_by_keyword(sheet, pat_study_duration)
        metadata["start_year"] = self._extract_field_by_keyword(sheet, pat_start_year)
        metadata["academic_year"] = self._extract_field_by_keyword(sheet, pat_academic_year)
        metadata["fgos_standard"] = self._extract_field_by_keyword(sheet, pat_fgos)
        metadata["faculty"] = self._extract_field_by_keyword(sheet, pat_faculty)
        metadata["department"] = self._extract_field_by_keyword(sheet, pat_department)
        metadata["profile"] = self._extract_field_by_keyword(sheet, pat_profile)

        # Парсинг направления подготовки (Код + Название)
        direction_raw = self._extract_field_by_keyword(sheet, pat_direction)
        if direction_raw:
            code_match = re.search(r"(\d{2}\.\d{2}\.\d{2})", direction_raw)
            if code_match:
                metadata["direction_code"] = code_match.group(1)
                name_part = direction_raw.replace(metadata["direction_code"], "").strip()
                metadata["direction_name"] = re.sub(r"^[-\s,.:\)]+", "", name_part).strip()
            else:
                metadata["direction_name"] = direction_raw

        logger.info(f"Успешно извлечены метаданные плана: {metadata}")
        return metadata

    def parse(self) -> Dict[str, Any]:
        """Парсит лист плана и возвращает структуру академической нагрузки и метаданные блоков."""
        wb = load_workbook(str(self.excel_path.absolute()), data_only=True)

        # Сбор общих метаданных учебного плана (форма обучения, квалификация, стандарты и др.)
        metadata = self._extract_title_metadata(wb)

        # Новое: Строим справочник соответствия кодов и названий кафедр из ПланСвод
        dept_directory = self._parse_department_directory(wb)

        sheet_name = self._find_plan_sheet(wb)
        sheet = wb[sheet_name]
        logger.info(f"Анализ учебной нагрузки на листе: '{sheet_name}'")

        # Базовые колонки (структура ИжГТУ)
        col_index = 2  # Код дисциплины (например, Б1.О.11)
        col_name = 3  # Наименование дисциплины

        # Колонки форм контроля
        col_exam = 4
        col_credit = 5
        col_graded_credit = 6
        col_kp = 7
        col_kr = 8

        col_credits = 9  # Колонка Зачетных Единиц (з.е.) - Экспертное значение

        col_dept = self._find_department_column(sheet)
        logger.info(f"Определен столбец кода закрепленной кафедры -> {col_dept}")

        semester_cols = self._map_semester_columns(sheet)
        logger.info(f"Обнаружена структура семестровых колок: {list(semester_cols.keys())}")

        disciplines = {}

        # Переменные для отслеживания текущего контекста иерархии
        current_block = ""
        current_part = ""
        current_elective_group: Optional[Dict[str, str]] = None

        for row in range(4, sheet.max_row + 1):
            idx_val = str(sheet.cell(row=row, column=col_index).value or "").strip()

            # --- СЦЕНАРИЙ А: Строка заголовка (Блок или Подблок) ---
            if not idx_val:
                row_texts = [str(sheet.cell(row=row, column=col).value or "").strip() for col in range(1, 6)]
                full_row_text = " ".join([t for t in row_texts if t])

                # 1. Поиск названия Блока
                if re.search(r"Блок\s+\d", full_row_text, re.IGNORECASE) or re.search(r"^ФТД\.", full_row_text,
                                                                                      re.IGNORECASE):
                    for t in row_texts:
                        if re.search(r"Блок\s+\d", t, re.IGNORECASE) or re.search(r"^ФТД\.", t, re.IGNORECASE):
                            current_block = t
                            current_part = ""  # Сбрасываем часть при смене блока
                            current_elective_group = None
                            logger.info(f"Текущий Блок -> {current_block}")
                            break
                    continue

                # 2. Поиск названия Части (Подблока)
                if "обязательная часть" in full_row_text.lower():
                    current_part = "Обязательная часть"
                    current_elective_group = None
                    logger.info(f"Текущая Часть -> {current_part}")
                    continue
                elif "часть, формируемая" in full_row_text.lower():
                    current_part = "Часть, формируемая участниками образовательных отношений"
                    current_elective_group = None
                    logger.info(f"Текущая Часть -> {current_part}")
                    continue

                continue

            # --- СЦЕНАРИЙ Б: Строка элемента (Дисциплина, Практика, ГИА, Факультатив) ---
            if re.match(r"^(Б\d|ФТД)", idx_val):
                name_val = str(sheet.cell(row=row, column=col_name).value or "").strip()
                if not name_val:
                    continue

                # 1. Определение заголовка элективной группы (Дисциплина по выбору)
                if re.match(r"^Б\d+\.[ОВ]\.ДВ\.\d+$", idx_val):
                    current_elective_group = {
                        "code": idx_val,
                        "name": name_val
                    }
                    logger.info(f"Обнаружена группа выбора -> {idx_val}: {name_val}")
                    continue  # Пропускаем строку-заголовок (она не несет нагрузки)

                # Проверка: относится ли текущая строка к активной группе элективов
                active_elective = None
                if current_elective_group:
                    if idx_val.startswith(current_elective_group["code"]):
                        active_elective = current_elective_group
                    else:
                        current_elective_group = None  # Вышли за пределы группы

                dept_val = sheet.cell(row=row, column=col_dept).value
                try:
                    department_code = int(float(dept_val)) if dept_val is not None else None
                except (ValueError, TypeError):
                    department_code = None

                if department_code is None:
                    continue

                # Извлечение текстового названия кафедры по ее коду через справочник ПланСвод
                department_name = dept_directory.get(department_code, "")

                # 3. Извлечение зачетных единиц (з.е.) из 9-го столбца
                credits_cell = sheet.cell(row=row, column=col_credits).value
                try:
                    credit_units = float(credits_cell) if credits_cell is not None else 0.0
                    if credit_units.is_integer():
                        credit_units = int(credit_units)
                except (ValueError, TypeError):
                    credit_units = 0

                # 4. Извлечение форм контроля
                exam_sems = str(sheet.cell(row=row, column=col_exam).value or "").strip()
                credit_sems = str(sheet.cell(row=row, column=col_credit).value or "").strip()
                graded_credit_sems = str(sheet.cell(row=row, column=col_graded_credit).value or "").strip()
                kp_sems = str(sheet.cell(row=row, column=col_kp).value or "").strip()
                kr_sems = str(sheet.cell(row=row, column=col_kr).value or "").strip()

                # 5. Распределение семестровой нагрузки
                load_by_semester = {}
                total_lectures = 0
                total_labs = 0
                total_practicals = 0
                total_cp = 0

                for sem_num, cols in semester_cols.items():
                    lek = sheet.cell(row=row, column=cols["Lek"]).value
                    lab = sheet.cell(row=row, column=cols["Lab"]).value
                    pr = sheet.cell(row=row, column=cols["Pr"]).value
                    cp = sheet.cell(row=row, column=cols["CP"]).value

                    try:
                        lek_h = int(float(lek)) if lek else 0
                        lab_h = int(float(lab)) if lab else 0
                        pr_h = int(float(pr)) if pr else 0
                        cp_h = int(float(cp)) if cp else 0
                    except (ValueError, TypeError):
                        continue

                    if lek_h > 0 or lab_h > 0 or pr_h > 0 or cp_h > 0:
                        load_by_semester[str(sem_num)] = {
                            "lectures": lek_h,
                            "laboratory_works": lab_h,
                            "practical_classes": pr_h,
                            "self_study": cp_h
                        }
                        total_lectures += lek_h
                        total_labs += lab_h
                        total_practicals += pr_h
                        total_cp += cp_h

                disciplines[idx_val] = {
                    "code": idx_val,
                    "name": name_val,
                    "department_code": department_code,
                    "department_name": department_name,  # Сохраняем текстовое имя кафедры
                    "credit_units": credit_units,
                    "structure": {
                        "block": current_block,
                        "part": current_part,
                        "elective_group": active_elective
                    },
                    "control_forms": {
                        "exams": parse_semesters_string(exam_sems),
                        "credits": parse_semesters_string(credit_sems),
                        "graded_credits": parse_semesters_string(graded_credit_sems),
                        "course_projects": parse_semesters_string(kp_sems),
                        "course_works": parse_semesters_string(kr_sems)
                    },
                    "total_hours": {
                        "lectures": total_lectures,
                        "laboratory_works": total_labs,
                        "practical_classes": total_practicals,
                        "self_study": total_cp,
                        "total": total_lectures + total_labs + total_practicals + total_cp
                    },
                    "load_by_semester": load_by_semester
                }

        return {
            "metadata": metadata,
            "disciplines": disciplines
        }


def main():
    print("=== Модуль разбора академической нагрузки и планов ===")
    user_excel = input("Введите путь к файлу Excel (например, plan.xlsx): ").strip()
    if not user_excel:
        user_excel = "plan.xlsx"

    excel_path = Path(user_excel)
    if not excel_path.exists():
        print("Ошибка: Файл плана не найден.")
        return

    user_output_dir = input(
        "Введите путь к папке для сохранения результатов (по умолчанию 'services/rp_generator'): ").strip()
    if not user_output_dir:
        user_output_dir = "services/rp_generator"

    parser = AcademicPlanParser(excel_path)
    try:
        data = parser.parse()

        output_dir = Path(user_output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "rp_academic_workload.json"

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"[Успешно] Данные по академической нагрузке сохранены в:\n{output_path.absolute()}")
    except Exception as e:
        logger.error(f"Ошибка при обработке файла: {e}", exc_info=True)


if __name__ == "__main__":
    main()