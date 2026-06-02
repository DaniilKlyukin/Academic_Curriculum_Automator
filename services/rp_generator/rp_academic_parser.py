import os
import re
import json
import logging
from pathlib import Path
from typing import Dict, Any, List

from openpyxl import load_workbook

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
logger = logging.getLogger(__name__)


def parse_semesters_string(sem_str: str) -> List[int]:
    """Интеллектуально разбирает строку семестров контроля (например, '12345', '78', '1-5', '1, 2')."""
    if not sem_str:
        return []

    sem_str = sem_str.strip()
    semesters = set()

    # 1. Если строка состоит только из цифр без разделителей (например, "12345" или "78")
    if sem_str.isdigit():
        for char in sem_str:
            val = int(char)
            if 1 <= val <= 12:  # Диапазон семестров (бакалавриат, магистратура, специалитет)
                semesters.add(val)
        return sorted(list(semesters))

    # 2. Обработка диапазонов вида "1-5" или "1 - 3"
    range_match = re.match(r"^(\d+)\s*[-—–]\s*(\d+)$", sem_str)
    if range_match:
        start = int(range_match.group(1))
        end = int(range_match.group(2))
        if start <= end and end <= 12:
            return list(range(start, end + 1))

    # 3. Разделение по запятым, точкам с запятой или пробелам
    parts = re.split(r'[,\s;]+', sem_str)
    for part in parts:
        part = part.strip()
        if not part:
            continue

        if part.isdigit():
            # Если часть состоит из нескольких цифр подряд без разделителей (например, "78" внутри списка)
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
    """Парсер учебной нагрузки и форм промежуточной аттестации из Excel."""

    def __init__(self, excel_path: Path):
        self.excel_path = excel_path

    def _find_plan_sheet(self, wb) -> str:
        """Ищет подходящий лист учебного плана."""
        for name in ["План", "ПланСвод"]:
            if name in wb.sheetnames:
                return name
        for name in wb.sheetnames:
            if "план" in name.lower():
                return name
        raise ValueError(f"Лист учебного плана не найден в книге. Доступные листы: {wb.sheetnames}")

    def _map_semester_columns(self, sheet) -> Dict[int, Dict[str, int]]:
        """Автоматически определяет номера колонок нагрузки (Лек, Лаб, Пр, СР) для каждого семестра (1-8)."""
        semester_map = {}
        # Сканируем заголовки семестров во 2-й строке
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
        # Дефолтный fallback на колонку 50 (согласно структуре ИжГТУ)
        return 50

    def parse(self) -> Dict[str, Any]:
        wb = load_workbook(str(self.excel_path.absolute()), data_only=True)
        sheet_name = self._find_plan_sheet(wb)
        sheet = wb[sheet_name]
        logger.info(f"Анализ учебной нагрузки на листе: '{sheet_name}'")

        # Настройка базовых колонок (согласно структуре ИжГТУ)
        col_index = 2  # Код дисциплины (например, Б1.О.11)
        col_name = 3  # Наименование дисциплины

        # Колонки форм контроля (строка 3)
        col_exam = 4  # Экзамен
        col_credit = 5  # Зачет
        col_graded_credit = 6  # Зачет с оценкой
        col_kp = 7  # Курсовой проект
        col_kr = 8  # Курсовая работа

        # Динамический поиск колонки закрепленной кафедры
        col_dept = self._find_department_column(sheet)
        logger.info(f"Определен столбец кода закрепленной кафедры -> {col_dept}")

        semester_cols = self._map_semester_columns(sheet)
        logger.info(f"Обнаружена структура семестровых колонок: {list(semester_cols.keys())}")

        disciplines = {}

        for row in range(4, sheet.max_row + 1):
            idx_val = str(sheet.cell(row=row, column=col_index).value or "").strip()
            # Нас интересуют только строки дисциплин и практик (Б1, Б2, Б3, ФТД)
            if re.match(r"^(Б\d|ФТД)", idx_val):
                name_val = str(sheet.cell(row=row, column=col_name).value or "").strip()
                if not name_val:
                    continue

                # Извлечение кода закрепленной кафедры
                dept_val = sheet.cell(row=row, column=col_dept).value
                try:
                    department_code = int(float(dept_val)) if dept_val is not None else None
                except (ValueError, TypeError):
                    department_code = None

                # Извлечение форм контроля
                exam_sems = str(sheet.cell(row=row, column=col_exam).value or "").strip()
                credit_sems = str(sheet.cell(row=row, column=col_credit).value or "").strip()
                graded_credit_sems = str(sheet.cell(row=row, column=col_graded_credit).value or "").strip()
                kp_sems = str(sheet.cell(row=row, column=col_kp).value or "").strip()
                kr_sems = str(sheet.cell(row=row, column=col_kr).value or "").strip()

                # Парсинг распределения нагрузки по семестрам
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

                    # Считаем только активные семестры (где нагрузка > 0)
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
                    "department_code": department_code,  # Сохраняем код читающей кафедры
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

        return disciplines


def main():
    print("=== Модуль разбора академической нагрузки и планов ===")
    user_excel = input("Введите путь к файлу Excel (например, plan.xlsx): ").strip()
    if not user_excel:
        user_excel = "plan.xlsx"

    excel_path = Path(user_excel)
    if not excel_path.exists():
        print("Ошибка: Файл плана не найден.")
        return

    parser = AcademicPlanParser(excel_path)
    try:
        data = parser.parse()
        output_path = Path("services/rp_generator/rp_academic_workload.json")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"[Успешно] Данные по академической нагрузке сохранены в:\n{output_path.absolute()}")
    except Exception as e:
        logger.error(f"Ошибка при обработке файла: {e}", exc_info=True)


if __name__ == "__main__":
    main()