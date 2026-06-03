# -*- coding: utf-8 -*-
"""
Модуль rp_doc_sections.py
Содержит шаблоны разделов РПД и ФОС с динамическим вычислением дат.
"""

import re
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

from services.rp_generator.rp_doc_styles import (
    add_paragraph_with_spacing, set_cell_text, set_row_cant_split,
    set_repeat_table_header, merge_cells_vertically, set_cell_background,
    set_font, set_cell_width, get_department_acronym
)


def get_ugsn_info(direction_code: str, faculty_fallback: str = "") -> tuple[str, str]:
    """Динамически определяет код и наименование укрупненной группы специальностей."""
    if not direction_code or "." not in direction_code:
        return "01.00.00", faculty_fallback or "Математика и механика"
    prefix = direction_code.split(".")[0]
    ugsn_code = f"{prefix}.00.00"

    ugsn_names = {
        "01": "Математика и механика",
        "02": "Компьютерные и информационные науки",
        "03": "Физика и астрономия",
        "04": "Химия",
        "05": "Науки о земле",
        "08": "Техника и технологии строительства",
        "09": "Информатика и вычислительная техника",
        "10": "Информационная безопасность",
        "11": "Электроника, радиотехника и системы связи",
        "12": "Фотоника, приборостроение, оптические и биотехнические системы и технологии",
        "13": "Электро- и теплоэнергетика",
        "15": "Машиностроение",
        "20": "Техносферная безопасность и природообустройство",
        "27": "Управление в технических системах",
        "38": "Экономика и управление",
        "45": "Языкознание и литературоведение"
    }
    ugsn_name = ugsn_names.get(prefix, faculty_fallback or "Естественные и технические науки")
    return ugsn_code, ugsn_name


def generate_title_page(doc, metadata: dict, subj_info: dict, staff: dict):
    """Страница 1: Титульный лист."""
    start_year = metadata.get("start_year") or "2026"
    add_paragraph_with_spacing(doc, "МИНОБРНАУКИ РОССИИ", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_spacing(doc,
                               "Федеральное государственное бюджетное образовательное учреждение высшего образования",
                               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_paragraph_with_spacing(doc,
                               "«Ижевский государственный технический университет имени М.Т. Калашникова»",
                               bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)

    add_paragraph_with_spacing(doc, "УТВЕРЖДАЮ", bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    add_paragraph_with_spacing(doc, f"Декан/Директор\n_____________/ {staff['dean']}",
                               align=WD_ALIGN_PARAGRAPH.RIGHT)
    add_paragraph_with_spacing(doc, f"_________________ {start_year} г.",
                               align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=48)

    add_paragraph_with_spacing(doc, "РАБОЧАЯ ПРОГРАММА ДИСЦИПЛИНЫ", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                               space_after=6)
    add_paragraph_with_spacing(doc, subj_info["name"], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)

    add_paragraph_with_spacing(doc,
                               f"направление (специальность): {metadata.get('direction_code')} {metadata.get('direction_name')}",
                               align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_spacing(doc, f"направленность (профиль/программа/специализация): {metadata.get('profile')}",
                               align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_spacing(doc, f"уровень образования: {metadata.get('qualification')}",
                               align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_spacing(doc, f"форма обучения: {metadata.get('education_form')}",
                               align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_spacing(doc,
                               f"общая трудоемкость дисциплины составляет: {subj_info.get('credit_units')} зачетных единиц(ы)",
                               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=72)

    add_paragraph_with_spacing(doc, f"Ижевск {start_year}", align=WD_ALIGN_PARAGRAPH.CENTER)


def generate_compilers_page(doc, metadata: dict, subj_info: dict, staff: dict):
    """Страница 2: Составители и согласование."""
    doc.add_page_break()
    start_year = metadata.get("start_year") or "2026"
    add_paragraph_with_spacing(doc, f"Кафедра «{subj_info.get('department_name', '')}»", bold=True, space_after=18)

    if len(staff["compilers"]) > 1:
        add_paragraph_with_spacing(doc, "Составители:", bold=True, space_after=4)
        for compiler in staff["compilers"]:
            add_paragraph_with_spacing(doc, f"— {compiler}", space_before=2, space_after=2)
    else:
        comp_name = staff["compilers"][0] if staff["compilers"] else "Преподаватель кафедры"
        add_paragraph_with_spacing(doc, f"Составитель: {comp_name}", space_after=12)

    add_paragraph_with_spacing(doc, space_after=12)
    add_paragraph_with_spacing(doc,
                               f"Рабочая программа составлена в соответствии с требованиями образовательного стандарта ФГОС {metadata.get('fgos_standard')}, "
                               f"рассмотрена и одобрена на заседании кафедры.")

    add_paragraph_with_spacing(doc, f"Протокол от «____» ________________ {start_year} г. №_______", space_after=24)
    add_paragraph_with_spacing(doc, f"И.о. заведующего кафедрой __________________ {staff['head_of_department']}",
                               space_after=24)

    add_paragraph_with_spacing(doc, "СОГЛАСОВАНО", bold=True, space_after=12)
    add_paragraph_with_spacing(doc,
                               f"Количество часов рабочей программы и формируемые компетенции соответствуют учебному плану "
                               f"{metadata.get('direction_code')} «{metadata.get('direction_name')}» "
                               f"(профиль «{metadata.get('profile')}»)", space_after=18)

    # Динамическое определение УГСН
    dir_code = metadata.get("direction_code") or "01.03.04"
    faculty_name = metadata.get("faculty") or ""
    ugsn_code, ugsn_name = get_ugsn_info(dir_code, faculty_name)

    add_paragraph_with_spacing(doc, f"Протокол заседания учебно-методической комиссии по УГСН\n"
                                    f"{ugsn_code} «{ugsn_name}» от «____» _______________ {start_year} г. №_______",
                               space_after=24)

    add_paragraph_with_spacing(doc, f"Председатель учебно-методической комиссии по УГСН\n"
                                    f"{ugsn_code} «{ugsn_name}» _______________________ {staff['umk_chairman']}",
                               space_after=24)

    add_paragraph_with_spacing(doc,
                               f"Руководитель образовательной программы _______________________ {staff['program_director']}")


def generate_annotation_page(doc, metadata: dict, subj_info: dict, subj_ai: dict, mapped_comp: dict,
                             comp_registry: dict):
    """Страница 3: Аннотация дисциплины."""
    doc.add_page_break()
    add_paragraph_with_spacing(doc, "Аннотация к дисциплине", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                               space_after=12)

    table_annot = doc.add_table(rows=9, cols=2)
    table_annot.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_annot.style = "Table Grid"

    total_hours_sum = subj_info.get("total_hours", {}).get("total", 0)

    # Сбор компетенций
    comp_texts_list = []
    for c_code in mapped_comp.keys():
        c_desc = comp_registry.get(c_code, {}).get('competency_text', '')
        comp_texts_list.append(f"{c_code}. {c_desc}")
    competencies_val = "\n".join(comp_texts_list)

    # Сбор разделов
    sections_list = []
    for s in subj_ai["thematic_plan"]["sections"]:
        sections_list.append(f"Раздел {s['number']}. {s['name']}")
    sections_val = "\n".join(sections_list)

    # Сбор форм промежуточной аттестации
    control_forms_str_list = []
    if subj_info["control_forms"]["exams"]:
        control_forms_str_list.append(f"Экзамен ({', '.join(map(str, subj_info['control_forms']['exams']))} семестр)")
    if subj_info["control_forms"]["graded_credits"]:
        control_forms_str_list.append(
            f"Зачет с оценкой ({', '.join(map(str, subj_info['control_forms']['graded_credits']))} семестр)")
    if subj_info["control_forms"]["credits"]:
        control_forms_str_list.append(f"Зачет ({', '.join(map(str, subj_info['control_forms']['credits']))} семестр)")
    control_val = ", ".join(control_forms_str_list)

    annot_data = [
        ("Название дисциплины", subj_info["name"]),
        ("Направление подготовки (специальность)",
         f"{metadata.get('direction_code')} {metadata.get('direction_name')}"),
        ("Направленность (профиль/программа/специализация)", metadata.get('profile', '')),
        ("Место дисциплины",
         f"{subj_info.get('structure', {}).get('part', 'Обязательная часть')} Блока 1. Дисциплины (модули)"),
        ("Трудоемкость (з.е. / часы)", f"{subj_info.get('credit_units')} з.е. / {total_hours_sum} часов"),
        ("Цель изучения дисциплины", subj_ai["pedagogical_frame"]["goals"]),
        ("Компетенции, формируемые в результате освоения дисциплины", competencies_val),
        ("Содержание дисциплины (основные разделы и темы)", sections_val),
        ("Форма промежуточной аттестации", control_val)
    ]

    for i, (label, val) in enumerate(annot_data):
        row = table_annot.rows[i]
        set_row_cant_split(row)
        set_cell_text(row.cells[0], label, bold=True, size_pt=11)
        set_cell_text(row.cells[1], val, size_pt=11)
        set_cell_width(row.cells[0], 5.0)
        set_cell_width(row.cells[1], 11.5)


def generate_sections_1_2(doc, subj_ai: dict, mapped_comp: dict, comp_registry: dict):
    """Страница 4: Разделы 1 и 2 РПД (Цели, Задачи, Компетенции) со слиянием ячеек."""
    doc.add_page_break()
    add_paragraph_with_spacing(doc, "1.  Цели и задачи дисциплины:", bold=True, space_after=12)
    add_paragraph_with_spacing(doc, f"Целью освоения дисциплины является: {subj_ai['pedagogical_frame']['goals']}")
    add_paragraph_with_spacing(doc, "Задачи дисциплины:", bold=True)
    for task in subj_ai['pedagogical_frame']['tasks']:
        add_paragraph_with_spacing(doc, f"— {task}")

    add_paragraph_with_spacing(doc, "2. Планируемые результаты обучения", bold=True, space_after=12)
    add_paragraph_with_spacing(doc, "В результате освоения дисциплины у студента должны быть сформированы:")

    # Списки Знаний, Умений и Навыков
    add_paragraph_with_spacing(doc, "Знания, приобретаемые в ходе освоения дисциплины:", bold=True)
    knowledge_all = []
    for entry in subj_ai["pedagogical_frame"].get("indicators_ksa", []):
        knowledge_all.extend(entry["knowledge"])
    for idx_k, k in enumerate(list(set(knowledge_all))[:3], start=1):
        add_paragraph_with_spacing(doc, f"{idx_k}. {k}")

    add_paragraph_with_spacing(doc, "Умения, приобретаемые в ходе освоения дисциплины:", bold=True)
    skills_all = []
    for entry in subj_ai["pedagogical_frame"].get("indicators_ksa", []):
        skills_all.extend(entry["skills"])
    for idx_s, s in enumerate(list(set(skills_all))[:3], start=1):
        add_paragraph_with_spacing(doc, f"{idx_s}. {s}")

    add_paragraph_with_spacing(doc, "Навыки, приобретаемые в ходе освоения дисциплины:", bold=True)
    abilities_all = []
    for entry in subj_ai["pedagogical_frame"].get("indicators_ksa", []):
        abilities_all.extend(entry["abilities"])
    for idx_a, a in enumerate(list(set(abilities_all))[:3], start=1):
        add_paragraph_with_spacing(doc, f"{idx_a}. {a}")

    # Таблица распределения ЗУН по индикаторам
    add_paragraph_with_spacing(doc, "Компетенции, приобретаемые в ходе освоения дисциплины:", bold=True)
    table_comp = doc.add_table(rows=1, cols=5)
    table_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_comp.style = "Table Grid"
    set_repeat_table_header(table_comp.rows[0])
    set_row_cant_split(table_comp.rows[0])

    comp_headers = ["Компетенции", "Индикаторы", "Знания", "Умения", "Навыки"]
    for c_idx, h_text in enumerate(comp_headers):
        set_cell_text(table_comp.rows[0].cells[c_idx], h_text, bold=True, size_pt=11, align=WD_ALIGN_PARAGRAPH.CENTER,
                      fill_hex="F2F2F2")

    for c_code, ind_list in mapped_comp.items():
        comp_text = f"{c_code}. {comp_registry.get(c_code, {}).get('competency_text', '')}"
        for i_code in ind_list:
            row_cells = table_comp.add_row().cells
            set_row_cant_split(table_comp.rows[-1])
            ind_text = f"{i_code}. {comp_registry.get(c_code, {}).get('indicators', {}).get(i_code, {}).get('indicator_text', '')}"

            # Извлечение актуальных обозначений ЗУН по данному индикатору
            ksa_entry = next((item for item in subj_ai["pedagogical_frame"].get("indicators_ksa", []) if
                              item["indicator_code"] == i_code), None)
            k_labels = ", ".join(
                f"З{idx}" for idx in range(1, len(ksa_entry.get("knowledge", [])) + 1)) if ksa_entry else ""
            s_labels = ", ".join(
                f"У{idx}" for idx in range(1, len(ksa_entry.get("skills", [])) + 1)) if ksa_entry else ""
            a_labels = ", ".join(
                f"Н{idx}" for idx in range(1, len(ksa_entry.get("abilities", [])) + 1)) if ksa_entry else ""

            set_cell_text(row_cells[0], comp_text, size_pt=10)
            set_cell_text(row_cells[1], ind_text, size_pt=10)
            set_cell_text(row_cells[2], k_labels or "–", size_pt=10, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_text(row_cells[3], s_labels or "–", size_pt=10, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_text(row_cells[4], a_labels or "–", size_pt=10, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Применение вертикального объединения дубликатов компетенций в первом столбце
    merge_cells_vertically(table_comp, 0)


def generate_section_3(doc, subj_info: dict, subj_ai: dict):
    """Страница 5: Раздел 3 Место дисциплины в структуре ООП с расчетом курсов."""
    add_paragraph_with_spacing(doc)
    add_paragraph_with_spacing(doc, "3. Место дисциплины в структуре ООП", bold=True, space_after=12)
    add_paragraph_with_spacing(doc,
                               f"Дисциплина относится к обязательной части Блока 1 «Дисциплины (модули)» ООП.")

    # Динамический расчет курсов обучения на основе семестров
    sems = sorted([int(s) for s in subj_info.get("load_by_semester", {}).keys()])
    courses = sorted(list(set([(s + 1) // 2 for s in sems])))
    sems_str = ", ".join(map(str, sems))
    courses_str = ", ".join(map(str, courses))

    add_paragraph_with_spacing(doc, f"Дисциплина изучается на {courses_str} курсе(ах) в {sems_str} семестре(ах).")

    add_paragraph_with_spacing(doc,
                               f"Изучение дисциплины базируется на знаниях, умениях и навыках, полученных при изучении школьной программы.")
    add_paragraph_with_spacing(doc, f"Последующие дисциплины: {subj_ai['pedagogical_frame']['postrequisites_text']}")


def generate_fos_appendix(doc, metadata: dict, subj_info: dict, subj_ai: dict, mapped_comp: dict, comp_registry: dict,
                          sems_active: list):
    """Генерация приложения оценочных средств (ФОС) с динамическим получением тестов и вопросов из JSON."""
    doc.add_page_break()
    start_year = metadata.get("start_year") or "2026"
    add_paragraph_with_spacing(doc, "Приложение к рабочей программе дисциплины", italic=True,
                               align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=24)

    add_paragraph_with_spacing(doc, "МИНОБРНАУКИ РОССИИ", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_spacing(doc,
                               "Федеральное государственное бюджетное образовательное учреждение высшего образования",
                               align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_spacing(doc, "«Ижевский государственный технический университет имени М.Т. Калашникова»",
                               bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=36)

    add_paragraph_with_spacing(doc, "Оценочные средства по дисциплине", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                               space_after=6)
    add_paragraph_with_spacing(doc, subj_info["name"], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)

    add_paragraph_with_spacing(doc,
                               f"направление (специальность) {metadata.get('direction_code')} {metadata.get('direction_name')}",
                               align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_spacing(doc, f"направленность (профиль) {metadata.get('profile')}",
                               align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_spacing(doc, f"уровень образования: {metadata.get('qualification')}",
                               align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_spacing(doc, f"форма обучения: {metadata.get('education_form')}",
                               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)

    # Принудительный перенос первой таблицы оценочных средств на новую чистую страницу
    doc.add_page_break()
    add_paragraph_with_spacing(doc, "1. Оценочные средства", bold=True, space_after=12)
    table_fos_map = doc.add_table(rows=1, cols=4)
    table_fos_map.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_fos_map.style = "Table Grid"
    set_repeat_table_header(table_fos_map.rows[0])

    fos_map_headers = ["№ п/п", "Коды компетенций и индикаторов", "Результат обучения (знания, умения и навыки)",
                       "Формы текущего и промежуточного контроля"]
    for idx_f, f_text in enumerate(fos_map_headers):
        set_cell_text(table_fos_map.rows[0].cells[idx_f], f_text, bold=True, size_pt=10,
                      align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="F2F2F2")

    idx_row = 1
    for c_code, ind_list in mapped_comp.items():
        for i_code in ind_list:
            r_cells = table_fos_map.add_row().cells
            set_row_cant_split(table_fos_map.rows[-1])

            set_cell_text(r_cells[0], str(idx_row), align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=10)
            set_cell_text(r_cells[1], i_code, align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=10)

            # Сбор ЗУН
            ksa_strings = []
            for ksa_entry in subj_ai["pedagogical_frame"].get("indicators_ksa", []):
                if ksa_entry["indicator_code"] == i_code:
                    for idx_k, k_text in enumerate(ksa_entry["knowledge"], start=1):
                        ksa_strings.append(f"З{idx_k}: {k_text}")
                    for idx_s, s_text in enumerate(ksa_entry["skills"], start=1):
                        ksa_strings.append(f"У{idx_s}: {s_text}")
                    for idx_a, a_text in enumerate(ksa_entry["abilities"], start=1):
                        ksa_strings.append(f"Н{idx_a}: {a_text}")

            total_hours_dict = subj_info.get("total_hours", {})
            labs_h = total_hours_dict.get("laboratory_works", 0)
            practicals_h = total_hours_dict.get("practical_classes", 0)

            controls = []
            if labs_h > 0:
                controls.append("Защита лабораторных работ")
            if practicals_h > 0:
                controls.append("Защита практических работ")
            control_text = "\n".join(controls) if controls else "Устный опрос"

            set_cell_text(r_cells[2], "\n".join(ksa_strings), size_pt=9)
            set_cell_text(r_cells[3], control_text, size_pt=10)
            idx_row += 1

    # Вертикальное объединение по первому столбцу кодов индикаторов в ФОС
    merge_cells_vertically(table_fos_map, 1)

    # Тестирование с ключами к тестам
    add_paragraph_with_spacing(doc)
    add_paragraph_with_spacing(doc, "Наименование: проверочный тест", bold=True)
    add_paragraph_with_spacing(doc, "Представление в ФОС: набор вопросов для проведения тестирования", italic=True)

    # Динамический сбор тестовых вопросов из JSON-базы
    test_questions = subj_ai.get("resources_and_evaluation", {}).get("test_questions", [])

    # Резервный универсальный вариант на случай отсутствия тестов в исходном JSON
    if not test_questions:
        test_questions = [
            {
                "question": f"Что является главным объектом изучения дисциплины «{subj_info['name']}»?",
                "options": [
                    "Теоретические концепции и базовые определения дисциплины",
                    "Вспомогательные инструменты сторонних предметных областей",
                    "Программные методы без привязки к теоретическому базису",
                    "Организационно-правовые аспекты деятельности сторонних ведомств"
                ],
                "correct_answer": "1"
            },
            {
                "question": f"Какой метод наиболее часто применяется в рамках «{subj_info['name']}»?",
                "options": [
                    "Метод субъективных экспертных допущений",
                    "Системный анализ и комплексное моделирование процессов",
                    "Случайный подбор экспериментальных параметров",
                    "Интуитивный подход к проектированию систем"
                ],
                "correct_answer": "2"
            },
            {
                "question": "Что представляет собой методология предметной области?",
                "options": [
                    "Хаотичный набор практических рекомендаций",
                    "Система принципов и способов организации теоретической и практической деятельности",
                    "Второстепенный раздел учебной программы",
                    "Субъективное видение процесса разработки"
                ],
                "correct_answer": "2"
            },
            {
                "question": f"Какой результат освоения дисциплины «{subj_info['name']}» является приоритетным?",
                "options": [
                    "Отказ от использования современных инструментальных средств",
                    "Способность эффективно решать профессиональные задачи на основе полученных знаний и умений",
                    "Изучение только теоретических аспектов без практической применимости",
                    "Ориентация исключительно на исторический опыт ведения разработок"
                ],
                "correct_answer": "2"
            },
            {
                "question": "Какие требования предъявляются к уровню освоения учебного материала?",
                "options": [
                    "Поверхностное ознакомление без закрепления практических навыков",
                    "Отказ от самостоятельного выполнения разделов программы",
                    "Формирование компетенций, установленных государственным образовательным стандартом",
                    "Оценка успеваемости на основе случайного распределения баллов"
                ],
                "correct_answer": "3"
            }
        ]

    # Вывод вопросов теста на страницу
    for idx_q, q_item in enumerate(test_questions, start=1):
        add_paragraph_with_spacing(doc, f"{idx_q}. {q_item['question']}")
        for opt_idx, option in enumerate(q_item["options"], start=1):
            add_paragraph_with_spacing(doc, f"   {opt_idx}) {option}", space_before=2, space_after=2)

    # Динамический вывод таблицы Ключей тестов в соответствии с фактическим количеством вопросов
    add_paragraph_with_spacing(doc, "Ключи теста:", bold=True, space_before=12, space_after=6)
    num_cols = len(test_questions) + 1
    table_keys = doc.add_table(rows=2, cols=num_cols)
    table_keys.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_keys.style = "Table Grid"

    set_cell_text(table_keys.rows[0].cells[0], "Вопрос", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="F2F2F2")
    set_cell_text(table_keys.rows[1].cells[0], "Ответ", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="F2F2F2")

    for idx_tk, q_item in enumerate(test_questions, start=1):
        set_cell_text(table_keys.rows[0].cells[idx_tk], str(idx_tk), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(table_keys.rows[1].cells[idx_tk], str(q_item["correct_answer"]), align=WD_ALIGN_PARAGRAPH.CENTER)

    # Экзаменационный билет
    add_paragraph_with_spacing(doc, space_before=18)
    add_paragraph_with_spacing(doc, "Пример экзаменационного билета:", bold=True, space_after=12)

    table_ticket = doc.add_table(rows=1, cols=1)
    table_ticket.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_ticket.style = "Table Grid"
    set_row_cant_split(table_ticket.rows[0])
    cell_ticket = table_ticket.rows[0].cells[0]
    cell_ticket.width = Cm(15.0)

    # Оформление рамки экзаменационного билета
    p_ticket = cell_ticket.paragraphs[0]
    p_ticket.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_h1 = p_ticket.add_run("ФГБОУ ВО «Ижевский государственный технический университет имени М.Т. Калашникова»\n\n")
    set_font(run_h1, size_pt=9)

    run_title = p_ticket.add_run("ЭКЗАМЕНАЦИОННЫЙ БИЛЕТ № 1\n")
    set_font(run_title, size_pt=12, bold=True)

    run_sub = p_ticket.add_run(
        f"по дисциплине «{subj_info['name']}»\nдля направления {metadata.get('direction_code')} «{metadata.get('direction_name')}»\n\n")
    set_font(run_sub, size_pt=10, italic=True)

    # Извлечение контрольных вопросов
    questions_ticket = subj_ai.get("resources_and_evaluation", {}).get("control_questions", [])[:3]

    # Универсальный непогрешимый fallback на случай, если вопросов в базе недостаточно
    if len(questions_ticket) < 3:
        questions_ticket = [
            f"Теоретические основы и терминологический аппарат дисциплины «{subj_info['name']}».",
            f"Анализ ключевых концепций, алгоритмов и методов, изученных в рамках курса «{subj_info['name']}».",
            f"Практическое задание на применение методов и подходов дисциплины «{subj_info['name']}» для решения прикладной задачи."
        ]

    for idx_q, q_text in enumerate(questions_ticket, start=1):
        run_q = p_ticket.add_run(f"{idx_q}. {q_text}\n")
        set_font(run_q, size_pt=11)

    # Извлечение аббревиатуры кафедры
    dept_name = subj_info.get("department_name") or metadata.get(
        "department") or "Прикладная математика и информационные технологии"
    dept_abbr = get_department_acronym(dept_name)

    # Исключение захардкоденной даты (заменено на стандартное поле заполнения по ГОСТ)
    run_footer = p_ticket.add_run(
        f"\nБилет рассмотрен на заседании кафедры {dept_abbr} от «____» _______________ {start_year} г.")
    set_font(run_footer, size_pt=9, italic=True)