import os
import json
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")


def set_font(run, font_name="Times New Roman", size_pt=14, bold=False, italic=False, color_rgb=(0, 0, 0)):
    """Устанавливает шрифт, размер и начертание для текстового прогона."""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)

    try:
        rPr = run._r.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:cs'), font_name)
        rFonts.set(qn('w:eastAsia'), font_name)
        rPr.append(rFonts)
    except Exception as e:
        logger.debug(f"Не удалось применить стили шрифтов: {e}")


def add_paragraph_with_spacing(doc, text="", style="Normal", bold=False, italic=False, align=None, space_after=0,
                               space_before=0) -> Any:
    """Добавляет абзац с фиксированными отступами и одинарным интервалом."""
    p = doc.add_paragraph()
    try:
        p.style = style
    except Exception:
        pass
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.0
    if align:
        p.alignment = align
    if text:
        run = p.add_run(text)
        set_font(run, bold=bold, italic=italic)
    return p


def set_repeat_table_header(row):
    """Повторение шапки таблицы на каждой странице."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:tblHeader'))


def set_row_cant_split(row):
    """Запрет разрыва строки таблицы."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))


def set_cell_text(cell, text: str, bold=False, italic=False, size_pt=12, align=WD_ALIGN_PARAGRAPH.LEFT):
    """Форматирует ячейку таблицы и записывает в нее текст."""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    if text:
        run = p.add_run(text)
        set_font(run, size_pt=size_pt, bold=bold, italic=italic)


class RPGenerator:
    """Автоматический генератор рабочих программ дисциплин (РПД)."""

    def __init__(self, project_dir: Path, output_dir: Path):
        self.project_dir = project_dir
        self.output_dir = output_dir

        self.workload_path = project_dir / "rp_academic_workload.json"
        self.comp_map_path = project_dir / "rp_subject_competency_map.json"
        self.personnel_path = project_dir / "rp_personnel_mapping.json"
        self.ai_data_path = project_dir / "rp_ai_generated_data.json"

    def _load_json(self, path: Path) -> Dict[str, Any]:
        if not path.exists():
            raise FileNotFoundError(f"Файл не найден: {path.name}")
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    def _lookup_personnel(self, code: str, name: str, personnel_data: dict, metadata: dict) -> Dict[str, str]:
        """Интеллектуальный поиск кадрового состава по направлению, дисциплине или дефолту."""
        default_staff = personnel_data.get("default_department_personnel", {})

        # Получаем канонические ФИО по справочникам
        deans = personnel_data.get("deans", {})
        heads = personnel_data.get("heads_of_department", {})
        pds = personnel_data.get("program_directors", {})
        umks = personnel_data.get("umk_chairmen", {})
        oamrs = personnel_data.get("oamr_heads", {})
        teachers = personnel_data.get("teachers", {})

        dir_key = f"{metadata.get('direction_code')} {metadata.get('direction_name')}".strip()
        subjects_map = personnel_data.get("subjects_mapping", {}).get(dir_key, {})
        subj_staff = subjects_map.get(name, {})

        # Вспомогательный поиск по словарю
        def get_name(id_val, source_dict):
            return source_dict.get(id_val, {}).get("name", "") if id_val else ""

        resolved = {
            "dean": get_name(subj_staff.get("dean") or default_staff.get("dean"), deans),
            "head_of_department": get_name(
                subj_staff.get("head_of_department") or default_staff.get("head_of_department"), heads),
            "program_director": get_name(subj_staff.get("program_director") or default_staff.get("program_director"),
                                         pds),
            "umk_chairman": get_name(default_staff.get("umk_chairman"), umks),
            "oamr_head": get_name(default_staff.get("oamr_head"), oamrs),
            "compilers": []
        }

        # Сбор составителей
        compiler_ids = subj_staff.get("teachers", [])
        for tid in compiler_ids:
            t_info = teachers.get(tid, {})
            if t_info:
                resolved["compilers"].append(f"{t_info.get('name')}, {t_info.get('degree_and_title')}")

        if not resolved["compilers"]:
            resolved["compilers"].append("Преподаватель кафедры")

        return resolved

    def generate_all(self):
        # Загрузка баз данных
        workload = self._load_json(self.workload_path)
        comp_map = self._load_json(self.comp_map_path)
        personnel = self._load_json(self.personnel_path)
        ai_data = self._load_json(self.ai_data_path)

        metadata = workload.get("metadata", {})
        disciplines = workload.get("disciplines", {})
        comp_registry = comp_map.get("competencies_registry", {})
        subject_to_comp = comp_map.get("subject_to_competencies", {})

        self.output_dir.mkdir(parents=True, exist_ok=True)

        for idx, (code, subj_info) in enumerate(disciplines.items(), start=1):
            subj_name = subj_info["name"]

            # Проверяем наличие ИИ данных
            if code not in ai_data:
                logger.warning(f"Пропуск {code} {subj_name} — отсутствуют сгенерированные ИИ-данные.")
                continue

            logger.info(f"Генерация РПД [{idx}/{len(disciplines)}]: {code} {subj_name}...")

            subj_ai = ai_data[code]
            staff = self._lookup_personnel(code, subj_name, personnel, metadata)

            total_hours_dict = subj_info.get("total_hours", {})
            lectures_h = total_hours_dict.get("lectures", 0)
            practicals_h = total_hours_dict.get("practical_classes", 0)
            labs_h = total_hours_dict.get("laboratory_works", 0)
            # =======================================================

            doc = Document()

            # Настройка полей по умолчанию
            section = doc.sections[0]
            section.top_margin = Cm(2.0)
            section.bottom_margin = Cm(2.0)
            section.left_margin = Cm(3.0)
            section.right_margin = Cm(1.5)

            # === СТРАНИЦА 1: ТИТУЛЬНЫЙ ЛИСТ ===
            add_paragraph_with_spacing(doc, "МИНОБРНАУКИ РОССИИ", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            add_paragraph_with_spacing(doc,
                                       "Федеральное государственное бюджетное образовательное учреждение высшего образования",
                                       align=WD_ALIGN_PARAGRAPH.CENTER)
            add_paragraph_with_spacing(doc,
                                       f"«Ижевский государственный технический университет имени М.Т. Калашникова»",
                                       bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

            add_paragraph_with_spacing(doc)

            p_utv = add_paragraph_with_spacing(doc, "УТВЕРЖДАЮ", bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
            add_paragraph_with_spacing(doc, f"Декан/Директор\n_____________/ {staff['dean']}",
                                       align=WD_ALIGN_PARAGRAPH.RIGHT)
            add_paragraph_with_spacing(doc, f"_________________ {metadata.get('start_year', '20')} г.",
                                       align=WD_ALIGN_PARAGRAPH.RIGHT)

            for _ in range(4):
                add_paragraph_with_spacing(doc)

            add_paragraph_with_spacing(doc, "РАБОЧАЯ ПРОГРАММА ДИСЦИПЛИНЫ", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            add_paragraph_with_spacing(doc, subj_name, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

            add_paragraph_with_spacing(doc)

            add_paragraph_with_spacing(doc,
                                       f"направление (специальность) {metadata.get('direction_code')} {metadata.get('direction_name')}",
                                       align=WD_ALIGN_PARAGRAPH.CENTER)
            add_paragraph_with_spacing(doc, f"направленность (профиль) {metadata.get('profile')}",
                                       align=WD_ALIGN_PARAGRAPH.CENTER)
            add_paragraph_with_spacing(doc, f"уровень образования: {metadata.get('qualification')}",
                                       align=WD_ALIGN_PARAGRAPH.CENTER)
            add_paragraph_with_spacing(doc, f"форма обучения: {metadata.get('education_form')}",
                                       align=WD_ALIGN_PARAGRAPH.CENTER)
            add_paragraph_with_spacing(doc,
                                       f"общая трудоемкость дисциплины составляет: {subj_info.get('credit_units')} з.е.",
                                       align=WD_ALIGN_PARAGRAPH.CENTER)

            for _ in range(3):
                add_paragraph_with_spacing(doc)

            add_paragraph_with_spacing(doc, f"Ижевск {metadata.get('start_year', '')}", align=WD_ALIGN_PARAGRAPH.CENTER)

            # === СТРАНИЦА 2: СОСТАВИТЕЛИ И СОГЛАСОВАНИЕ ===
            doc.add_page_break()
            add_paragraph_with_spacing(doc, f"Кафедра {subj_info.get('department_name')}", bold=True)

            add_paragraph_with_spacing(doc)

            for comp in staff["compilers"]:
                add_paragraph_with_spacing(doc, f"Составитель: {comp}")

            add_paragraph_with_spacing(doc)

            add_paragraph_with_spacing(doc,
                                       f"Рабочая программа составлена в соответствии с требованиями образовательного стандарта ФГОС {metadata.get('fgos_standard')}")

            add_paragraph_with_spacing(doc)
            add_paragraph_with_spacing(doc, f"Заведующий кафедрой __________________ {staff['head_of_department']}")
            add_paragraph_with_spacing(doc, f"Председатель УМК __________________ {staff['umk_chairman']}")
            add_paragraph_with_spacing(doc,
                                       f"Руководитель образовательной программы __________________ {staff['program_director']}")

            # === СЕКЦИЯ 1: ЦЕЛИ И ЗАДАЧИ ===
            doc.add_page_break()
            add_paragraph_with_spacing(doc, "1. Цели и задачи дисциплины:", bold=True)
            add_paragraph_with_spacing(doc,
                                       f"Целью освоения дисциплины является: {subj_ai['pedagogical_frame']['goals']}")
            add_paragraph_with_spacing(doc, "Задачи дисциплины:", bold=True)
            for task in subj_ai['pedagogical_frame']['tasks']:
                add_paragraph_with_spacing(doc, f"— {task}")

            # === СЕКЦИЯ 2: РЕЗУЛЬТАТЫ (ЗУН) ===
            add_paragraph_with_spacing(doc)
            add_paragraph_with_spacing(doc, "2. Планируемые результаты обучения:", bold=True)

            # Таблица компетенций и индикаторов (ЗУН)
            table_comp = doc.add_table(rows=1, cols=5)
            table_comp.style = "Table Grid"
            hdr_cells = table_comp.rows[0].cells
            headers = ["Код компетенции", "Индикаторы", "Знания (З)", "Умения (У)", "Навыки (Н)"]
            for idx, text in enumerate(headers):
                set_cell_text(hdr_cells[idx], text, bold=True, size_pt=12, align=WD_ALIGN_PARAGRAPH.CENTER)

            mapped_comp = subject_to_comp.get(code, {}).get("competencies", {})
            for c_code, ind_list in mapped_comp.items():
                row_cells = table_comp.add_row().cells
                comp_text = f"{c_code}: {comp_registry.get(c_code, {}).get('competency_text', '')}"
                set_cell_text(row_cells[0], comp_text, size_pt=11)

                inds_text = []
                for i_code in ind_list:
                    inds_text.append(
                        f"{i_code}: {comp_registry.get(c_code, {}).get('indicators', {}).get(i_code, {}).get('indicator_text', '')}")
                set_cell_text(row_cells[1], "\n".join(inds_text), size_pt=11)

                # Поиск соответствующих ЗУН по индикаторам из ИИ-данных
                k_list, s_list, a_list = [], [], []
                for entry in subj_ai["pedagogical_frame"].get("indicators_ksa", []):
                    if entry["indicator_code"] in ind_list:
                        k_list.extend(entry["knowledge"])
                        s_list.extend(entry["skills"])
                        a_list.extend(entry["abilities"])

                set_cell_text(row_cells[2], "\n".join([f"— {x}" for x in k_list]), size_pt=11)
                set_cell_text(row_cells[3], "\n".join([f"— {x}" for x in s_list]), size_pt=11)
                set_cell_text(row_cells[4], "\n".join([f"— {x}" for x in a_list]), size_pt=11)

            # === СЕКЦИЯ 3: МЕСТО В СТРУКТУРЕ ООП ===
            add_paragraph_with_spacing(doc)
            add_paragraph_with_spacing(doc, "3. Место дисциплины в структуре ООП:", bold=True)
            struct = subj_info.get("structure", {})
            add_paragraph_with_spacing(doc,
                                       f"Дисциплина относится к разделу: {struct.get('part', '')} ({struct.get('block', '')}).")
            add_paragraph_with_spacing(doc,
                                       f"Предшествующие дисциплины: {subj_ai['pedagogical_frame']['prerequisites_text']}")
            add_paragraph_with_spacing(doc,
                                       f"Последующие дисциплины: {subj_ai['pedagogical_frame']['postrequisites_text']}")

            # === СЕКЦИЯ 4: СТРУКТУРА И СОДЕРЖАНИЕ ===
            add_paragraph_with_spacing(doc)
            add_paragraph_with_spacing(doc, "4. Структура и содержание дисциплины:", bold=True)

            # Математическое распределение СРС по разделам
            thematic = subj_ai["thematic_plan"]
            sections = thematic.get("sections", [])

            section_hours = {}
            for s in sections:
                s_num = s["number"]
                section_hours[s_num] = {"lectures": 0, "practicals": 0, "labs": 0, "cpc": 0}

            for l in thematic.get("lectures", []):
                s_num = l["section_number"]
                if s_num in section_hours:
                    section_hours[s_num]["lectures"] += l["hours"]
            for p in thematic.get("practicals", []):
                s_num = p["section_number"]
                if s_num in section_hours:
                    section_hours[s_num]["practicals"] += p["hours"]
            for lb in thematic.get("labs", []):
                s_num = lb["section_number"]
                if s_num in section_hours:
                    section_hours[s_num]["labs"] += lb["hours"]

            # Распределяем СРС пропорционально контактным часам
            total_contact = lectures_h + practicals_h + labs_h
            total_cpc = subj_info["total_hours"].get("self_study", 0)

            for s_num, h in section_hours.items():
                contact = h["lectures"] + h["practicals"] + h["labs"]
                if total_contact > 0:
                    h["cpc"] = int(round((contact / total_contact) * total_cpc))
                else:
                    h["cpc"] = 0

            # 4.1 Таблица структуры дисциплины
            add_paragraph_with_spacing(doc, "4.1 Структура учебной нагрузки по разделам:", bold=True)
            table_struct = doc.add_table(rows=1, cols=7)
            table_struct.style = "Table Grid"
            hdr_str = table_struct.rows[0].cells
            headers_str = ["№", "Наименование раздела", "Всего ч.", "Лекции", "Практ.", "Лаб.", "СРС"]
            for idx, text in enumerate(headers_str):
                set_cell_text(hdr_str[idx], text, bold=True, size_pt=12, align=WD_ALIGN_PARAGRAPH.CENTER)

            for s in sections:
                row_cells = table_struct.add_row().cells
                s_num = s["number"]
                h = section_hours[s_num]
                total_s_hours = h["lectures"] + h["practicals"] + h["labs"] + h["cpc"]

                set_cell_text(row_cells[0], str(s_num), size_pt=11, align=WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_text(row_cells[1], s["name"], size_pt=11)
                set_cell_text(row_cells[2], str(total_s_hours), size_pt=11, align=WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_text(row_cells[3], str(h["lectures"]), size_pt=11, align=WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_text(row_cells[4], str(h["practicals"]), size_pt=11, align=WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_text(row_cells[5], str(h["labs"]), size_pt=11, align=WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_text(row_cells[6], str(h["cpc"]), size_pt=11, align=WD_ALIGN_PARAGRAPH.CENTER)

            # 4.2 Содержание разделов
            add_paragraph_with_spacing(doc)
            add_paragraph_with_spacing(doc, "4.2 Краткое содержание разделов курса:", bold=True)
            for s in sections:
                add_paragraph_with_spacing(doc, f"Раздел {s['number']}. {s['name']}", bold=True)
                add_paragraph_with_spacing(doc, s["description"])

            # 4.3 Темы лекций
            add_paragraph_with_spacing(doc)
            add_paragraph_with_spacing(doc, "4.3 Наименование тем лекционных занятий:", bold=True)
            for l in thematic.get("lectures", []):
                add_paragraph_with_spacing(doc, f"— Раздел {l['section_number']}. Тема: {l['theme']} ({l['hours']} ч.)")
                add_paragraph_with_spacing(doc, f"  Содержание: {l['content']}", italic=True)

            # 4.4 Темы практик
            if practicals_h > 0:
                add_paragraph_with_spacing(doc)
                add_paragraph_with_spacing(doc, "4.4 Наименование тем практических (семинарских) занятий:", bold=True)
                for p in thematic.get("practicals", []):
                    add_paragraph_with_spacing(doc,
                                               f"— Раздел {p['section_number']}. Семинар: {p['theme']} ({p['hours']} ч.)")

            # 4.5 Темы лабораторных
            if labs_h > 0:
                add_paragraph_with_spacing(doc)
                add_paragraph_with_spacing(doc, "4.5 Наименование тем лабораторных работ:", bold=True)
                for lb in thematic.get("labs", []):
                    add_paragraph_with_spacing(doc,
                                               f"— Раздел {lb['section_number']}. Лабораторная работа: {lb['theme']} ({lb['hours']} ч.)")

            # === СЕКЦИЯ 5: ОЦЕНОЧНЫЕ МАТЕРИАЛЫ ===
            add_paragraph_with_spacing(doc)
            add_paragraph_with_spacing(doc, "5. Оценочные материалы для аттестации:", bold=True)
            add_paragraph_with_spacing(doc, "Вопросы для подготовки к промежуточной аттестации (зачету/экзамену):",
                                       bold=True)
            for idx_q, q in enumerate(subj_ai["resources_and_evaluation"].get("control_questions", []), start=1):
                add_paragraph_with_spacing(doc, f"{idx_q}. {q}")

            # === СЕКЦИЯ 6: УЧЕБНО-МЕТОДИЧЕСКОЕ ОБЕСПЕЧЕНИЕ ===
            add_paragraph_with_spacing(doc)
            add_paragraph_with_spacing(doc, "6. Учебно-методическое и информационное обеспечение:", bold=True)

            add_paragraph_with_spacing(doc, "а) основная литература:", bold=True)
            for book in subj_ai["resources_and_evaluation"].get("primary_literature", []):
                add_paragraph_with_spacing(doc, f"— {book}")

            add_paragraph_with_spacing(doc, "б) дополнительная литература:", bold=True)
            for book in subj_ai["resources_and_evaluation"].get("secondary_literature", []):
                add_paragraph_with_spacing(doc, f"— {book}")

            add_paragraph_with_spacing(doc, "в) методические указания:", bold=True)
            for guide in subj_ai["resources_and_evaluation"].get("methodological_guidelines", []):
                add_paragraph_with_spacing(doc, f"— {guide}")

            add_paragraph_with_spacing(doc, "г) интернет-ресурсы:", bold=True)
            for link in subj_ai["resources_and_evaluation"].get("internet_resources", []):
                add_paragraph_with_spacing(doc, f"— {link}")

            add_paragraph_with_spacing(doc, "д) программное обеспечение:", bold=True)
            standard_sw = [
                "Операционная система семейства Microsoft Windows / Linux",
                "Пакет офисных приложений LibreOffice / МойОфис / Microsoft Office",
                "Браузер Яндекс.Браузер / Mozilla Firefox / Google Chrome",
                "Свободно распространяемая среда разработки по профилю дисциплины"
            ]
            for sw in standard_sw:
                add_paragraph_with_spacing(doc, f"— {sw}")

            # === СЕКЦИЯ 7: МАТЕРИАЛЬНО-ТЕХНИЧЕСКОЕ ОБЕСПЕЧЕНИЕ ===
            add_paragraph_with_spacing(doc)
            add_paragraph_with_spacing(doc, "7. Материально-техническое обеспечение дисциплины:", bold=True)
            add_paragraph_with_spacing(doc,
                                       "— Учебная аудитория для проведения лекционных занятий, укомплектованная специализированной мебелью и техническими средствами обучения (мультимедийный проектор, экран).")
            add_paragraph_with_spacing(doc,
                                       "— Лаборатория / компьютерный класс для проведения практических занятий и лабораторных работ, оснащенный ПЭВМ с доступом в Интернет и локальную сеть вуза.")

            # Сохранение готового файла DOCX
            safe_filename = "".join(c for c in f"{code} {subj_name}" if c.isalnum() or c in (" ", "_", "-")).strip()
            output_file_path = self.output_dir / f"{safe_filename}.docx"

            try:
                doc.save(str(output_file_path.absolute()))
                logger.info(f"  [+] Успешно сгенерирована РПД: {output_file_path.name}")
            except Exception as e:
                logger.error(f"Не удалось сохранить РПД {subj_name}: {e}")

        print(
            f"\n[Успешно] Генерация всех рабочих программ завершена. Файлы сохранены в:\n{self.output_dir.absolute()}")


def main():
    print("=== Комплексный генератор рабочих программ дисциплин (РПД) ===")

    # Считывание путей к ресурсам
    user_project_dir = input("Введите путь к папке с файлами JSON (по умолчанию 'services/rp_generator'): ").strip()
    if not user_project_dir:
        user_project_dir = "services/rp_generator"
    project_dir = Path(user_project_dir)

    user_output_dir = input(
        "Введите путь к папке для сохранения сгенерированных РПД (по умолчанию 'RPD_Output'): ").strip()
    if not user_output_dir:
        user_output_dir = "RPD_Output"
    output_dir = Path(user_output_dir)

    generator = RPGenerator(project_dir=project_dir, output_dir=output_dir)
    try:
        generator.generate_all()
    except Exception as e:
        logger.error(f"Ошибка при работе генератора РПД: {e}", exc_info=True)


if __name__ == "__main__":
    main()