import os
import re
import json
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional
from docx import Document
from openpyxl import load_workbook

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)
logger = logging.getLogger(__name__)

# === ГИБКИЕ РЕГУЛЯРНЫЕ ВЫРАЖЕНИЯ ===
ROLE_DEAN = re.compile(
    r"(?:(?:и\.?\s*о\.?|врио|временно\s+исполняющий\s+обязанности)\s+)?(декан|директор|руководитель\s+института)\b",
    re.IGNORECASE
)
ROLE_HEAD = re.compile(
    r"(?:(?:и\.?\s*о\.?|врио)\s+)?(?:заведующий|зав\.?)\s*(?:кафедрой|каф\.?)\b",
    re.IGNORECASE
)
ROLE_DIRECTOR = re.compile(
    r"(?:руководитель\s+(?:образовательной\s+)?программы|руководитель\s+оп|руководитель\s+направления)\b",
    re.IGNORECASE
)
ROLE_COMPILER = re.compile(
    r"(?:составител[ьяи]|разработчик[и]?)\b",
    re.IGNORECASE
)

# Шаблон поиска ученых степеней, званий и должностей
DEGREE_PATTERN = re.compile(
    r"(?:преподаватель|доцент|профессор|ассистент|к\.?\s*[тфмэ]\.?\s*н\.?|д\.?\s*[тфмэ]\.?\s*н\.?|ст\.?\s*преп)",
    re.IGNORECASE
)

# Стоп-фразы, прерывающие сканирование блока составителей РП
STOP_COMPILER_PHRASES = [
    "рабочая программа составлена",
    "в соответствии с требованиями",
    "рассмотрена на заседании",
    "протокол от",
    "согласовано",
    "кафедра",
    "заведующий",
    "декан"
]

# Шаблоны метаданных учебного плана
DIRECTION_ROLE = re.compile(r"(?:направление|специальность|подготовк\w*)\b", re.IGNORECASE)
CODE_PATTERN = re.compile(r"(\d{2}\.\d{2}\.\d{2})")
PROFILE_ROLE = re.compile(r"(?:направленность|профиль|программа/специализация|специализация)\b", re.IGNORECASE)
DISCIPLINE_ROLE = re.compile(
    r"(?:рабочая\s+программа\s+дисциплины|оценочные\s+средства\s+по\s+дисциплине)",
    re.IGNORECASE
)

# Шаблон поиска ФИО с инициалами
INITIALS_PATTERN = re.compile(
    r'(?:[А-Я]\s*\.\s*[А-Я]\s*\.\s*[А-Я][а-я]+|[А-Я][а-я]+\s+[А-Я]\s*\.\s*[А-Я]\s*\.)'
)


def clean_text(text: str) -> str:
    """Нормализует пробелы и удаляет артефакты форматирования подписей."""
    if not text:
        return ""
    s = re.sub(r'[_/\\|]+', ' ', text)
    s = " ".join(s.split()).strip()
    return s


def clean_fio_spaces(name_str: str) -> str:
    """Приводит ФИО с инициалами к единому стандартному виду."""
    s = clean_text(name_str)
    s = re.sub(r'([А-Я]\.)\s+([А-Я]\.)', r'\1\2', s)
    s = re.sub(r'([А-Я][а-я]+)([А-Я]\.[А-Я]\.)', r'\1 \2', s)
    s = re.sub(r'([А-Я]\.[А-Я]\.)([А-Я][а-я]+)', r'\1 \2', s)
    return s


def has_name_pattern(text: str) -> bool:
    """Проверяет, содержит ли строка упоминание имени или фамилии составителя."""
    if INITIALS_PATTERN.search(text):
        return True
    capitalized = re.findall(r'\b[А-Я][а-я]+\b', text)
    if len(capitalized) >= 2:
        return True
    return False


def split_compiler_name_and_degree(compiler_str: str) -> tuple[str, str]:
    """Разделяет ФИО преподавателя и его ученые степени/звания."""
    parts = compiler_str.split(",", 1)
    name = clean_text(parts[0])
    degree = clean_text(parts[1]) if len(parts) > 1 else ""
    degree = re.sub(r'^[,\s]+|[,\s]+$', '', degree)
    return name, degree


def levenshtein_distance(s1: str, s2: str) -> int:
    """Вычисляет редакционное расстояние Левенштейна между строками."""
    if len(s1) < len(s2):
        return levenshtein_distance(s2, s1)
    if len(s2) == 0:
        return len(s1)

    previous_row = range(len(s2) + 1)
    for i, c1 in enumerate(s1):
        current_row = [i + 1]
        for j, c2 in enumerate(s2):
            insertions = previous_row[j + 1] + 1
            deletions = current_row[j] + 1
            substitutions = previous_row[j] + (c1 != c2)
            current_row.append(min(insertions, deletions, substitutions))
        previous_row = current_row

    return previous_row[-1]


def is_similar_name(name1: str, name2: str) -> bool:
    """Нечетко сравнивает два ФИО на предмет опечаток или сокращений до инициалов."""
    n1 = name1.lower().replace("ё", "е").strip()
    n2 = name2.lower().replace("ё", "е").strip()

    if n1 == n2:
        return True

    # Разбиваем на слова для анализа фамилии и инициалов
    w1 = [w for w in re.split(r'[^а-яa-z]', n1) if w]
    w2 = [w for w in re.split(r'[^а-яa-z]', n2) if w]

    if not w1 or not w2:
        return False

    # Фамилия (обычно идет первым словом)
    surname1, surname2 = w1[0], w2[0]

    # Допускаем опечатку в 1 символ в фамилии (например, Нефедов/Нефёдов или окончание)
    if levenshtein_distance(surname1, surname2) <= 1:
        # Извлекаем первые буквы имени и отчества
        init1 = "".join([w[0] for w in w1[1:] if len(w) > 0])
        init2 = "".join([w[0] for w in w2[1:] if len(w) > 0])

        if init1 and init2:
            # Сопоставляем по минимальной длине инициалов (например, "дг" совпадет с "денисгеннадьевич")
            min_len = min(len(init1), len(init2))
            if init1[:min_len] == init2[:min_len]:
                return True
    return False


def find_person_by_role(lines: List[str], role_regex: re.Pattern, name_regex: re.Pattern, scan_offset: int = 4) -> str:
    """Ищет должностное лицо по гибкому шаблону роли в окрестности совпадения."""
    for idx, line in enumerate(lines):
        line_clean = clean_text(line)
        if role_regex.search(line_clean):
            for offset in range(0, scan_offset):
                if idx + offset < len(lines):
                    pot_line = clean_text(lines[idx + offset])
                    m = name_regex.search(pot_line)
                    if m:
                        return clean_fio_spaces(m.group(0))
    return ""


def extract_default_personnel_from_excel(sheet) -> Dict[str, str]:
    """Парсит лист 'Титул' и извлекает должностных лиц по умолчанию для выпускающей кафедры."""
    default_staff = {
        "dean": "",
        "head_of_department": "",
        "program_director": "",
        "umk_chairman": "",
        "oamr_head": ""
    }

    role_dean_pattern = re.compile(r"декан", re.IGNORECASE)
    role_head_pattern = re.compile(r"(?:зав\.\s*кафедрой|заведующий\s*кафедрой|зав\. кафедра)", re.IGNORECASE)
    role_director_pattern = re.compile(r"руководитель\s+(?:ооп|направления|ооп)\b", re.IGNORECASE)
    role_umk_pattern = re.compile(r"председатель\s+умк", re.IGNORECASE)
    role_oamr_pattern = re.compile(r"начальник\s+оамр", re.IGNORECASE)

    fio_excel_pattern = re.compile(
        r"/\s*([А-Яа-я]\s*\.\s*[А-Яа-я]\s*\.\s*[А-Яа-я]+|[А-Яа-я]+\s+[А-Яа-я]\s*\.\s*[А-Яа-я]\s*\.)\s*/"
    )
    fio_fallback_pattern = re.compile(
        r"([А-Яа-я]\s*\.\s*[А-Яа-я]\s*\.\s*[А-Яа-я]+|[А-Яа-я]+\s+[А-Яа-я]\s*\.\s*[А-Яа-я]\s*\.)"
    )

    for row in range(1, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            cell_val = str(sheet.cell(row=row, column=col).value or "").strip()
            if not cell_val:
                continue

            target_key = None
            if role_dean_pattern.search(cell_val):
                target_key = "dean"
            elif role_head_pattern.search(cell_val):
                target_key = "head_of_department"
            elif role_director_pattern.search(cell_val):
                target_key = "program_director"
            elif role_umk_pattern.search(cell_val):
                target_key = "umk_chairman"
            elif role_oamr_pattern.search(cell_val):
                target_key = "oamr_head"

            if target_key:
                for offset in range(1, 15):
                    if col + offset <= sheet.max_column:
                        test_cell = str(sheet.cell(row=row, column=col + offset).value or "").strip()
                        if test_cell:
                            m = fio_excel_pattern.search(test_cell)
                            if m:
                                default_staff[target_key] = clean_fio_spaces(m.group(1))
                                break
                            m_fb = fio_fallback_pattern.search(test_cell)
                            if m_fb:
                                default_staff[target_key] = clean_fio_spaces(m_fb.group(1))
                                break
    return default_staff


def parse_rp_file(file_path: Path) -> Optional[dict]:
    """Анализирует docx файл РП с использованием гибких регулярных выражений."""
    try:
        doc = Document(file_path)
    except Exception as e:
        logger.error(f"Не удалось открыть файл {file_path.name}: {e}")
        return None

    lines: List[str] = []
    for p in doc.paragraphs:
        val = p.text.strip()
        if val:
            lines.append(val)

    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                val = cell.text.strip()
                if val:
                    for line in val.split("\n"):
                        line_strip = line.strip()
                        if line_strip and line_strip not in lines:
                            lines.append(line_strip)

    # 1. Извлечение направления (код и наименование)
    direction_code = ""
    direction_name = ""
    for idx, line in enumerate(lines):
        line_clean = clean_text(line)
        if DIRECTION_ROLE.search(line_clean):
            m_code = CODE_PATTERN.search(line_clean)
            if m_code:
                direction_code = m_code.group(1)
                parts = line_clean.split(direction_code, 1)
                if len(parts) > 1:
                    pot_name = parts[1].strip()
                    pot_name = re.sub(r'(?:код|наименование|полностью).*', '', pot_name, flags=re.IGNORECASE).strip()
                    pot_name = re.sub(r'^[-\s,.:\)]+', '', pot_name).strip()
                    if pot_name:
                        direction_name = pot_name
                break

    # 2. Извлечение профиля/направленности
    profile = ""
    for idx, line in enumerate(lines):
        line_clean = clean_text(line)
        if PROFILE_ROLE.search(line_clean):
            parts = re.split(r'[:\)]', line_clean, 1)
            pot_profile = parts[-1].strip() if len(parts) > 1 else ""

            if not pot_profile or len(pot_profile) < 5 or any(
                    x in pot_profile.lower() for x in ["направленность", "профиль", "наименование"]):
                for offset in range(1, 3):
                    if idx + offset < len(lines):
                        pot = clean_text(lines[idx + offset])
                        if pot and len(pot) > 5 and not any(
                                x in pot.lower() for x in ["направленность", "профиль", "наименование"]):
                            pot_profile = pot
                            break

            if pot_profile:
                # Универсальное отсечение остатков левых скобок
                pot_profile = re.sub(r"^.*?\)\s*", "", pot_profile)
                pot_profile = re.sub(r'(?:наименование|полностью).*', '', pot_profile, flags=re.IGNORECASE).strip()
                pot_profile = re.sub(r'^[-\s,.:\)]+', '', pot_profile).strip()
                if pot_profile:
                    profile = pot_profile
                    break

    # 3. Извлечение названия дисциплины
    discipline = ""
    for idx, line in enumerate(lines):
        line_clean = clean_text(line)
        if DISCIPLINE_ROLE.search(line_clean):
            for offset in range(1, 4):
                if idx + offset < len(lines):
                    pot = clean_text(lines[idx + offset])
                    if pot and not any(x in pot.lower() for x in ["наименование", "направление", "код"]):
                        discipline = pot
                        break
            if discipline:
                break

    # 4. Извлечение должностных лиц
    dean = find_person_by_role(lines, ROLE_DEAN, INITIALS_PATTERN, scan_offset=4)
    head_of_department = find_person_by_role(lines, ROLE_HEAD, INITIALS_PATTERN, scan_offset=4)
    program_director = find_person_by_role(lines, ROLE_DIRECTOR, INITIALS_PATTERN, scan_offset=4)

    # 5. Извлечение списка составителей РП
    raw_compilers: List[str] = []
    for idx, line in enumerate(lines):
        line_clean = clean_text(line)
        if ROLE_COMPILER.search(line_clean):
            for offset in range(0, 7):
                if idx + offset < len(lines):
                    pot = clean_text(lines[idx + offset])

                    if offset > 0 and (ROLE_DEAN.search(pot) or ROLE_HEAD.search(pot) or ROLE_DIRECTOR.search(pot) or
                                       any(p in pot.lower() for p in STOP_COMPILER_PHRASES)):
                        break

                    if pot:
                        cleaned = ROLE_COMPILER.sub("", pot).strip()
                        cleaned = re.sub(r'^[-\s,.:]+', '', cleaned).strip()
                        cleaned = re.sub(r'\s*Ф\.И\.О\..*', '', cleaned, flags=re.IGNORECASE).strip()
                        if cleaned and not cleaned.startswith("_") and len(cleaned) > 4:
                            raw_compilers.append(cleaned)
            break

    compilers: List[str] = []
    for item in raw_compilers:
        if has_name_pattern(item):
            compilers.append(item)
        else:
            if compilers:
                compilers[-1] = (compilers[-1].rstrip(",") + ", " + item).strip()
                compilers[-1] = re.sub(r'\s*,\s*,', ',', compilers[-1])
                compilers[-1] = " ".join(compilers[-1].split())
            else:
                compilers.append(item)

    if not discipline:
        logger.warning(f"В файле {file_path.name} не удалось определить название дисциплины.")
        return None

    direction_full = f"{direction_code} {direction_name}".strip() if direction_code else "Неизвестное направление"

    return {
        "direction": direction_full,
        "profile": profile,
        "discipline": discipline,
        "data": {
            "dean": dean,
            "compilers": compilers,
            "head_of_department": head_of_department,
            "program_director": program_director,
            "profile": profile
        }
    }


def main():
    print("=== Комплексный экстрактор и маппер должностных лиц РП ===")

    user_excel = input("Шаг 1. Введите путь к файлу Excel учебного плана (например, plan.xlsx): ").strip()
    if not user_excel:
        user_excel = "plan.xlsx"

    excel_path = Path(user_excel)
    default_staff = {}
    if excel_path.exists():
        logger.info(f"Парсинг листа 'Титул' из Excel плана: {excel_path.name}...")
        try:
            wb = load_workbook(str(excel_path.absolute()), data_only=True)
            if "Титул" in wb.sheetnames:
                default_staff = extract_default_personnel_from_excel(wb["Титул"])
                logger.info(f"  [+] Извлечены должностные лица по умолчанию: {default_staff}")
            else:
                logger.warning("Лист 'Титул' не обнаружен в Excel.")
        except Exception as e:
            logger.error(f"Не удалось распарсить Excel: {e}")
    else:
        logger.warning("Файл Excel не найден. Слияние с должностями по умолчанию будет пропущено.")

    user_rp_folder = input("\nШаг 2. Введите путь к папке с архивными файлами РП (docx): ").strip()
    if not user_rp_folder:
        user_rp_folder = "."

    rp_folder = Path(user_rp_folder)
    subjects_mapping: Dict[str, Dict[str, Any]] = {}

    if rp_folder.exists():
        docx_files = list(rp_folder.glob("*.docx"))
        print(f"Найдено файлов для анализа: {len(docx_files)}")

        for docx_path in docx_files:
            if docx_path.name.startswith("~$"):
                continue

            logger.info(f"Анализ файла РП: {docx_path.name}")
            result = parse_rp_file(docx_path)
            if result:
                dir_key = result["direction"]
                disc_key = result["discipline"]

                if dir_key not in subjects_mapping:
                    subjects_mapping[dir_key] = {}

                extracted_data = result["data"]

                # Слияние с должностями по умолчанию из Excel
                for key in ["dean", "head_of_department", "program_director"]:
                    if not extracted_data.get(key) and default_staff.get(key):
                        extracted_data[key] = default_staff[key]
                        logger.info(
                            f"    [*] Поле '{key}' заполнено значением по умолчанию из Excel: {default_staff[key]}")

                subjects_mapping[dir_key][disc_key] = extracted_data
                logger.info(f"  [+] Успешно сопоставлены кадры для: {disc_key} ({dir_key})")
    else:
        logger.warning("Папка с Word РП не найдена или пропущена.")

    # === АЛГОРИТМ НЕЧЕТКОЙ КЛАСТЕРИЗАЦИИ И ЧАСТОТНОГО ВЫБОРА (MAJORITY VOTING) ===
    # Сначала собираем все "сырые" упоминания преподавателей со всех файлов
    raw_occurrences = []
    for dir_key, subjects in subjects_mapping.items():
        for disc_key, data in subjects.items():
            compilers_list = data.get("compilers", [])
            for comp_str in compilers_list:
                name, degree = split_compiler_name_and_degree(comp_str)
                if name:
                    raw_occurrences.append({
                        "name": name,
                        "degree": degree,
                        "discipline": disc_key,
                        "direction": dir_key,
                        "profile": data.get("profile", ""),
                        "dean": data.get("dean", ""),
                        "head_of_department": data.get("head_of_department", ""),
                        "program_director": data.get("program_director", "")
                    })

    # Группируем схожие записи в кластеры
    clusters = []  # Список словарей: {"names_freq": {}, "degrees_freq": {}, "disciplines": []}

    for occ in raw_occurrences:
        occ_name = occ["name"]
        occ_degree = occ["degree"]

        matched_cluster = None
        for cl in clusters:
            # Сравниваем с наиболее популярным вариантом имени в текущем кластере
            canonical_name = max(cl["names_freq"], key=cl["names_freq"].get)
            if is_similar_name(occ_name, canonical_name):
                matched_cluster = cl
                break

        if matched_cluster is None:
            matched_cluster = {
                "names_freq": {},
                "degrees_freq": {},
                "disciplines": []
            }
            clusters.append(matched_cluster)

        # Подсчет частоты упоминаний вариантов ФИО и ученых степеней
        matched_cluster["names_freq"][occ_name] = matched_cluster["names_freq"].get(occ_name, 0) + 1
        if occ_degree:
            matched_cluster["degrees_freq"][occ_degree] = matched_cluster["degrees_freq"].get(occ_degree, 0) + 1

        matched_cluster["disciplines"].append({
            "name": occ["discipline"],
            "direction": occ["direction"],
            "profile": occ["profile"],
            "dean": occ["dean"],
            "head_of_department": occ["head_of_department"],
            "program_director": occ["program_director"]
        })

    # Формируем итоговый маппинг по каноническим преподавателям
    teachers_mapping: Dict[str, Dict[str, Any]] = {}
    for cl in clusters:
        # Выбираем самый часто встречающийся вариант ФИО (Majority Voting)
        canonical_name = max(cl["names_freq"], key=cl["names_freq"].get)

        # Выбираем наиболее полную или часто встречающуюся ученую степень/звание
        canonical_degree = ""
        if cl["degrees_freq"]:
            # Приоритет отдается наиболее часто встречающемуся значению, а при равенстве - более длинному
            canonical_degree = max(cl["degrees_freq"], key=lambda k: (cl["degrees_freq"][k], len(k)))

        teachers_mapping[canonical_name] = {
            "degree_and_title": canonical_degree,
            "disciplines": cl["disciplines"],
            "variations_found": cl["names_freq"]  # Полезно для отладки
        }

    output_mapping_path = Path("services/rp_generator/rp_personnel_mapping.json")
    output_mapping_path.parent.mkdir(parents=True, exist_ok=True)

    final_output = {
        "default_department_personnel": default_staff,
        "teachers_mapping": teachers_mapping
    }

    try:
        with open(output_mapping_path, "w", encoding="utf-8") as f:
            json.dump(final_output, f, ensure_ascii=False, indent=2)
        print(
            f"\n[Успешно] Маппинг преподавателей с нечетким сопоставлением сохранен в:\n{output_mapping_path.absolute()}")
    except Exception as e:
        logger.error(f"Не удалось сохранить итоговый JSON: {e}")


if __name__ == "__main__":
    main()