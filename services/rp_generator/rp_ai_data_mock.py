import json
import logging
from pathlib import Path
from typing import Dict, Any, List, Tuple

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")


class RPAIMockGenerator:
    """Генератор синтетических тестовых метаданных РП (Mock) без обращения к API."""

    def __init__(self, project_dir: Path):
        self.project_dir = project_dir
        self.workload_path = project_dir / "rp_academic_workload.json"
        self.comp_map_path = project_dir / "rp_subject_competency_map.json"
        self.cache_path = project_dir / "rp_ai_generated_data.json"

    def _load_json(self, path: Path) -> Dict[str, Any]:
        if not path.exists():
            raise FileNotFoundError(f"Отсутствует важный файл: {path.name}. Сначала выполните его парсинг.")
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    def _save_cache(self, data: Dict[str, Any]):
        with open(self.cache_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def _get_semester_relationships(self, current_code: str, all_workload: Dict[str, Any]) -> Tuple[
        List[str], List[str]]:
        current_data = all_workload.get(current_code, {})
        current_sems = [int(s) for s in current_data.get("load_by_semester", {}).keys()]
        if not current_sems:
            return [], []

        min_sem = min(current_sems)
        max_sem = max(current_sems)

        prior_candidates = []
        future_candidates = []

        for code, data in all_workload.items():
            if code == current_code:
                continue
            sub_sems = [int(s) for s in data.get("load_by_semester", {}).keys()]
            if not sub_sems:
                continue

            sub_max = max(sub_sems)
            sub_min = min(sub_sems)

            if sub_max < min_sem:
                prior_candidates.append(data["name"])
            elif sub_min > max_sem:
                future_candidates.append(data["name"])

        return prior_candidates[:20], future_candidates[:20]

    def generate_all(self):
        workload_data = self._load_json(self.workload_path)
        comp_map_data = self._load_json(self.comp_map_path)

        metadata = workload_data.get("metadata", {})
        disciplines = workload_data.get("disciplines", {})
        comp_registry = comp_map_data.get("competencies_registry", {})
        subject_to_comp = comp_map_data.get("subject_to_competencies", {})

        cache = {}
        if self.cache_path.exists():
            try:
                cache = self._load_json(self.cache_path)
                logger.info(f"Загружен существующий кэш. Найдено предметов: {len(cache)}")
            except Exception:
                logger.warning("Кэш ИИ поврежден, создаем новый файл.")

        print(f"\nНайдено дисциплин в плане: {len(disciplines)}")
        target_subject = input(
            "Введите код дисциплины для генерации мока (или Enter для запуска по ВСЕМ предметам): ").strip()

        subject_keys = [target_subject] if target_subject in disciplines else list(disciplines.keys())

        for idx, code in enumerate(subject_keys, start=1):
            if code in cache:
                logger.info(f"[{idx}/{len(subject_keys)}] Дисциплина {code} уже есть в кэше. Пропускаем.")
                continue

            subj_info = disciplines[code]
            subj_name = subj_info["name"]
            logger.info(f"=== [{idx}/{len(subject_keys)}] Локальная имитация (Mock): {code} {subj_name} ===")

            # Часы нагрузки
            lectures_h = subj_info["total_hours"].get("lectures", 0)
            practicals_h = subj_info["total_hours"].get("practical_classes", 0)
            labs_h = subj_info["total_hours"].get("laboratory_works", 0)

            # Определение связанных компетенций
            mapped_comp = subject_to_comp.get(code, {}).get("competencies", {})
            priors, futures = self._get_semester_relationships(code, disciplines)

            # Блок 1. Мок-Фрейм (Глобальный пул ЗУН независимой длины)
            mock_knowledge = [
                f"Синтаксис и семантика алгоритмического языка программирования, принципы и методология построения систем по курсу «{subj_name}»",
                f"Концепции и идеи объектно-ориентированного программирования и проектирования программных решений",
                f"Технология работы на ПК в современных операционных средах, основные методы разработки эффективных структур данных"
            ]
            mock_skills = [
                f"Формализовать прикладную задачу, выбирать для неё подходящие структуры данных и алгоритмы обработки в «{subj_name}»",
                f"Программировать прикладные алгоритмы, используя программные средства языков высокого уровня; разрабатывать тесты",
                f"Проектировать программные компоненты IT-систем и модули обработки информации"
            ]
            mock_abilities = [
                f"Использовать современные математические методы и программные средства для решения задач науки, образования и бизнеса",
                f"Реализовать идеи объектно-ориентированного подхода и паттернов проектирования в предметной области «{subj_name}»",
                f"Применять современные цифровые технологии моделирования, алгоритмизации и оптимизации бизнес-процессов"
            ]

            # Построение реляционного маппинга ЗУН на индикаторы строго по диагонали
            mock_indicator_mappings = []
            for c_code, ind_list in mapped_comp.items():
                for i_code in ind_list:
                    # Извлекаем последнюю цифру индикатора после точки (например, "1" из "ОПК-3.1")
                    parts = i_code.split(".")
                    last_digit = parts[-1] if parts else ""

                    k_idx = [1, 2, 3] if last_digit == "1" else []
                    s_idx = [1, 2, 3] if last_digit == "2" else []
                    a_idx = [1, 2, 3] if last_digit == "3" else []

                    mock_indicator_mappings.append({
                        "indicator_code": i_code,
                        "knowledge_indices": k_idx,
                        "skills_indices": s_idx,
                        "abilities_indices": a_idx
                    })

            frame_data = {
                "goals": f"Формирование систематизированных знаний, умений и практических навыков по дисциплине «{subj_name}».",
                "tasks": [
                    f"Изучить теоретические основы и понятия курса «{subj_name}».",
                    f"Освоить методы решения прикладных задач в предметной области.",
                    f"Получить опыт выполнения самостоятельных проектов по теме «{subj_name}»."
                ],
                "prerequisites_text": f"Дисциплина опирается на компетенции, полученные при изучении: {', '.join(priors) if priors else 'школьной программы'}.",
                "postrequisites_text": f"Знания пригодятся при изучении последующих дисциплин: {', '.join(futures) if futures else 'дипломного проектирования'}.",
                "knowledge_list": mock_knowledge,
                "skills_list": mock_skills,
                "abilities_list": mock_abilities,
                "indicator_mappings": mock_indicator_mappings
            }

            # Блок 2. Мок-План (Лекции/Практики/Лабы)
            mock_sections = [
                {"number": 1, "name": f"Введение в предметную область «{subj_name}»",
                 "description": "Базовые понятия, терминология, предмет исследования."},
                {"number": 2, "name": f"Основной прикладной раздел «{subj_name}»",
                 "description": "Методология, ключевые алгоритмы и технологии разработки."}
            ]

            mock_lectures = []
            if lectures_h > 0:
                h_lec = max(1, lectures_h // 2)
                mock_lectures.append({"section_number": 1, "theme": f"Тема лекции 1. Раздел {subj_name}",
                                      "content": "Вводный обзор дисциплины.", "hours": h_lec})
                mock_lectures.append({"section_number": 2, "theme": f"Тема лекции 2. Применение {subj_name}",
                                      "content": "Специфика практической реализации.", "hours": lectures_h - h_lec})

            mock_practicals = []
            if practicals_h > 0:
                h_pr = max(1, practicals_h // 2)
                mock_practicals.append(
                    {"section_number": 1, "theme": f"Семинар 1. Базовые концепции {subj_name}", "hours": h_pr})
                mock_practicals.append({"section_number": 2, "theme": f"Семинар 2. Решение кейсов {subj_name}",
                                        "hours": practicals_h - h_pr})

            mock_labs = []
            if labs_h > 0:
                h_lb = max(1, labs_h // 2)
                mock_labs.append({"section_number": 1, "theme": f"Лабораторная работа 1. Начало работы", "hours": h_lb})
                mock_labs.append({"section_number": 2, "theme": f"Лабораторная работа 2. Оптимизация системы",
                                  "hours": labs_h - h_lb})

            themes_data = {"sections": mock_sections, "lectures": mock_lectures, "practicals": mock_practicals,
                           "labs": mock_labs}

            mock_competency_tests = []
            for c_code in mapped_comp.keys():
                mock_competency_tests.append({
                    "competency_code": c_code,
                    "questions": [
                        {
                            "question": f"Вопрос 1 для контроля уровня сформированности компетенции {c_code} в рамках курса «{subj_name}»?",
                            "options": [
                                "Правильный вариант ответа по теоретическим основам и синтаксису",
                                "Вспомогательный некорректный параметр инструментальной среды",
                                "Альтернативное решение без системной поддержки архитектуры",
                                "Абстрактная формулировка понятия, не относящаяся к делу"
                            ],
                            "correct_answer": "1"
                        },
                        {
                            "question": f"Вопрос 2 для контроля уровня сформированности компетенции {c_code} в рамках курса «{subj_name}»?",
                            "options": [
                                "Правильный вариант ответа по практическому применению и методологии",
                                "Вспомогательный некорректный параметр инструментальной среды",
                                "Альтернативное решение без системной поддержки архитектуры",
                                "Абстрактная формулировка понятия, не относящаяся к делу"
                            ],
                            "correct_answer": "1"
                        },
                        {
                            "question": f"Вопрос 3 для контроля уровня сформированности компетенции {c_code} в рамках курса «{subj_name}»?",
                            "options": [
                                "Правильный вариант ответа по отладке и интеграции компонентов",
                                "Вспомогательный некорректный параметр инструментальной среды",
                                "Альтернативное решение без системной поддержки архитектуры",
                                "Абстрактная формулировка понятия, не относящаяся к делу"
                            ],
                            "correct_answer": "1"
                        },
                        {
                            "question": f"Вопрос 4 для контроля уровня сформированности компетенции {c_code} в рамках курса «{subj_name}»?",
                            "options": [
                                "Правильный вариант ответа по оценке сложности используемых алгоритмов",
                                "Вспомогательный некорректный параметр инструментальной среды",
                                "Альтернативное решение без системной поддержки архитектуры",
                                "Абстрактная формулировка понятия, не относящаяся к делу"
                            ],
                            "correct_answer": "1"
                        },
                        {
                            "question": f"Вопрос 5 для контроля уровня сформированности компетенции {c_code} в рамках курса «{subj_name}»?",
                            "options": [
                                "Правильный вариант ответа по верификации и тестированию программных модулей",
                                "Вспомогательный некорректный параметр инструментальной среды",
                                "Альтернативное решение без системной поддержки архитектуры",
                                "Абстрактная формулировка понятия, не относящаяся к делу"
                            ],
                            "correct_answer": "1"
                        }
                    ]
                })

            # Блок 3. Мок-Ресурсы (Литература IPR SMART)
            resources_data = {
                "primary_literature": [
                    f"Иванов И.И. Теоретические основы курса «{subj_name}»: учебник для вузов. — Москва: IPR SMART, 2024. — 320 c. — URL: https://www.iprbookshop.ru/115123.html.",
                    f"Петров П.П. Практическое руководство по дисциплине «{subj_name}»: учебное пособие. — Саратов: ВУЗ-Издат, 2023. — 150 c. — URL: https://www.iprbookshop.ru/124321.html."
                ],
                "secondary_literature": [
                    f"Сидоров С.С. Архитектура и проектирование решений «{subj_name}». — Санкт-Петербург: Наука, 2022. — 200 c."
                ],
                "methodological_guidelines": [
                    f"Методические указания к лабораторным работам по «{subj_name}» / сост. Кафедра информационных технологий, 2025."
                ],
                "internet_resources": [
                    "Электронный библиотечный каталог — http://istu.ru",
                    "Российское образование — https://edu.ru"
                ],
                "control_questions": [
                    f"Каковы базовые понятия дисциплины «{subj_name}»?",
                    "Опишите структуру и принципы работы базовых методов.",
                    "Сформулируйте критерии эффективности примененных подходов.",
                    "Опишите технологию отладки и верификации результатов.",
                    "Каковы особенности стандартизации процессов в предметной области?"
                ],
                "competency_tests": mock_competency_tests
            }

            # Сохранение мок-данных в кэш
            cache[code] = {
                "discipline_code": code,
                "discipline_name": subj_name,
                "pedagogical_frame": frame_data,
                "thematic_plan": themes_data,
                "resources_and_evaluation": resources_data
            }

            self._save_cache(cache)
            logger.info(f"  [+] Заглушка для «{subj_name}» успешно сохранена в локальный кэш.")

        print(f"\n[Успешно] Мок-генерация завершена. Данные сохранены в файл:\n{self.cache_path.absolute()}")


def main():
    print("=== Тестовый оффлайн генератор мок-данных РП ===")

    # Запрос пути к папке с результатами парсинга
    user_project_dir = input(
        "Введите путь к папке с файлами JSON учебного плана (по умолчанию 'services/rp_generator'): ").strip()
    if not user_project_dir:
        user_project_dir = "services/rp_generator"

    project_dir = Path(user_project_dir)
    project_dir.mkdir(parents=True, exist_ok=True)

    generator = RPAIMockGenerator(project_dir=project_dir)
    try:
        generator.generate_all()
    except Exception as e:
        logger.error(f"Ошибка при работе мок-генератора: {e}", exc_info=True)


if __name__ == "__main__":
    main()