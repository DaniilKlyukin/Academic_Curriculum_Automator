import os
import sys
import json
import logging
import time
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

# Импорт официального SDK Google GenAI
try:
    from google import genai
    from google.genai import types

    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")

# === СХЕМЫ ОТВЕТОВ ДЛЯ НАДЁЖНОЙ СЕРИАЛИЗАЦИИ ===
if HAS_GEMINI:
    # 1. Схема педагогического фрейма
    PedagogicalFrameSchema = types.Schema(
        type=types.Type.OBJECT,
        properties={
            "goals": types.Schema(type=types.Type.STRING, description="Цель изучения дисциплины"),
            "tasks": types.Schema(
                type=types.Type.ARRAY,
                items=types.Schema(type=types.Type.STRING),
                description="Задачи дисциплины (список)"
            ),
            "prerequisites_text": types.Schema(
                type=types.Type.STRING,
                description="Связь с предшествующими дисциплинами (какие знания и разделы необходимы)"
            ),
            "postrequisites_text": types.Schema(
                type=types.Type.STRING,
                description="Связь с последующими дисциплинами (для чего пригодятся знания этого предмета)"
            ),
            "indicators_ksa": types.Schema(
                type=types.Type.ARRAY,
                description="Сопоставление индикаторов компетенций со знаниями, умениями, навыками (З1, У1, Н1)",
                items=types.Schema(
                    type=types.Type.OBJECT,
                    properties={
                        "indicator_code": types.Schema(type=types.Type.STRING,
                                                       description="Код индикатора, например УК-1.1"),
                        "knowledge": types.Schema(
                            type=types.Type.ARRAY,
                            items=types.Schema(type=types.Type.STRING),
                            description="Знания (З1, З2...) по этому индикатору в рамках предмета"
                        ),
                        "skills": types.Schema(
                            type=types.Type.ARRAY,
                            items=types.Schema(type=types.Type.STRING),
                            description="Умения (У1, У2...) по этому индикатору"
                        ),
                        "abilities": types.Schema(
                            type=types.Type.ARRAY,
                            items=types.Schema(type=types.Type.STRING),
                            description="Навыки/Владения (Н1, Н2...) по этому индикатору"
                        )
                    },
                    required=["indicator_code", "knowledge", "skills", "abilities"]
                )
            )
        },
        required=["goals", "tasks", "prerequisites_text", "postrequisites_text", "indicators_ksa"]
    )

    # 2. Схема тематического планирования
    ThematicPlanSchema = types.Schema(
        type=types.Type.OBJECT,
        properties={
            "sections": types.Schema(
                type=types.Type.ARRAY,
                description="Основные разделы/модули дисциплины",
                items=types.Schema(
                    type=types.Type.OBJECT,
                    properties={
                        "number": types.Schema(type=types.Type.INTEGER, description="Порядковый номер раздела"),
                        "name": types.Schema(type=types.Type.STRING, description="Название раздела"),
                        "description": types.Schema(type=types.Type.STRING, description="Краткое содержание раздела")
                    },
                    required=["number", "name", "description"]
                )
            ),
            "lectures": types.Schema(
                type=types.Type.ARRAY,
                description="Темы лекционных занятий",
                items=types.Schema(
                    type=types.Type.OBJECT,
                    properties={
                        "section_number": types.Schema(type=types.Type.INTEGER,
                                                       description="Номер родительского раздела"),
                        "theme": types.Schema(type=types.Type.STRING, description="Тема лекции"),
                        "content": types.Schema(type=types.Type.STRING, description="Краткий план лекционного занятия"),
                        "hours": types.Schema(type=types.Type.INTEGER, description="Часы на эту тему")
                    },
                    required=["section_number", "theme", "content", "hours"]
                )
            ),
            "practicals": types.Schema(
                type=types.Type.ARRAY,
                description="Темы практических/семинарских занятий",
                items=types.Schema(
                    type=types.Type.OBJECT,
                    properties={
                        "section_number": types.Schema(type=types.Type.INTEGER),
                        "theme": types.Schema(type=types.Type.STRING),
                        "hours": types.Schema(type=types.Type.INTEGER)
                    },
                    required=["section_number", "theme", "hours"]
                )
            ),
            "labs": types.Schema(
                type=types.Type.ARRAY,
                description="Темы лабораторных работ",
                items=types.Schema(
                    type=types.Type.OBJECT,
                    properties={
                        "section_number": types.Schema(type=types.Type.INTEGER),
                        "theme": types.Schema(type=types.Type.STRING),
                        "hours": types.Schema(type=types.Type.INTEGER)
                    },
                    required=["section_number", "theme", "hours"]
                )
            )
        },
        required=["sections", "lectures", "practicals", "labs"]
    )

    # 3. Схема литературы и аттестации
    ResourcesAndEvaluationSchema = types.Schema(
        type=types.Type.OBJECT,
        properties={
            "primary_literature": types.Schema(
                type=types.Type.ARRAY,
                description="Список основной учебной литературы. Ссылки строго на www.iprbookshop.ru / IPR SMART, год издания 2020-2025",
                items=types.Schema(type=types.Type.STRING)
            ),
            "secondary_literature": types.Schema(
                type=types.Type.ARRAY,
                description="Список дополнительной учебной литературы",
                items=types.Schema(type=types.Type.STRING)
            ),
            "methodological_guidelines": types.Schema(
                type=types.Type.ARRAY,
                description="Методические указания и рекомендации для обучающихся",
                items=types.Schema(type=types.Type.STRING)
            ),
            "internet_resources": types.Schema(
                type=types.Type.ARRAY,
                description="Электронно-библиотечные ресурсы и интернет-ссылки",
                items=types.Schema(type=types.Type.STRING)
            ),
            "control_questions": types.Schema(
                type=types.Type.ARRAY,
                description="Контрольные вопросы для подготовки к зачету/экзамену",
                items=types.Schema(type=types.Type.STRING)
            )
        },
        required=["primary_literature", "secondary_literature", "methodological_guidelines", "internet_resources",
                  "control_questions"]
    )


class RateLimiter:
    """Ограничитель частоты вызовов API Gemini."""

    def __init__(self, rpm: int = 15):
        self.delay = 60.0 / rpm if rpm > 0 else 0.0
        self.last_call = 0.0

    def wait(self):
        if self.delay <= 0:
            return
        now = time.time()
        elapsed = now - self.last_call
        if elapsed < self.delay:
            time.sleep(self.delay - elapsed)
        self.last_call = time.time()


class RPAIGenerator:
    """Генератор интеллектуальных метаданных РП на базе API Gemini."""

    def __init__(self, project_dir: Path, rpm_limit: int = 15):
        self.project_dir = project_dir
        self.rate_limiter = RateLimiter(rpm=rpm_limit)
        self.api_key = os.environ.get("GEMINI_API_KEY", "").strip()
        self.model_name = "gemini-3.1-flash-lite-preview"

        # Инициализация путей файлов
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
        """Анализирует семестры в плане для определения логических предшественников и последователей."""
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

        return prior_candidates[:20], future_candidates[:20]  # Ограничение размера контекста

    def generate_all(self):
        if not HAS_GEMINI:
            logger.error("Библиотека google-genai не найдена. Выполните: pip install google-genai")
            return

        if not self.api_key:
            logger.error("Переменная окружения GEMINI_API_KEY не задана. Прерывание работы.")
            return

        # Загрузка данных
        workload_data = self._load_json(self.workload_path)
        comp_map_data = self._load_json(self.comp_map_path)

        metadata = workload_data.get("metadata", {})
        disciplines = workload_data.get("disciplines", {})
        comp_registry = comp_map_data.get("competencies_registry", {})
        subject_to_comp = comp_map_data.get("subject_to_competencies", {})

        # Чтение кэша
        cache = {}
        if self.cache_path.exists():
            try:
                cache = self._load_json(self.cache_path)
                logger.info(f"Загружен существующий кэш ИИ. Найдено предметов: {len(cache)}")
            except Exception:
                logger.warning("Кэш ИИ поврежден, создаем новый файл.")

        client = genai.Client(api_key=self.api_key)

        print(f"\nНайдено дисциплин в плане: {len(disciplines)}")
        target_subject = input(
            "Введите код дисциплины для генерации (или Enter для запуска по ВСЕМ предметам): ").strip()

        subject_keys = [target_subject] if target_subject in disciplines else list(disciplines.keys())

        for idx, code in enumerate(subject_keys, start=1):
            if code in cache:
                logger.info(f"[{idx}/{len(subject_keys)}] Дисциплина {code} уже есть в кэше. Пропускаем.")
                continue

            subj_info = disciplines[code]
            subj_name = subj_info["name"]
            logger.info(f"=== [{idx}/{len(subject_keys)}] Запуск генерации ИИ: {code} {subj_name} ===")

            # Часы нагрузки
            lectures_h = subj_info["total_hours"].get("lectures", 0)
            practicals_h = subj_info["total_hours"].get("practical_classes", 0)
            labs_h = subj_info["total_hours"].get("laboratory_works", 0)

            # Определение связанных компетенций
            mapped_comp = subject_to_comp.get(code, {}).get("competencies", {})
            competency_context = []
            for c_code, ind_list in mapped_comp.items():
                c_text = comp_registry.get(c_code, {}).get("competency_text", "")
                ind_info = []
                for ind_code in ind_list:
                    ind_text = comp_registry.get(c_code, {}).get("indicators", {}).get(ind_code, {}).get(
                        "indicator_text", "")
                    ind_info.append(f"  - {ind_code}: {ind_text}")
                competency_context.append(f"Компетенция {c_code}: {c_text}\nИндикаторы:\n" + "\n".join(ind_info))

            competencies_prompt = "\n\n".join(
                competency_context) if competency_context else "Общие профессиональные навыки."

            # Анализ предшественников и последователей по семестрам плана
            priors, futures = self._get_semester_relationships(code, disciplines)

            # === ЗАПРОС 1: Педагогический фрейм (Цели, задачи, связи, ЗУН/КУН) ===
            self.rate_limiter.wait()
            prompt_frame = f"""Вы — ведущий профессор вуза. Сгенерируйте педагогический фрейм (цель, задачи, связи и ЗУНы) для дисциплины «{subj_name}».
Направление подготовки: {metadata.get('direction_code')} {metadata.get('direction_name')}
Профиль: {metadata.get('profile')}
Квалификация: {metadata.get('qualification')}

Связанные компетенции и индикаторы:
{competencies_prompt}

Возможные предшествующие предметы из учебного плана: {', '.join(priors) if priors else 'Школьная программа'}
Возможные последующие предметы из учебного плана: {', '.join(futures) if futures else 'Дипломное проектирование'}

Требования:
1. Задачи сформулируйте списком (не менее 3).
2. Сделайте ссылки на предшествующие и последующие предметы строго из предложенных списков.
3. Распишите Знания, Умения, Владения для КАЖДОГО связанного индикатора (например, для {list(mapped_comp.keys())[0] if mapped_comp else 'УК-1.1'}).
"""
            try:
                logger.info("  -> Вызов Запроса 1 (Педагогический фрейм)...")
                res_frame = client.models.generate_content(
                    model=self.model_name,
                    contents=prompt_frame,
                    config=types.GenerateContentConfig(response_mime_type="application/json",
                                                       response_schema=PedagogicalFrameSchema, temperature=0.3)
                )
                frame_data = json.loads(res_frame.text)
            except Exception as e:
                logger.error(f"Не удалось выполнить Запрос 1 для {code}: {e}")
                continue

            # === ЗАПРОС 2: Тематическое планирование (Разделы, лекции, практики, лабы) ===
            self.rate_limiter.wait()
            prompt_themes = f"""Сгенерируйте подробный календарный план занятий для дисциплины «{subj_name}».
Вам нужно распределить часы на лекции, практические и лабораторные занятия.

Выделенная учебная нагрузка:
- Лекции: {lectures_h} ч.
- Практические занятия: {practicals_h} ч.
- Лабораторные работы: {labs_h} ч.

Требования:
1. Создайте от 3 до 6 логических разделов (сегментов) дисциплины.
2. Сформируйте темы лекций. Суммарное количество часов всех тем лекций должно строго равняться {lectures_h}. Темы должны быть привязаны к разделам.
3. Сформируйте темы практических занятий (семинаров). Сумма часов должна строго равняться {practicals_h}. Если практических занятий нет (0 ч.), верните пустой массив.
4. Сформируйте темы лабораторных работ. Сумма часов должна строго равняться {labs_h}. Если лаб нет (0 ч.), верните пустой массив.
5. Распределите темы по соответствующим номерам разделов. Темы должны быть профессионально сформулированными и современными.
"""
            try:
                logger.info("  -> Вызов Запроса 2 (Тематическое планирование)...")
                res_themes = client.models.generate_content(
                    model=self.model_name,
                    contents=prompt_themes,
                    config=types.GenerateContentConfig(response_mime_type="application/json",
                                                       response_schema=ThematicPlanSchema, temperature=0.2)
                )
                themes_data = json.loads(res_themes.text)
            except Exception as e:
                logger.error(f"Не удалось выполнить Запрос 2 для {code}: {e}")
                continue

            # === ЗАПРОС 3: Литература и аттестация (Ресурсы, IPRbookshop) ===
            self.rate_limiter.wait()
            current_year = metadata.get("start_year") or "2026"
            prompt_resources = f"""Сгенерируйте список литературы и контрольные вопросы для дисциплины «{subj_name}».
Направление подготовки: {metadata.get('direction_code')} {metadata.get('direction_name')}
Текущий год разработки программы: {current_year}

Требования:
1. Основная литература: 2-3 учебника или учебных пособия, обязательно содержащие ссылки на электронно-библиотечную систему IPRbookshop (iprbookshop.ru / IPR SMART) с указанием года издания строго в диапазоне от 2020 до {current_year}. Книги должны быть тематически связаны с предметом.
2. Дополнительная литература: 2-3 книги.
3. Методические указания: ссылки на методические рекомендации для практических/лабораторных работ по теме дисциплины.
4. Контрольные вопросы: 15-20 вопросов для подготовки студентов к промежуточной аттестации.
"""
            try:
                logger.info("  -> Вызов Запроса 3 (Ресурсы и аттестация)...")
                res_res = client.models.generate_content(
                    model=self.model_name,
                    contents=prompt_resources,
                    config=types.GenerateContentConfig(response_mime_type="application/json",
                                                       response_schema=ResourcesAndEvaluationSchema, temperature=0.4)
                )
                resources_data = json.loads(res_res.text)
            except Exception as e:
                logger.error(f"Не удалось выполнить Запрос 3 для {code}: {e}")
                continue

            # Объединение всех блоков данных воедино
            cache[code] = {
                "discipline_code": code,
                "discipline_name": subj_name,
                "pedagogical_frame": frame_data,
                "thematic_plan": themes_data,
                "resources_and_evaluation": resources_data
            }

            # Сохранение кэша на каждом шаге
            self._save_cache(cache)
            logger.info(f"  [+] Данные для «{subj_name}» успешно сгенерированы и сохранены в кэш.")

        print(f"\n[Успешно] Генерация ИИ завершена. Данные сохранены в файл:\n{self.cache_path.absolute()}")


def main():
    print("=== Интеллектуальный ИИ-генератор данных РП (Gemini) ===")

    # Путь по умолчанию
    project_dir = Path("services/rp_generator")
    if not project_dir.exists():
        project_dir.mkdir(parents=True, exist_ok=True)

    rpm_limit = 15
    rpm_input = input("Укажите лимит запросов в минуту (RPM) [По умолчанию: 15]: ").strip()
    if rpm_input.isdigit():
        rpm_limit = int(rpm_input)

    generator = RPAIGenerator(project_dir=project_dir, rpm_limit=rpm_limit)
    try:
        generator.generate_all()
    except Exception as e:
        logger.error(f"Ошибка при работе ИИ-генератора: {e}", exc_info=True)


if __name__ == "__main__":
    main()