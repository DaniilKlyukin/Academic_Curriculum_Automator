import os
import logging
import zipfile
import shutil
import tempfile
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn, nsmap

# Настройка логирования
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# ПРИНУДИТЕЛЬНАЯ РЕГИСТРАЦИЯ пространства имен VML (исправляет KeyError: 'v')
if 'v' not in nsmap:
    nsmap['v'] = 'urn:schemas-microsoft-com:vml'

THRESHOLD_EMU = 20 * 360000
THRESHOLD_PT = 20 * 28.35


class WordImageCleanerDocx:
    def __init__(self, input_dir: str):
        self.input_dir = Path(input_dir)

    def process_all(self):
        files = list(self.input_dir.glob("*.docx"))
        if not files:
            logger.warning(f"В папке {self.input_dir} не найдено .docx файлов.")
            return

        for file_path in files:
            if file_path.name.startswith("~$"): continue
            self._clean_single_document(file_path)

    def _clean_single_document(self, file_path: Path):
        try:
            doc = Document(file_path)
            removed_count = 0

            # Список всех частей документа, где могут быть картинки
            parts = [doc]
            for section in doc.sections:
                parts.extend([section.header, section.footer])

            # Обработка основного текста, колонтитулов и таблиц
            for part in parts:
                removed_count += self._remove_large_elements(part)

            # Обработка таблиц
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        removed_count += self._remove_large_elements(cell)

            if removed_count > 0:
                doc.save(file_path)
                # Чистим мусор (картинки физически и связи в .rels)
                self._garbage_collect_media(file_path)
                logger.info(f"Файл {file_path.name}: Удалено {removed_count} объектов.")
            else:
                logger.info(f"Файл {file_path.name}: Крупных сканов не найдено.")

        except Exception as e:
            logger.error(f"Ошибка при обработке {file_path.name}: {e}")

    def _remove_large_elements(self, container):
        count = 0
        element = container._element if hasattr(container, '_element') else container

        # 1. Новые рисунки (w:drawing)
        for drawing in element.findall(".//" + qn('w:drawing')):
            extent = drawing.find(".//" + qn('wp:extent'))
            if extent is not None:
                try:
                    w = int(extent.get('cx', 0))
                    h = int(extent.get('cy', 0))
                    if w > THRESHOLD_EMU or h > THRESHOLD_EMU:
                        drawing.getparent().remove(drawing)
                        count += 1
                except:
                    continue

        # 2. Старые VML (w:pict) - ИСПРАВЛЕНО
        for pict in element.findall(".//" + qn('w:pict')):
            # Используем полный путь пространства имен, если qn('v:shape') падает
            v_shape_tag = f"{{{nsmap['v']}}}shape"
            shapes = pict.findall(".//" + v_shape_tag)

            for shape in shapes:
                style = shape.get('style', '')
                w_m = re.search(r'width:(\d+\.?\d*)pt', style)
                h_m = re.search(r'height:(\d+\.?\d*)pt', style)

                is_large = False
                if w_m and float(w_m.group(1)) > THRESHOLD_PT: is_large = True
                if h_m and float(h_m.group(1)) > THRESHOLD_PT: is_large = True

                if is_large:
                    parent = pict.getparent()
                    if parent is not None:
                        parent.remove(pict)
                        count += 1
                        break
        return count

    def _garbage_collect_media(self, file_path: Path):
        """ Удаление файлов картинок из архива и записей в .rels """
        temp_dir = Path(tempfile.mkdtemp())
        rel_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
        ET.register_namespace('', rel_ns)  # Важно для корректного сохранения .rels

        try:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            # 1. Находим все rId, которые реально остались в XML-файлах
            used_rids = set()
            for xml_file in temp_dir.rglob('*.xml'):
                if '_rels' in xml_file.parts: continue
                try:
                    with open(xml_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                        # Ищем все ссылки на связи
                        used_rids.update(re.findall(r'r:(?:embed|id|pict|link)="([^"]+)"', content))
                except:
                    continue

            # 2. Чистим .rels файлы
            targets_to_keep = set()
            for rels_file in list(temp_dir.rglob('*.rels')):
                tree = ET.parse(rels_file)
                root = tree.getroot()
                changed = False

                for rel in root.findall(f'{{{rel_ns}}}Relationship'):
                    rid = rel.get('Id')
                    target = rel.get('Target')
                    rel_type = rel.get('Type', '')

                    # Удаляем только если это медиа/изображение и оно не используется
                    is_media = "image" in rel_type or "media" in target.lower()

                    if is_media and rid not in used_rids:
                        root.remove(rel)
                        changed = True
                    else:
                        targets_to_keep.add(os.path.basename(target))

                if changed:
                    tree.write(rels_file, encoding='utf-8', xml_declaration=True)

            # 3. Удаляем сами файлы из word/media
            media_dir = temp_dir / 'word' / 'media'
            if media_dir.exists():
                for img_file in media_dir.glob('*'):
                    if img_file.name not in targets_to_keep:
                        os.remove(img_file)

            # 4. Собираем обратно
            with zipfile.ZipFile(file_path, 'w', compression=zipfile.ZIP_DEFLATED) as new_zip:
                for file in temp_dir.rglob('*'):
                    if file.is_file():
                        new_zip.write(file, file.relative_to(temp_dir))

        finally:
            shutil.rmtree(temp_dir)


if __name__ == "__main__":
    print("--- Очистка DOCX от крупных картинок (Версия 3.2) ---")
    path_input = input("Введите путь до папки с .docx файлами: ").strip().strip('"')
    if os.path.isdir(path_input):
        cleaner = WordImageCleanerDocx(path_input)
        cleaner.process_all()
        print("\nГотово!")
    else:
        print("Ошибка: Путь не найден.")