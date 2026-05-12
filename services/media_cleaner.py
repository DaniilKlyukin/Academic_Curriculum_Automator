import os
import zipfile
import shutil
import tempfile
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Set, Any, Union, Optional
from docx import Document
from docx.oxml.ns import qn, nsmap
from docx.document import Document as DocumentType
from docx.table import _Cell
from docx.section import _Header, _Footer

if 'v' not in nsmap:
    nsmap['v'] = 'urn:schemas-microsoft-com:vml'

THRESHOLD_EMU: int = 20 * 360000
THRESHOLD_PT: float = 20 * 28.35


class WordImageCleanerDocx:
    """
    Класс для очистки документов Word от крупных изображений (сканов).
    Удаляет элементы, превышающие заданный порог, и выполняет очистку неиспользуемых медиа-файлов внутри архива docx.
    """

    def __init__(self, input_dir: Union[str, Path]) -> None:
        self.input_dir: Path = Path(input_dir)

    def process_all(self) -> None:
        """
        Рекурсивно находит все файлы .docx в директории и запускает процесс очистки.
        """
        files: List[Path] = list(self.input_dir.rglob("*.docx"))
        total: int = len(files)
        print(f"\n{'№':<9} | {'Статус':<8} | {'Файл'}")
        print("-" * 80)

        for i, file_path in enumerate(files, 1):
            if not file_path.name.startswith("~$"):
                self._clean_single_document(file_path)
                print(f"[{i:03}/{total:03}] | {'DONE':<8} | {file_path.name[:60]}")

    def _clean_single_document(self, file_path: Path) -> None:
        """
        Открывает документ, удаляет крупные изображения из всех частей и сохраняет результат.
        """
        try:
            doc: DocumentType = Document(file_path)
            removed: int = 0

            parts: List[Any] = [doc]
            for section in doc.sections:
                parts.extend([section.header, section.footer])

            for part in parts:
                removed += self._remove_large_elements(part)

            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        removed += self._remove_large_elements(cell)

            if removed > 0:
                doc.save(file_path)
                self._garbage_collect_media(file_path)
        except Exception:
            pass

    def _remove_large_elements(self, container: Any) -> int:
        """
        Ищет и удаляет элементы w:drawing и w:pict, если их размеры превышают порог.
        """
        count: int = 0
        element: Any = container._element if hasattr(container, '_element') else container

        for drawing in element.findall(".//" + qn('w:drawing')):
            extent: Optional[ET.Element] = drawing.find(".//" + qn('wp:extent'))
            if extent is not None:
                try:
                    cx: int = int(extent.get('cx', 0))
                    cy: int = int(extent.get('cy', 0))
                    if cx > THRESHOLD_EMU or cy > THRESHOLD_EMU:
                        parent: Optional[Any] = drawing.getparent()
                        if parent is not None:
                            parent.remove(drawing)
                            count += 1
                except (ValueError, TypeError):
                    continue

        for pict in element.findall(".//" + qn('w:pict')):
            for shape in pict.findall(f".//{{{nsmap['v']}}}shape"):
                style: str = shape.get('style', '')
                w_m: Optional[re.Match] = re.search(r'width:(\d+\.?\d*)pt', style)
                h_m: Optional[re.Match] = re.search(r'height:(\d+\.?\d*)pt', style)

                is_large: bool = False
                if w_m and float(w_m.group(1)) > THRESHOLD_PT:
                    is_large = True
                if h_m and float(h_m.group(1)) > THRESHOLD_PT:
                    is_large = True

                if is_large:
                    parent_pict: Optional[Any] = pict.getparent()
                    if parent_pict is not None:
                        parent_pict.remove(pict)
                        count += 1
                        break
        return count

    def _garbage_collect_media(self, file_path: Path) -> None:
        """
        Распаковывает docx как ZIP, анализирует связи и удаляет файлы в media/, которые не упоминаются в XML.
        """
        temp_dir: Path = Path(tempfile.mkdtemp())
        rel_ns: str = 'http://schemas.openxmlformats.org/package/2006/relationships'
        ET.register_namespace('', rel_ns)

        try:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            used_rids: Set[str] = set()
            for xml_file in temp_dir.rglob('*.xml'):
                if '_rels' in xml_file.parts:
                    continue
                try:
                    with open(xml_file, 'r', encoding='utf-8') as f:
                        content: str = f.read()
                        used_rids.update(re.findall(r'r:(?:embed|id|pict|link)="([^"]+)"', content))
                except Exception:
                    continue

            targets_to_keep: Set[str] = set()
            for rels_file in list(temp_dir.rglob('*.rels')):
                tree: ET.ElementTree = ET.parse(rels_file)
                root: ET.Element = tree.getroot()
                changed: bool = False

                for rel in root.findall(f'{{{rel_ns}}}Relationship'):
                    rid: Optional[str] = rel.get('Id')
                    target: Optional[str] = rel.get('Target')
                    rel_type: str = rel.get('Type', '')

                    if not rid or not target:
                        continue

                    if ("image" in rel_type or "media" in target.lower()) and rid not in used_rids:
                        root.remove(rel)
                        changed = True
                    else:
                        targets_to_keep.add(os.path.basename(target))

                if changed:
                    tree.write(rels_file, encoding='utf-8', xml_declaration=True)

            media_dir: Path = temp_dir / 'word' / 'media'
            if media_dir.exists():
                for img_file in media_dir.glob('*'):
                    if img_file.name not in targets_to_keep:
                        try:
                            os.remove(img_file)
                        except OSError:
                            pass

            with zipfile.ZipFile(file_path, 'w', compression=zipfile.ZIP_DEFLATED) as new_zip:
                for file in temp_dir.rglob('*'):
                    if file.is_file():
                        new_zip.write(file, file.relative_to(temp_dir))
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)