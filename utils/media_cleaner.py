import os
import zipfile
import shutil
import tempfile
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn, nsmap

if 'v' not in nsmap:
    nsmap['v'] = 'urn:schemas-microsoft-com:vml'

THRESHOLD_EMU = 20 * 360000
THRESHOLD_PT = 20 * 28.35

class WordImageCleanerDocx:
    def __init__(self, input_dir: str):
        self.input_dir = Path(input_dir)

    def process_all(self):
        files = list(self.input_dir.rglob("*.docx"))
        for file_path in files:
            if not file_path.name.startswith("~$"):
                self._clean_single_document(file_path)

    def _clean_single_document(self, file_path: Path):
        try:
            doc = Document(file_path)
            removed = 0
            parts = [doc]
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

    def _remove_large_elements(self, container):
        count = 0
        element = container._element if hasattr(container, '_element') else container
        for drawing in element.findall(".//" + qn('w:drawing')):
            extent = drawing.find(".//" + qn('wp:extent'))
            if extent is not None:
                try:
                    if int(extent.get('cx', 0)) > THRESHOLD_EMU or int(extent.get('cy', 0)) > THRESHOLD_EMU:
                        drawing.getparent().remove(drawing)
                        count += 1
                except: continue
        for pict in element.findall(".//" + qn('w:pict')):
            for shape in pict.findall(f".//{{{nsmap['v']}}}shape"):
                style = shape.get('style', '')
                w_m = re.search(r'width:(\d+\.?\d*)pt', style)
                h_m = re.search(r'height:(\d+\.?\d*)pt', style)
                if (w_m and float(w_m.group(1)) > THRESHOLD_PT) or (h_m and float(h_m.group(1)) > THRESHOLD_PT):
                    parent = pict.getparent()
                    if parent is not None:
                        parent.remove(pict)
                        count += 1
                        break
        return count

    def _garbage_collect_media(self, file_path: Path):
        temp_dir = Path(tempfile.mkdtemp())
        rel_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
        ET.register_namespace('', rel_ns)
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            used_rids = set()
            for xml_file in temp_dir.rglob('*.xml'):
                if '_rels' in xml_file.parts: continue
                try:
                    with open(xml_file, 'r', encoding='utf-8') as f:
                        used_rids.update(re.findall(r'r:(?:embed|id|pict|link)="([^"]+)"', f.read()))
                except: continue

            targets_to_keep = set()
            for rels_file in list(temp_dir.rglob('*.rels')):
                tree = ET.parse(rels_file)
                root = tree.getroot()
                changed = False
                for rel in root.findall(f'{{{rel_ns}}}Relationship'):
                    rid = rel.get('Id')
                    target = rel.get('Target')
                    if ("image" in rel.get('Type', '') or "media" in target.lower()) and rid not in used_rids:
                        root.remove(rel)
                        changed = True
                    else:
                        targets_to_keep.add(os.path.basename(target))
                if changed:
                    tree.write(rels_file, encoding='utf-8', xml_declaration=True)

            media_dir = temp_dir / 'word' / 'media'
            if media_dir.exists():
                for img_file in media_dir.glob('*'):
                    if img_file.name not in targets_to_keep:
                        os.remove(img_file)

            with zipfile.ZipFile(file_path, 'w', compression=zipfile.ZIP_DEFLATED) as new_zip:
                for file in temp_dir.rglob('*'):
                    if file.is_file():
                        new_zip.write(file, file.relative_to(temp_dir))
        finally:
            shutil.rmtree(temp_dir)