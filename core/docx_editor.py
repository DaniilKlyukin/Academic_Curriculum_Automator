from docx import Document
from docx.shared import Mm
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import os


class DocxEditor:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.doc = None

    def __enter__(self):
        self.doc = Document(self.file_path)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.doc and exc_type is None:
            self.doc.save(self.file_path)

    def _make_image_floating(self, shape, width_mm, height_mm):
        """Превращает обычное изображение в 'плавающее' поверх текста по центру страницы."""
        # Размеры в EMU (1 мм = 36000 EMU)
        width_emu = int(width_mm * 36000)
        height_emu = int(height_mm * 36000)

        inline = shape._inline
        rId = inline.graphic.graphicData.pic.blipFill.blip.embed

        # XML для Anchor (плавающего объекта)
        # simplePos="0" позволяет использовать детальное позиционирование
        # relativeFrom="page" привязывает картинку к краям листа, а не к полям
        anchor_xml = f"""
        <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240" 
            behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" 
            {nsdecls('wp', 'wp14', 'pic', 'r')}>
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="page">
                <wp:posOffset>0</wp:posOffset>
            </wp:positionH>
            <wp:positionV relativeFrom="page">
                <wp:posOffset>0</wp:posOffset>
            </wp:positionV>
            <wp:extent cx="{width_emu}" cy="{height_emu}"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="1" name="Scan Overlay"/>
            <wp:cNvGraphicFramePr>
                <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/package/2006/main" noChangeAspect="1"/>
            </wp:cNvGraphicFramePr>
            {inline.graphic.xml}
        </wp:anchor>
        """
        anchor = parse_xml(anchor_xml)
        inline.getparent().replace(inline, anchor)

    def add_floating_scan(self, paragraph, image_path, width_mm=210, height_mm=297):
        """Вставляет изображение, которое перекрывает всю страницу."""
        run = paragraph.add_run()
        # Сначала создаем временное inline изображение, чтобы получить rId
        shape = run.add_picture(image_path, width=Mm(width_mm))
        # Превращаем его в плавающее поверх страницы
        self._make_image_floating(shape, width_mm, height_mm)

    def add_image_at_beginning(self, image_path: str):
        """Титульник: Скан поверх первой страницы."""
        self.add_floating_scan(self.doc.paragraphs[0], image_path)

    def add_image_at_end(self, image_path: str):
        """Лист согласования: Скан поверх последней страницы."""
        self.add_floating_scan(self.doc.add_paragraph(), image_path)

    def insert_image_after_text(self, search_text: str, image_path: str) -> bool:
        """Поиск текста и наложение скана на текущую страницу."""
        # Поиск в параграфах
        for paragraph in self.doc.paragraphs:
            if search_text.lower() in paragraph.text.lower():
                self.add_floating_scan(paragraph, image_path)
                return True
        # Поиск в таблицах
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if search_text.lower() in cell.text.lower():
                        self.add_floating_scan(cell.paragraphs[0], image_path)
                        return True
        return False