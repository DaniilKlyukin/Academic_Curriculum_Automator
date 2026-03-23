from typing import Optional
from comtypes.gen import Word as word


class WordDocumentEditor:
    """Класс для работы с документом Word"""

    def __init__(self, word_app, doc_path: str):
        self.word_app = word_app
        self.doc_path = doc_path
        self.doc = None

    def __enter__(self):
        self.doc = self.word_app.Documents.Open(self.doc_path)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.doc:
            save_changes = exc_type is None
            self.doc.Close(SaveChanges=save_changes)

    def insert_image(self, image_path: str, page_num: int):
        """Вставляет изображение на указанную страницу"""
        if not self.doc:
            raise RuntimeError("Документ не открыт")

        range_start = self.doc.Range().GoTo(
            What=word.wdGoToPage,
            Which=word.wdGoToAbsolute,
            Count=page_num
        ).Start

        range_end = self.doc.Range().GoTo(
            What=word.wdGoToPage,
            Which=word.wdGoToAbsolute,
            Count=page_num + 1
        ).Start - 1 if page_num < self.doc.ComputeStatistics(word.wdStatisticPages) else self.doc.Content.End

        page_range = self.doc.Range(range_start, range_end)

        mm_to_points = 2.83465

        shape = self.doc.Shapes.AddPicture(
            FileName=image_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0,
            Top=0,
            Width=min(self.doc.PageSetup.PageWidth, 210.0 * mm_to_points),
            Height=min(self.doc.PageSetup.PageHeight, 297.0 * mm_to_points),
            Anchor=page_range
        )

        shape.WrapFormat.Type = word.wdWrapFront
        shape.RelativeHorizontalPosition = word.wdRelativeHorizontalPositionPage
        shape.RelativeVerticalPosition = word.wdRelativeVerticalPositionPage
        shape.Left = word.wdShapeLeft
        shape.Top = word.wdShapeTop

    def find_page_with_text(self, text: str) -> Optional[int]:
        """Находит страницу с указанным текстом"""
        find = self.doc.Content.Find
        find.ClearFormatting()
        find.Text = text
        find.Forward = True
        find.MatchCase = False
        find.MatchWholeWord = False

        if find.Execute():
            return find.Parent.Information(word.wdActiveEndPageNumber)
        return None
