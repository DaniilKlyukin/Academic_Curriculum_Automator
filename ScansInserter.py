from ScanInsertRules import first_page_resolver, second_page_resolver, coordination_page_resolver
from ScanInsertRule import ScanInsertRule
from ScanFinder import ScanFinder
from ScanInsertionManager import ScanInsertionManager

from os import path, listdir
from os.path import join
from typing import List, Tuple, Optional
import logging
from comtypes.client import CreateObject
from comtypes.gen import Word as word


logger = logging.getLogger(__name__)


def _find_coordination_page(doc) -> Optional[int]:
    """Находит номер страницы с 'Лист согласования'."""
    find = doc.Content.Find
    find.ClearFormatting()
    find.Text = "Лист согласования"
    find.Forward = True
    find.MatchCase = False
    find.MatchWholeWord = False

    if find.Execute():
        return find.Parent.Information(word.wdActiveEndPageNumber)
    return None


def _insert_image_over_page(doc, page_num: int, image_path: str):
    """Вставляет изображение поверх указанной страницы без учета полей."""
    # Находим диапазон, соответствующий нужной странице
    range_start = doc.Range().GoTo(
        What=word.wdGoToPage,
        Which=word.wdGoToAbsolute,
        Count=page_num
    ).Start

    range_end = doc.Range().GoTo(
        What=word.wdGoToPage,
        Which=word.wdGoToAbsolute,
        Count=page_num + 1
    ).Start - 1 if page_num < doc.ComputeStatistics(word.wdStatisticPages) else doc.Content.End

    page_range = doc.Range(range_start, range_end)

    # Добавляем изображение в этот диапазон
    shape = doc.Shapes.AddPicture(
        FileName=image_path,
        LinkToFile=False,
        SaveWithDocument=True,
        Left=0,
        Top=0,
        Width=doc.PageSetup.PageWidth,
        Height=doc.PageSetup.PageHeight,
        Anchor=page_range
    )

    # Настройка обтекания - перед текстом
    shape.WrapFormat.Type = word.wdWrapFront

    # Позиционирование относительно страницы
    shape.RelativeHorizontalPosition = word.wdRelativeHorizontalPositionPage
    shape.RelativeVerticalPosition = word.wdRelativeVerticalPositionPage
    shape.Left = word.wdShapeLeft  # 0 - край страницы
    shape.Top = word.wdShapeTop  # 0 - верх страницы


def _find_scans_files(scan_folder: str, images_extensions: Optional[List[str]] = None) -> List[Tuple[str, str, str]]:
    """Находит все тройки файлов сканов."""
    images_extensions = images_extensions or ['.jpg', '.jpeg', '.png']

    files = [
        f for f in listdir(scan_folder)
        if any(f.lower().endswith(ext.lower()) for ext in images_extensions) and '~' not in f
    ]
    files.sort()

    if len(files) % 3 != 0:
        logger.warning(
            f"Количество файлов изображений ({len(files)}) не кратно 3. "
            "Возможно, не все документы будут обработаны."
        )

    files = [join(scan_folder, f) for f in files]

    return [tuple(files[i:i + 3]) for i in range(0, len(files), 3)]


class ScansInserter:
    """Класс для наложения сканов подписей поверх страниц Word документов."""

    def insert_scans(self, words_files: List[str], scans_triples: List[Tuple[str, str, str]]) -> None:
        """
        Накладывает подписи поверх страниц Word документов.

        :param words_files: Список путей к Word файлам в правильном порядке
        :param scans_triples: Список троек путей к сканам
        """
        if len(scans_triples) != len(words_files):
            logger.warning(
                f"Количество троек подписей ({len(scans_triples)}) "
                f"не совпадает с количеством Word файлов ({len(words_files)})."
            )

        word_app = CreateObject("Word.Application")
        word_app.Visible = False

        try:
            for i, (word_file, scans) in enumerate(zip(words_files, scans_triples)):
                try:
                    self._process_single_document(word_app, word_file, scans)
                    logger.info(f"Успешно обработан документ: {word_file}")
                except Exception as e:
                    logger.error(f"Ошибка при обработке {word_file}: {str(e)}")
                    continue
        finally:
            word_app.Quit()

    def _process_single_document(self, word_app, word_path: str, scans: Tuple[str, str, str]) -> None:
        """Обрабатывает один документ Word."""
        doc = word_app.Documents.Open(word_path)

        try:
            self._insert_image_on_specific_page(doc, 1, scans[0])
            self._insert_image_on_specific_page(doc, 2, scans[1])

            coord_page_num = _find_coordination_page(doc)
            if coord_page_num:
                self._insert_image_on_specific_page(doc, coord_page_num, scans[2])
            else:
                logger.warning(f"Не найдена страница согласования в {word_path}")

            doc.Save()
        except Exception as e:
            logger.error(f"Ошибка при обработке {word_path}: {str(e)}")
            doc.Close(SaveChanges=False)
            raise
        else:
            doc.Close(SaveChanges=True)

    def _insert_image_on_specific_page(self, doc, page_num: int, image_path: str):
        """Вспомогательный метод для вставки на конкретную страницу."""
        _insert_image_over_page(doc, page_num, image_path)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    try:
        words_folder = input("Введите путь до папки с рабочими программами:\n")
        scans_folder = input("Введите путь до папки со сканами:\n")

        # Настройка правил вставки
        insertion_rules = [
            ScanInsertRule(0, first_page_resolver),  # Первый скан на первую страницу
            ScanInsertRule(1, second_page_resolver),  # Второй скан на вторую страницу
            ScanInsertRule(2, coordination_page_resolver)  # Третий скан на страницу согласования
        ]

        scan_finder = ScanFinder(scans_folder)
        manager = ScanInsertionManager(scan_finder, insertion_rules)

        word_files = sorted(
            path.join(words_folder, f) for f in listdir(words_folder)
            if f.lower().endswith(('.doc', '.docx')) and '~' not in f
        )

        if not word_files:
            logger.warning(f"Не найдено Word файлов в {words_folder}")
            exit(1)

        manager.process_documents(word_files)

    except Exception as e:
        logger.error(f"Критическая ошибка: {str(e)}", exc_info=True)
        exit(1)