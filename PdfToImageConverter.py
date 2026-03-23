from os import cpu_count
import logging
from pathlib import Path
from typing import List
from concurrent.futures import ThreadPoolExecutor
from PIL import Image
import fitz

logger = logging.getLogger(__name__)


class PdfToImageConverter:
    """Конвертер PDF страниц в изображения с поддержкой многопоточной обработки."""

    def __init__(
            self,
            input_dir: str,
            output_dir: str,
            page_numbers: List[int] = None,
            dpi: int = 300,
            img_format: str = "png",
            max_workers: int = 4,
    ):
        """
        Инициализация конвертера.

        :param input_dir: Путь к папке с PDF файлами
        :param output_dir: Путь для сохранения изображений
        :param page_numbers: Номера страниц для конвертации (None - все страницы)
        :param dpi: Качество DPI для изображений
        :param img_format: Формат изображений (png, jpeg и т.д.)
        :param max_workers: Максимальное количество потоков
        """
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        self.page_numbers = page_numbers if page_numbers else []
        self.dpi = dpi
        self.img_format = img_format.lower()
        self.max_workers = max_workers

        self._validate_params()
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def _validate_params(self):
        """Проверка входных параметров."""
        if not self.input_dir.exists():
            raise ValueError(f"Input directory does not exist: {self.input_dir}")
        if not self.input_dir.is_dir():
            raise ValueError(f"Input path is not a directory: {self.input_dir}")
        if self.dpi < 72 or self.dpi > 1200:
            raise ValueError("DPI should be between 72 and 1200")
        if self.img_format not in ["png", "jpeg", "jpg", "tiff", "bmp"]:
            raise ValueError(f"Unsupported image format: {self.img_format}")

    def process_all_files(self):
        """Обработка всех PDF файлов в директории."""
        pdf_files = list(self.input_dir.glob("*.pdf"))
        if not pdf_files:
            logger.warning(f"No PDF files found in {self.input_dir}")
            return

        logger.info(f"Found {len(pdf_files)} PDF files to process")

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            for pdf_file in pdf_files:
                executor.submit(self._process_single_pdf, pdf_file)

    def _process_single_pdf(self, pdf_path: Path):
        """Обработка одного PDF файла."""
        try:
            with fitz.open(pdf_path) as doc:
                pages_to_convert = (
                    self.page_numbers if self.page_numbers else range(doc.page_count)
                )

                for page_num in pages_to_convert:
                    if page_num >= doc.page_count:
                        logger.warning(
                            f"Page {page_num} not found in {pdf_path.name} "
                            f"(total pages: {doc.page_count})"
                        )
                        continue

                    self._convert_page(doc, page_num, pdf_path)

        except Exception as e:
            logger.error(f"Error processing {pdf_path.name}: {str(e)}")

    def _convert_page(self, doc: fitz.Document, page_num: int, pdf_path: Path):
        """Конвертация одной страницы в изображение."""
        page = doc.load_page(page_num)
        zoom = self.dpi / 72  # 72 - стандартное DPI PDF
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)

        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        output_filename = self._generate_output_filename(pdf_path, page_num)

        # Сохраняем с оптимальными параметрами
        save_kwargs = {}
        if self.img_format in ["jpeg", "jpg"]:
            save_kwargs["quality"] = 70
            save_kwargs["optimize"] = True
        elif self.img_format == "png":
            save_kwargs["compress_level"] = 6

        img.save(output_filename, format=self.img_format, **save_kwargs)
        logger.debug(f"Saved: {output_filename}")

    def _generate_output_filename(self, pdf_path: Path, page_num: int) -> str:
        """Генерация имени выходного файла."""
        stem = pdf_path.stem
        ext = f".{self.img_format}"
        return str(self.output_dir / f"{stem}{ext}")


if __name__ == "__main__":
    logger = logging.getLogger(__name__)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )

    try:
        input_folder = input("Введите путь до папки с PDF:\n")
        output_folder = input("Введите путь до папки куда будут сохранены изображения:\n")

        converter = PdfToImageConverter(
            input_dir=input_folder,
            output_dir=output_folder,
            page_numbers=[0, 1],  # Первая и вторая страницы
            dpi=300,
            img_format="png",
            max_workers=cpu_count(),
        )
        converter.process_all_files()
        logger.info("Conversion completed successfully")
    except Exception as e:
        logger.error(f"Conversion failed: {str(e)}", exc_info=True)
