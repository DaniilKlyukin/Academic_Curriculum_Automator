import os
import logging
from pathlib import Path
from fpdf import FPDF
from typing import List

# Настройка логгера для вывода в консоль
logger = logging.getLogger(__name__)


class ImageToPDFService:
    """Сервис для группировки изображений (сканов) в PDF документы."""

    def __init__(self, images_per_pdf: int = 3):
        self.images_per_pdf = images_per_pdf
        self.supported_ext = ('.jpg', '.jpeg', '.png')

    def generate_pdfs(self, input_path: str, output_path: str = None):
        """
        Сканирует папку, группирует картинки по 3 и сохраняет PDF.
        """
        src_dir = Path(input_path)
        if not src_dir.exists():
            logger.error(f"Путь не найден: {input_path}")
            return

        # Если выходная папка не указана, создаем 'PDF' внутри входной
        dst_dir = Path(output_path) if output_path else src_dir / "PDF_Output"
        dst_dir.mkdir(parents=True, exist_ok=True)

        # Получаем список всех подходящих изображений
        files = sorted([
            f for f in os.listdir(src_dir)
            if f.lower().endswith(self.supported_ext) and not f.startswith('~$')
        ])

        if not files:
            logger.warning(f"В папке {src_dir} не найдено подходящих изображений.")
            return

        logger.info(f"Найдено изображений: {len(files)}. Группировка по {self.images_per_pdf}...")

        # Разбиваем список на тройки (или другое кол-во)
        groups = [files[i:i + self.images_per_pdf] for i in range(0, len(files), self.images_per_pdf)]

        success_count = 0
        for group in groups:
            if len(group) < self.images_per_pdf:
                logger.warning(f"Последняя группа содержит меньше {self.images_per_pdf} изображений: {group}")

            # Формируем имя PDF по первому файлу в группе
            # Убираем расширение и последний символ (обычно это цифра 1, 2 или 3)
            base_file = Path(group[0])
            pdf_name = base_file.stem
            if len(pdf_name) > 1:
                # Если имя заканчивается на цифру или индекс (напр. 'Математика1'), отрезаем его
                pdf_name = pdf_name.rstrip('0123456789_ ')

            pdf_file_path = dst_dir / f"{pdf_name}.pdf"

            try:
                self._create_pdf_from_images(src_dir, group, pdf_file_path)
                logger.info(f"Создан PDF: {pdf_file_path.name}")
                success_count += 1
            except Exception as e:
                logger.error(f"Ошибка при создании {pdf_file_path.name}: {e}")

        print(f"\n--- Готово! Успешно создано PDF: {success_count} ---")

    def _create_pdf_from_images(self, root_dir: Path, image_names: List[str], output_path: Path):
        """Внутренний метод генерации PDF через FPDF."""
        pdf = FPDF()
        for img_name in image_names:
            pdf.add_page()
            img_full_path = str(root_dir / img_name)
            # Растягиваем на весь лист A4 (210x297 мм)
            pdf.image(img_full_path, 0, 0, 210, 297)

        pdf.output(str(output_path), "F")


def main():
    print("=== Конвертер Сканов (JPG -> PDF) ===")
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    input_dir = input("Введите путь к папке со сканами (.jpg): ").strip().strip('"')
    output_dir = input("Введите путь для сохранения PDF (Enter для папки по умолчанию): ").strip().strip('"')

    if not output_dir:
        output_dir = None

    service = ImageToPDFService(images_per_pdf=3)
    service.generate_pdfs(input_dir, output_dir)


if __name__ == "__main__":
    main()