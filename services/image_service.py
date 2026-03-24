import os
from pathlib import Path
from fpdf import FPDF


class ImageToPDFService:
    def __init__(self, images_per_pdf: int = 3):
        self.images_per_pdf = images_per_pdf
        self.supported_ext = ('.jpg', '.jpeg', '.png')

    def generate_pdfs(self, input_path: str, output_path: str = None):
        src_dir = Path(input_path)
        if not src_dir.exists(): return
        dst_dir = Path(output_path) if output_path else src_dir / "PDF_Output"
        dst_dir.mkdir(parents=True, exist_ok=True)

        all_files = []
        for root, _, filenames in os.walk(src_dir):
            for f in sorted(filenames):
                if f.lower().endswith(self.supported_ext) and not f.startswith('~$'):
                    all_files.append(Path(root) / f)

        groups = [all_files[i:i + self.images_per_pdf] for i in range(0, len(all_files), self.images_per_pdf)]
        for group in groups:
            pdf_name = group[0].stem.rstrip('0123456789_ ')
            try:
                pdf = FPDF()
                for img_path in group:
                    pdf.add_page()
                    pdf.image(str(img_path), 0, 0, 210, 297)
                pdf.output(str(dst_dir / f"{pdf_name}.pdf"), "F")
            except Exception:
                pass

    def _create_pdf(self, root_dir: Path, image_names: list, output_path: Path):
        pdf = FPDF()
        for img_name in image_names:
            pdf.add_page()
            pdf.image(str(root_dir / img_name), 0, 0, 210, 297)
        pdf.output(str(output_path), "F")