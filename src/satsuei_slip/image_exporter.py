from __future__ import annotations

import tempfile
from pathlib import Path

import fitz

from satsuei_slip.models import DeliveryInfo, VideoItem
from satsuei_slip.pdf_exporter import DeliverySlipPdfExporter


class DeliverySlipImageExporter:
    def __init__(self) -> None:
        self.pdf_exporter = DeliverySlipPdfExporter()

    def export(self, output_path: Path, delivery_info: DeliveryInfo, items: list[VideoItem]) -> list[Path]:
        output_path = output_path.with_suffix(".png")
        output_paths: list[Path] = []
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_pdf = Path(temp_dir) / "delivery_slip.pdf"
            self.pdf_exporter.export(temp_pdf, delivery_info, items)
            with fitz.open(temp_pdf) as doc:
                page_count = doc.page_count
                for page_index in range(page_count):
                    page = doc.load_page(page_index)
                    pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                    if page_count == 1:
                        page_output = output_path
                    else:
                        page_output = output_path.with_name(
                            f"{output_path.stem}_{page_index + 1:02d}{output_path.suffix}"
                        )
                    pixmap.save(page_output)
                    output_paths.append(page_output)
        return output_paths
