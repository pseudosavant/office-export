from __future__ import annotations

from pathlib import Path

import pypdfium2 as pdfium


def create_blank_pdf(path: Path, sizes: list[tuple[float, float]]) -> Path:
    document = pdfium.PdfDocument.new()
    try:
        for width, height in sizes:
            page = document.new_page(width, height)
            page.close()
        document.save(path)
    finally:
        document.close()
    return path
