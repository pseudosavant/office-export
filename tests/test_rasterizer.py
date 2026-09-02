from __future__ import annotations

from pathlib import Path

import pytest
from conftest import create_blank_pdf
from PIL import Image

from office_export.errors import ConversionError, SecurityError
from office_export.rasterizer import RasterPage, _reject_security_handler, inspect_pdf, rasterize_pdf


def test_inspect_pdf_reports_physical_pages_and_sizes(tmp_path: Path) -> None:
    source = create_blank_pdf(tmp_path / "source.pdf", [(72, 144), (216, 72)])
    result = inspect_pdf(source)
    assert result["page_count"] == 2
    assert result["pages"][0]["width_points"] == 72
    assert result["pages"][0]["height_points"] == 144
    assert result["pages"][1]["width_points"] == 216
    assert result["dependencies"]["pdfium"]


def test_rasterize_selected_png_with_dpi_metadata(tmp_path: Path) -> None:
    source = create_blank_pdf(tmp_path / "source.pdf", [(72, 144), (216, 72)])
    output = tmp_path / "images"
    files, mappings, warnings = rasterize_pdf(
        source,
        output,
        [RasterPage(2, "page-002.png", {"page": 2})],
        output_format="png",
        dpi=144,
    )
    assert files == [output / "page-002.png"]
    assert mappings[0]["page"] == 2
    assert mappings[0]["width"] == 432
    assert mappings[0]["height"] == 144
    assert warnings == []
    with Image.open(files[0]) as image:
        assert image.size == (432, 144)
        assert image.info["dpi"][0] == pytest.approx(144, abs=1)


def test_rasterize_jpeg_and_enforce_pixel_limit(tmp_path: Path) -> None:
    source = create_blank_pdf(tmp_path / "source.pdf", [(72, 72)])
    output = tmp_path / "images"
    files, _, _ = rasterize_pdf(
        source,
        output,
        [RasterPage(1, "page.jpg")],
        output_format="jpeg",
        dpi=72,
        background="#ffeecc",
    )
    with Image.open(files[0]) as image:
        assert image.mode == "RGB"
        assert image.size == (72, 72)

    with pytest.raises(ConversionError) as raised:
        rasterize_pdf(
            source,
            output,
            [RasterPage(1, "too-large.png")],
            output_format="png",
            dpi=720,
            max_megapixels=0.1,
        )
    assert raised.value.context.code == "pixel_limit_exceeded"


def test_security_handler_is_rejected() -> None:
    class Raw:
        @staticmethod
        def FPDF_GetSecurityHandlerRevision(document: object) -> int:
            return 4

        @staticmethod
        def FPDF_GetDocUserPermissions(document: object) -> int:
            return 4

    with pytest.raises(SecurityError) as raised:
        _reject_security_handler(Raw, object(), Path("restricted.pdf"))
    assert raised.value.context.code == "pdf_restricted"
