from __future__ import annotations

import json
from pathlib import Path

import pytest
from conftest import create_blank_pdf
from PIL import Image

from office_export.core import ExportOptions, export_document, inspect_document
from office_export.errors import UsageError


def test_direct_pdf_export_to_selected_png_and_manifest(tmp_path: Path) -> None:
    source = create_blank_pdf(tmp_path / "Document.pdf", [(72, 72), (144, 72)])
    output = tmp_path / "previews"
    manifest = tmp_path / "conversion.json"
    result = export_document(
        ExportOptions(
            source=source,
            output_format="png",
            output=output,
            pages="2",
            dpi=72,
            manifest=manifest,
        )
    )
    exported = output / "document-page-002.png"
    assert result["ok"] is True
    assert result["office"] is None
    assert result["mapping"][0]["page"] == 2
    assert result["mapping"][0]["engine"] == "pdfium"
    assert result["rendering"] == {"engine": "pdfium", "temporary_pdf_used": False}
    assert result["mapping"][0]["output"] == str(exported)
    assert result["outputs"][0]["sha256"]
    assert json.loads(manifest.read_text(encoding="utf-8")) == result
    with Image.open(exported) as image:
        assert image.size == (144, 72)


def test_direct_pdf_explicit_single_jpeg_and_collision_safety(tmp_path: Path) -> None:
    source = create_blank_pdf(tmp_path / "source.pdf", [(72, 72)])
    output = tmp_path / "preview.jpg"
    first = export_document(
        ExportOptions(source=source, output_format="jpeg", output=output, pages="1", jpeg_quality=80)
    )
    assert first["outputs"][0]["path"] == str(output)
    original = output.read_bytes()
    with pytest.raises(UsageError) as raised:
        export_document(ExportOptions(source=source, output_format="jpeg", output=output, pages="1"))
    assert raised.value.context.code == "output_exists"
    assert output.read_bytes() == original
    export_document(ExportOptions(source=source, output_format="jpeg", output=output, pages="1", force=True))


def test_pdf_inspection_uses_stable_result_envelope(tmp_path: Path) -> None:
    source = create_blank_pdf(tmp_path / "source.pdf", [(72, 72)])
    result = inspect_document(source)
    assert result["mode"] == "inspect"
    assert result["source"]["format"] == "pdf"
    assert result["inspection"]["page_count"] == 1


def test_pdf_to_pdf_is_rejected_before_open(tmp_path: Path) -> None:
    source = tmp_path / "source.pdf"
    source.write_bytes(b"invalid but present")
    with pytest.raises(UsageError, match="PDF-to-PDF"):
        export_document(ExportOptions(source=source, output_format="pdf"))


def test_failed_conversion_can_write_a_failure_manifest(tmp_path: Path) -> None:
    source = create_blank_pdf(tmp_path / "source.pdf", [(72, 72)])
    manifest = tmp_path / "failed.json"
    with pytest.raises(UsageError, match="PDF-to-PDF"):
        export_document(ExportOptions(source=source, output_format="pdf", manifest=manifest))
    payload = json.loads(manifest.read_text(encoding="utf-8"))
    assert payload["ok"] is False
    assert payload["error"]["code"] == "usage_error"


def test_image_output_rejects_a_mismatched_file_extension(tmp_path: Path) -> None:
    source = create_blank_pdf(tmp_path / "source.pdf", [(72, 72)])
    with pytest.raises(UsageError, match=r"must end in \.png"):
        export_document(ExportOptions(source=source, output_format="png", output=tmp_path / "preview.jpg"))
