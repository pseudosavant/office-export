from __future__ import annotations

from pathlib import Path

import pypdfium2 as pdfium
import pytest

from office_export import skill


@pytest.fixture(autouse=True)
def isolated_skills_dir(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """No test, including ordinary CLI calls, may touch the user's real skills."""
    root = tmp_path / "home" / ".agents" / "skills"
    monkeypatch.setattr(skill, "default_skills_dir", lambda: root)
    return root


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
