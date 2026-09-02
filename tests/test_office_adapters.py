from __future__ import annotations

from pathlib import Path
from typing import Any

from office_export.office_adapters import export_word_pdf


class FakeCollection:
    Count = 0


class FakeWordDocument:
    TablesOfContents = FakeCollection()

    def __init__(self, output: Path) -> None:
        self.output = output
        self.repaginate_called = False
        self.export_options: dict[str, Any] | None = None

    def Repaginate(self) -> None:
        self.repaginate_called = True

    def ComputeStatistics(self, statistic: int) -> int:
        assert statistic == 2
        return 5

    def ExportAsFixedFormat(self, **kwargs: Any) -> None:
        self.export_options = kwargs
        self.output.write_bytes(b"pdf")


def test_word_pdf_uses_native_range_and_fidelity_options(tmp_path: Path) -> None:
    output = tmp_path / "word.pdf"
    document = FakeWordDocument(output)
    result = export_word_pdf(
        document,
        output,
        {
            "pages": [2, 3],
            "quality": "print",
            "bookmarks": "headings",
            "include_markup": False,
            "pdf_a": False,
            "update_toc": True,
        },
    )
    assert result == {"page_count": 5, "exported_pages": [2, 3]}
    assert document.repaginate_called is True
    assert document.export_options is not None
    assert document.export_options["Range"] == 3
    assert document.export_options["From"] == 2
    assert document.export_options["To"] == 3
    assert document.export_options["IncludeDocProps"] is True
    assert document.export_options["DocStructureTags"] is True
    assert document.export_options["KeepIRM"] is False
