from __future__ import annotations

import gc
import sys
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

import pytest

from office_export.core import ExportOptions, export_document, inspect_document
from office_export.results import sha256_file
from office_export.worker import _application_pid, _snapshot_processes

pytestmark = pytest.mark.office


@contextmanager
def _office_application(progid: str, process_name: str) -> Iterator[Any]:
    if sys.platform != "win32":
        pytest.skip("Office integration tests require Windows.")
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    application = None
    owned = False
    try:
        before = _snapshot_processes(process_name)
        application = win32com.client.DispatchEx(progid)
        after = _snapshot_processes(process_name)
        created = after - before
        pid = _application_pid(application)
        if pid is None and len(created) == 1:
            pid = next(iter(created))
        owned = pid is not None and pid in created
        yield application
    finally:
        if application is not None and owned:
            application.Quit()
        application = None
        gc.collect()
        pythoncom.CoUninitialize()


@pytest.fixture(scope="module")
def office_files(tmp_path_factory: pytest.TempPathFactory) -> dict[str, Path]:
    root = tmp_path_factory.mktemp("office-files")
    files: dict[str, Path] = {}

    with _office_application("Word.Application", "WINWORD.EXE") as word:
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        document = word.Documents.Add()
        try:
            document.Content.Text = "Office Export Integration\rThis document verifies native Word PDF output."
            files["docx"] = root / "word-modern.docx"
            document.SaveAs2(str(files["docx"]), FileFormat=16)
            files["doc"] = root / "word-legacy.doc"
            document.SaveAs2(str(files["doc"]), FileFormat=0)
        finally:
            document.Close(SaveChanges=0)
            document = None

    with _office_application("Excel.Application", "EXCEL.EXE") as excel:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        workbook = excel.Workbooks.Add()
        try:
            summary = workbook.Worksheets.Item(1)
            summary.Name = "Summary"
            summary.Range("A1").Value = "Quarter"
            summary.Range("B1").Value = "Revenue"
            summary.Range("A2").Value = "Q1"
            summary.Range("B2").Value = 100
            summary.Range("A3").Value = "Q2"
            summary.Range("B3").Value = 140
            chart_object = summary.ChartObjects().Add(200, 20, 320, 180)
            chart_object.Name = "Revenue Chart"
            chart_object.Chart.SetSourceData(summary.Range("A1:B3"))
            chart_object.Chart.HasTitle = True
            chart_object.Chart.ChartTitle.Text = "Revenue"
            details = workbook.Worksheets.Add(After=summary)
            details.Name = "Details"
            details.Range("A1").Value = "Detail row"
            files["xlsx"] = root / "excel-modern.xlsx"
            workbook.SaveAs(str(files["xlsx"]), FileFormat=51)
            files["xls"] = root / "excel-legacy.xls"
            workbook.SaveAs(str(files["xls"]), FileFormat=56)
        finally:
            workbook.Close(SaveChanges=False)
            chart_object = None
            summary = None
            details = None
            workbook = None

    with _office_application("PowerPoint.Application", "POWERPNT.EXE") as powerpoint:
        powerpoint.DisplayAlerts = 1
        powerpoint.AutomationSecurity = 3
        presentation = powerpoint.Presentations.Add(WithWindow=False)
        try:
            first = presentation.Slides.Add(1, 2)
            first.Shapes.Title.TextFrame.TextRange.Text = "Office Export"
            first.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Native PowerPoint rendering"
            second = presentation.Slides.Add(2, 2)
            second.Shapes.Title.TextFrame.TextRange.Text = "Hidden slide"
            second.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Selection fixture"
            second.SlideShowTransition.Hidden = True
            third = presentation.Slides.Add(3, 2)
            third.Shapes.Title.TextFrame.TextRange.Text = "Visible closing slide"
            third.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Noncontiguous selection fixture"
            files["pptx"] = root / "powerpoint-modern.pptx"
            presentation.SaveAs(str(files["pptx"]), 24)
            files["ppt"] = root / "powerpoint-legacy.ppt"
            presentation.SaveCopyAs(str(files["ppt"]), 1)
        finally:
            presentation.Close()
            first = None
            second = None
            third = None
            presentation = None
    return files


@pytest.mark.parametrize("extension", ["docx", "doc"])
def test_word_modern_and_legacy_native_pdf_preserve_source(
    office_files: dict[str, Path], tmp_path: Path, extension: str
) -> None:
    source = office_files[extension]
    before = sha256_file(source)
    inspection = inspect_document(source, timeout=60)
    assert inspection["inspection"]["page_count"] >= 1
    output = tmp_path / f"word-{extension}.pdf"
    result = export_document(ExportOptions(source=source, output_format="pdf", output=output, timeout=60))
    assert result["office"]["id"] == "word"
    assert result["outputs"][0]["size"] > 0
    assert sha256_file(source) == before


def test_word_pdfium_image_path(office_files: dict[str, Path], tmp_path: Path) -> None:
    source = office_files["docx"]
    output = tmp_path / "word-images"
    result = export_document(
        ExportOptions(source=source, output_format="png", output=output, pages="1", dpi=96, timeout=60)
    )
    assert result["mapping"][0]["page"] == 1
    assert Path(result["outputs"][0]["path"]).suffix == ".png"


@pytest.mark.parametrize("extension", ["xlsx", "xls"])
def test_excel_modern_and_legacy_native_pdf_preserve_source(
    office_files: dict[str, Path], tmp_path: Path, extension: str
) -> None:
    source = office_files[extension]
    before = sha256_file(source)
    inspection = inspect_document(source, timeout=60)
    assert {sheet["name"] for sheet in inspection["inspection"]["sheets"]} == {"Summary", "Details"}
    assert inspection["inspection"]["charts"][0]["name"] == "Revenue Chart"
    output = tmp_path / f"excel-{extension}.pdf"
    result = export_document(
        ExportOptions(source=source, output_format="pdf", output=output, sheets=["Summary"], timeout=60)
    )
    assert result["office"]["id"] == "excel"
    assert result["outputs"][0]["size"] > 0
    assert sha256_file(source) == before


def test_excel_chart_png_and_jpeg(office_files: dict[str, Path], tmp_path: Path) -> None:
    source = office_files["xlsx"]
    png = export_document(
        ExportOptions(
            source=source,
            output_format="png",
            output=tmp_path / "chart-png",
            charts=["Summary!Revenue Chart"],
            timeout=60,
        )
    )
    assert png["mapping"][0]["chart"] == "Revenue Chart"
    assert Path(png["outputs"][0]["path"]).suffix == ".png"
    jpeg = export_document(
        ExportOptions(
            source=source,
            output_format="jpeg",
            output=tmp_path / "chart-jpeg",
            charts_all=True,
            jpeg_quality=85,
            timeout=60,
        )
    )
    assert Path(jpeg["outputs"][0]["path"]).suffix == ".jpg"


@pytest.mark.parametrize("extension", ["pptx", "ppt"])
def test_powerpoint_modern_and_legacy_native_pdf_preserve_source(
    office_files: dict[str, Path], tmp_path: Path, extension: str
) -> None:
    source = office_files[extension]
    before = sha256_file(source)
    inspection = inspect_document(source, timeout=60)
    assert inspection["inspection"]["slide_count"] == 3
    assert inspection["inspection"]["slides"][1]["hidden"] is True
    output = tmp_path / f"powerpoint-{extension}.pdf"
    result = export_document(ExportOptions(source=source, output_format="pdf", output=output, slides="1,3", timeout=60))
    assert result["mapping"][0]["slides"] == [1, 3]
    assert result["outputs"][0]["size"] > 0
    assert inspect_document(output)["inspection"]["page_count"] == 2
    assert sha256_file(source) == before


def test_powerpoint_both_image_engines(office_files: dict[str, Path], tmp_path: Path) -> None:
    source = office_files["pptx"]
    native = export_document(
        ExportOptions(
            source=source,
            output_format="png",
            output=tmp_path / "native",
            slides="1",
            image_engine="office",
            dpi=96,
            timeout=60,
        )
    )
    assert native["mapping"][0]["engine"] == "office"
    pdfium = export_document(
        ExportOptions(
            source=source,
            output_format="jpeg",
            output=tmp_path / "pdfium",
            slides="1",
            dpi=96,
            timeout=60,
        )
    )
    assert pdfium["mapping"][0]["slide"] == 1
    assert Path(pdfium["outputs"][0]["path"]).suffix == ".jpg"
