from __future__ import annotations

from pathlib import Path

from office_export.naming import (
    default_output_path,
    excel_chart_image_name,
    excel_page_image_name,
    powerpoint_image_name,
    safe_component,
    word_image_name,
)


def test_default_output_paths() -> None:
    source = Path("C:/work/My Report.docx")
    assert default_output_path(source, "pdf") == Path("C:/work/My Report.pdf")
    assert default_output_path(source, "png") == Path("C:/work/My Report - PNG export")
    assert default_output_path(source, "jpeg") == Path("C:/work/My Report - JPEG export")


def test_safe_component_handles_windows_names_and_characters() -> None:
    assert safe_component(" Q4: Revenue / Margin ") == "q4-revenue-margin"
    assert safe_component("CON") == "con-item"
    assert safe_component("...", fallback="sheet") == "sheet"


def test_deterministic_image_names() -> None:
    source = Path("Quarterly Report.docx")
    assert word_image_name(source, 2, 12, "png") == "quarterly-report-page-002.png"
    assert powerpoint_image_name(source.with_suffix(".pptx"), 3, 20, "jpeg") == "quarterly-report-slide-003.jpg"
    assert (
        excel_page_image_name(source.with_suffix(".xlsx"), "Q4 Summary", 1, 2, "png")
        == "quarterly-report-sheet-q4-summary-page-001.png"
    )
    assert (
        excel_chart_image_name(source.with_suffix(".xlsx"), "Dashboard", "Revenue %", "jpeg")
        == "quarterly-report-sheet-dashboard-chart-revenue.jpg"
    )
