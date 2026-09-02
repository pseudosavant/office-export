from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any


class AdapterFailure(Exception):
    def __init__(
        self,
        code: str,
        message: str,
        *,
        category: str = "conversion",
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.category = category
        self.details = details


def inspect_word(document: Any) -> dict[str, Any]:
    page_count = _word_page_count(document)
    sections: list[dict[str, Any]] = []
    for index in range(1, int(document.Sections.Count) + 1):
        section = document.Sections.Item(index)
        setup = section.PageSetup
        sections.append(
            {
                "index": index,
                "width_points": _float_or_none(lambda setup=setup: setup.PageWidth),
                "height_points": _float_or_none(lambda setup=setup: setup.PageHeight),
                "orientation": _int_or_none(lambda setup=setup: setup.Orientation),
            }
        )
    return {
        "page_count": page_count,
        "title": _word_property(document, "Title"),
        "section_count": int(document.Sections.Count),
        "sections": sections,
        "has_tracked_changes": int(document.Revisions.Count) > 0,
        "has_comments": int(document.Comments.Count) > 0,
        "protection_type": _int_or_none(lambda: document.ProtectionType),
    }


def export_word_pdf(document: Any, output: Path, options: dict[str, Any]) -> dict[str, Any]:
    if options.get("update_toc", True):
        for index in range(1, int(document.TablesOfContents.Count) + 1):
            document.TablesOfContents.Item(index).Update()
    document.Repaginate()
    page_count = _word_page_count(document)
    selected = options.get("pages")
    if selected:
        _validate_number_selection(selected, page_count, "Word page")
        if selected != list(range(min(selected), max(selected) + 1)):
            raise AdapterFailure(
                "word_noncontiguous_pdf_selection",
                "Word PDF output supports only a contiguous page range.",
                category="usage",
            )
        export_range = 3
        first_page = min(selected)
        last_page = max(selected)
    else:
        export_range = 0
        first_page = 1
        last_page = page_count
    document.ExportAsFixedFormat(
        OutputFileName=str(output),
        ExportFormat=17,
        OpenAfterExport=False,
        OptimizeFor=1 if options.get("quality") == "screen" else 0,
        Range=export_range,
        From=first_page,
        To=last_page,
        Item=7 if options.get("include_markup") else 0,
        IncludeDocProps=True,
        KeepIRM=False,
        CreateBookmarks={"none": 0, "headings": 1, "word": 2}[options.get("bookmarks", "headings")],
        DocStructureTags=True,
        BitmapMissingFonts=True,
        UseISO19005_1=bool(options.get("pdf_a")),
    )
    _require_output(output, "Word")
    exported = selected or list(range(1, page_count + 1))
    return {"page_count": page_count, "exported_pages": exported}


def inspect_powerpoint(presentation: Any) -> dict[str, Any]:
    slides: list[dict[str, Any]] = []
    for index in range(1, int(presentation.Slides.Count) + 1):
        slide = presentation.Slides.Item(index)
        slides.append(
            {
                "index": index,
                "id": int(slide.SlideID),
                "name": str(slide.Name),
                "title": _powerpoint_title(slide),
                "hidden": bool(slide.SlideShowTransition.Hidden),
                "has_notes": _powerpoint_has_notes(slide),
            }
        )
    return {
        "slide_count": int(presentation.Slides.Count),
        "width_points": float(presentation.PageSetup.SlideWidth),
        "height_points": float(presentation.PageSetup.SlideHeight),
        "slides": slides,
        "title": _powerpoint_property(presentation, "Title"),
    }


def export_powerpoint_pdf(presentation: Any, output: Path, options: dict[str, Any]) -> dict[str, Any]:
    slide_count = int(presentation.Slides.Count)
    include_hidden = bool(options.get("include_hidden"))
    selected = _selected_powerpoint_slides(presentation, options.get("slides"), include_hidden)
    ranges = presentation.PrintOptions.Ranges
    ranges.ClearAll()
    print_range = None
    range_type = 1
    if selected != list(range(1, slide_count + 1)):
        selected_set = set(selected)
        for index in range(slide_count, 0, -1):
            if index not in selected_set:
                presentation.Slides.Item(index).Delete()
    output_type = {
        "slides": 1,
        "handout2": 2,
        "handout3": 3,
        "handout6": 4,
        "notes": 5,
        "outline": 6,
        "handout4": 8,
        "handout9": 9,
        "handout1": 10,
    }[options.get("output_type", "slides")]
    presentation.ExportAsFixedFormat(
        str(output),
        2,
        1 if options.get("quality") == "screen" else 2,
        bool(options.get("frame_slides")),
        1,
        output_type,
        include_hidden,
        print_range,
        range_type,
        "",
        True,
        False,
        True,
        True,
        bool(options.get("pdf_a")),
    )
    ranges.ClearAll()
    _require_output(output, "PowerPoint")
    return {"slide_count": slide_count, "exported_slides": selected}


def export_powerpoint_images(presentation: Any, output_dir: Path, options: dict[str, Any]) -> dict[str, Any]:
    if options.get("output_type", "slides") != "slides":
        raise AdapterFailure(
            "office_image_engine_output_type",
            "The PowerPoint Office image engine supports slide output only.",
            category="usage",
        )
    slide_count = int(presentation.Slides.Count)
    include_hidden = bool(options.get("include_hidden"))
    selected = _selected_powerpoint_slides(presentation, options.get("slides"), include_hidden)
    width_points = float(presentation.PageSetup.SlideWidth)
    height_points = float(presentation.PageSetup.SlideHeight)
    scale = int(options.get("dpi", 150)) / 72
    width = max(1, round(width_points * scale))
    height = max(1, round(height_points * scale))
    image_format = options["output_format"]
    extension = ".png" if image_format == "png" else ".jpg"
    filter_name = "PNG" if image_format == "png" else "JPG"
    digits = max(3, len(str(slide_count)))
    source_stem = options["source_stem"]
    outputs: list[dict[str, Any]] = []
    for slide_number in selected:
        filename = f"{source_stem}-slide-{slide_number:0{digits}d}{extension}"
        target = output_dir / filename
        presentation.Slides.Item(slide_number).Export(str(target), filter_name, width, height)
        _require_output(target, f"PowerPoint slide {slide_number}")
        outputs.append({"slide": slide_number, "filename": filename, "width": width, "height": height})
    return {"slide_count": slide_count, "exported_slides": selected, "images": outputs}


def inspect_excel(workbook: Any, application: Any) -> dict[str, Any]:
    sheets: list[dict[str, Any]] = []
    charts: list[dict[str, Any]] = []
    for position in range(1, int(workbook.Sheets.Count) + 1):
        sheet = workbook.Sheets.Item(position)
        kind = "worksheet" if _is_worksheet(sheet) else "chart_sheet"
        item: dict[str, Any] = {
            "position": position,
            "name": str(sheet.Name),
            "kind": kind,
            "visibility": _excel_visibility(_int_or_none(lambda sheet=sheet: sheet.Visible)),
        }
        if kind == "worksheet":
            item.update(
                {
                    "used_range": _excel_address(sheet.UsedRange),
                    "print_area": str(sheet.PageSetup.PrintArea or "") or None,
                    "orientation": _int_or_none(lambda sheet=sheet: sheet.PageSetup.Orientation),
                    "paper_size": _int_or_none(lambda sheet=sheet: sheet.PageSetup.PaperSize),
                    "estimated_page_count": _excel_page_count(sheet),
                }
            )
            for chart_index in range(1, int(sheet.ChartObjects().Count) + 1):
                chart_object = sheet.ChartObjects().Item(chart_index)
                chart = chart_object.Chart
                charts.append(
                    {
                        "sheet": str(sheet.Name),
                        "name": str(chart_object.Name),
                        "title": _excel_chart_title(chart),
                        "type": _int_or_none(lambda chart=chart: chart.ChartType),
                        "visible": bool(chart_object.Visible),
                        "width_points": float(chart_object.Width),
                        "height_points": float(chart_object.Height),
                        "kind": "embedded",
                    }
                )
        else:
            chart = sheet
            charts.append(
                {
                    "sheet": str(sheet.Name),
                    "name": str(sheet.Name),
                    "title": _excel_chart_title(chart),
                    "type": _int_or_none(lambda chart=chart: chart.ChartType),
                    "visible": _int_or_none(lambda sheet=sheet: sheet.Visible) == -1,
                    "width_points": _float_or_none(lambda chart=chart: chart.ChartArea.Width),
                    "height_points": _float_or_none(lambda chart=chart: chart.ChartArea.Height),
                    "kind": "chart_sheet",
                }
            )
        sheets.append(item)
    external_links = _excel_link_sources(workbook)
    return {
        "sheets": sheets,
        "charts": charts,
        "external_links": external_links,
        "has_external_links": bool(external_links),
        "connection_count": int(workbook.Connections.Count),
        "calculation_mode": _int_or_none(lambda: application.Calculation),
        "title": _excel_property(workbook, "Title"),
    }


def export_excel_pdf(workbook: Any, application: Any, output: Path, options: dict[str, Any]) -> dict[str, Any]:
    _prepare_excel_view(workbook, application, options)
    quality = 1 if options.get("quality") == "screen" else 0
    ignore_print_areas = bool(options.get("ignore_print_area"))
    range_value = options.get("range")
    selected_sheets = _resolve_excel_sheets(workbook, options.get("sheets"))
    if range_value:
        sheet_name, address = _split_excel_reference(range_value, label="range")
        sheet = _sheet_by_name(workbook, sheet_name)
        target = sheet.Range(address)
        _excel_export_pdf(target, output, quality, ignore_print_areas)
        units = [{"sheet": str(sheet.Name), "page_count": _excel_page_count(sheet), "range": address}]
    elif selected_sheets and len(selected_sheets) == 1:
        sheet = selected_sheets[0]
        _excel_export_pdf(sheet, output, quality, ignore_print_areas)
        units = [{"sheet": str(sheet.Name), "page_count": _excel_page_count(sheet)}]
    elif selected_sheets:
        names = [str(sheet.Name) for sheet in selected_sheets]
        workbook.Sheets(names).Select()
        _excel_export_pdf(application.ActiveSheet, output, quality, ignore_print_areas)
        units = [{"sheet": str(sheet.Name), "page_count": _excel_page_count(sheet)} for sheet in selected_sheets]
    else:
        _excel_export_pdf(workbook, output, quality, ignore_print_areas)
        units = [
            {"sheet": str(sheet.Name), "page_count": _excel_page_count(sheet)}
            for sheet in _visible_worksheets(workbook)
        ]
    _require_output(output, "Excel")
    return {"units": units}


def export_excel_charts(workbook: Any, output_dir: Path, options: dict[str, Any]) -> dict[str, Any]:
    available = _excel_charts(workbook)
    selectors = options.get("charts") or []
    select_all = bool(options.get("charts_all"))
    if select_all:
        selected = [item for item in available if item["visible"]]
    else:
        selected = []
        for selector in selectors:
            sheet_name, chart_name = _split_excel_reference(selector, label="chart")
            matches = [
                item
                for item in available
                if item["sheet"].casefold() == sheet_name.casefold()
                and item["name"].casefold() == chart_name.casefold()
            ]
            if not matches:
                raise AdapterFailure(
                    "excel_chart_not_found",
                    f"Excel chart was not found: {selector}",
                    category="usage",
                )
            selected.append(matches[0])
    outputs: list[dict[str, Any]] = []
    for item in selected:
        filename = item["filename"]
        target = output_dir / filename
        exported = bool(item["chart"].Export(str(target), "PNG", False))
        if not exported:
            raise AdapterFailure(
                "excel_chart_export_failed", f"Excel could not export chart {item['sheet']}!{item['name']}."
            )
        _require_output(target, f"Excel chart {item['sheet']}!{item['name']}")
        outputs.append(
            {
                "sheet": item["sheet"],
                "chart": item["name"],
                "filename": filename,
                "width_points": item["width_points"],
                "height_points": item["height_points"],
            }
        )
    for chart_item in available:
        chart_item["chart"] = None
    return {"images": outputs}


def _word_page_count(document: Any) -> int:
    return int(document.ComputeStatistics(2))


def _word_property(document: Any, name: str) -> str | None:
    return _property_value(lambda: document.BuiltInDocumentProperties(name).Value)


def _powerpoint_property(presentation: Any, name: str) -> str | None:
    return _property_value(lambda: presentation.BuiltInDocumentProperties(name).Value)


def _excel_property(workbook: Any, name: str) -> str | None:
    return _property_value(lambda: workbook.BuiltinDocumentProperties(name).Value)


def _property_value(getter: Callable[[], Any]) -> str | None:
    try:
        value = getter()
    except Exception:
        return None
    text = str(value).strip() if value is not None else ""
    return text or None


def _powerpoint_title(slide: Any) -> str | None:
    try:
        if not slide.Shapes.HasTitle:
            return None
        text = str(slide.Shapes.Title.TextFrame.TextRange.Text).strip()
        return text or None
    except Exception:
        return None


def _powerpoint_has_notes(slide: Any) -> bool:
    try:
        shapes = slide.NotesPage.Shapes
        for index in range(1, int(shapes.Count) + 1):
            shape = shapes.Item(index)
            if bool(shape.HasTextFrame) and bool(shape.TextFrame.HasText):
                text = str(shape.TextFrame.TextRange.Text).strip()
                if text and text not in {str(slide.SlideIndex), ""}:
                    return True
    except Exception:
        return False
    return False


def _is_worksheet(sheet: Any) -> bool:
    try:
        _ = sheet.UsedRange
        return True
    except Exception:
        return False


def _excel_visibility(value: int | None) -> str:
    return {-1: "visible", 0: "hidden", 2: "very_hidden"}.get(value, "unknown")


def _excel_address(cell_range: Any) -> str | None:
    try:
        return str(cell_range.Address(True, True, 1, False))
    except Exception:
        try:
            return str(cell_range.Address)
        except Exception:
            return None


def _excel_page_count(sheet: Any) -> int | None:
    try:
        return max(1, (int(sheet.HPageBreaks.Count) + 1) * (int(sheet.VPageBreaks.Count) + 1))
    except Exception:
        return None


def _excel_chart_title(chart: Any) -> str | None:
    try:
        if not bool(chart.HasTitle):
            return None
        text = str(chart.ChartTitle.Text).strip()
        return text or None
    except Exception:
        return None


def _excel_link_sources(workbook: Any) -> list[str]:
    try:
        values = workbook.LinkSources(1)
    except Exception:
        return []
    if values is None:
        return []
    if isinstance(values, str):
        return [values]
    try:
        return [str(value) for value in values]
    except TypeError:
        return [str(values)]


def _resolve_excel_sheets(workbook: Any, selectors: list[int | str] | None) -> list[Any]:
    if not selectors:
        return []
    resolved: list[Any] = []
    for selector in selectors:
        try:
            sheet = workbook.Sheets.Item(selector)
        except Exception as exc:
            raise AdapterFailure(
                "excel_sheet_not_found",
                f"Excel sheet was not found: {selector}",
                category="usage",
            ) from exc
        if any(str(existing.Name).casefold() == str(sheet.Name).casefold() for existing in resolved):
            raise AdapterFailure(
                "excel_sheet_duplicate",
                f"Excel sheet was selected more than once: {sheet.Name}",
                category="usage",
            )
        resolved.append(sheet)
    return resolved


def _visible_worksheets(workbook: Any) -> list[Any]:
    values: list[Any] = []
    for index in range(1, int(workbook.Worksheets.Count) + 1):
        sheet = workbook.Worksheets.Item(index)
        if _int_or_none(lambda sheet=sheet: sheet.Visible) == -1:
            values.append(sheet)
    return values


def _sheet_by_name(workbook: Any, name: str) -> Any:
    try:
        sheet = workbook.Worksheets.Item(name)
    except Exception as exc:
        raise AdapterFailure(
            "excel_sheet_not_found", f"Excel worksheet was not found: {name}", category="usage"
        ) from exc
    return sheet


def _split_excel_reference(value: str, *, label: str) -> tuple[str, str]:
    sheet, separator, reference = value.partition("!")
    if not separator or not sheet.strip() or not reference.strip():
        raise AdapterFailure(
            f"excel_{label}_syntax",
            f"Excel {label} must use SHEET!VALUE syntax.",
            category="usage",
        )
    sheet = sheet.strip()
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, reference.strip()


def _prepare_excel_view(workbook: Any, application: Any, options: dict[str, Any]) -> None:
    recalculate = options.get("recalculate", "never")
    if recalculate == "auto":
        application.Calculate()
    elif recalculate == "full":
        application.CalculateFullRebuild()
    if options.get("show_formulas"):
        for index in range(1, int(workbook.Worksheets.Count) + 1):
            workbook.Worksheets.Item(index).Activate()
            application.ActiveWindow.DisplayFormulas = True
    if options.get("show_headings"):
        for index in range(1, int(workbook.Worksheets.Count) + 1):
            workbook.Worksheets.Item(index).PageSetup.PrintHeadings = True


def _excel_charts(workbook: Any) -> list[dict[str, Any]]:
    from office_export.naming import excel_chart_image_name

    source = Path(str(workbook.FullName))
    items: list[dict[str, Any]] = []
    for sheet_index in range(1, int(workbook.Worksheets.Count) + 1):
        sheet = workbook.Worksheets.Item(sheet_index)
        for chart_index in range(1, int(sheet.ChartObjects().Count) + 1):
            chart_object = sheet.ChartObjects().Item(chart_index)
            items.append(
                {
                    "sheet": str(sheet.Name),
                    "name": str(chart_object.Name),
                    "chart": chart_object.Chart,
                    "visible": bool(chart_object.Visible) and _int_or_none(lambda sheet=sheet: sheet.Visible) == -1,
                    "width_points": float(chart_object.Width),
                    "height_points": float(chart_object.Height),
                    "filename": excel_chart_image_name(source, str(sheet.Name), str(chart_object.Name), "png"),
                }
            )
    for index in range(1, int(workbook.Charts.Count) + 1):
        chart = workbook.Charts.Item(index)
        items.append(
            {
                "sheet": str(chart.Name),
                "name": str(chart.Name),
                "chart": chart,
                "visible": _int_or_none(lambda chart=chart: chart.Visible) == -1,
                "width_points": _float_or_none(lambda chart=chart: chart.ChartArea.Width),
                "height_points": _float_or_none(lambda chart=chart: chart.ChartArea.Height),
                "filename": excel_chart_image_name(source, str(chart.Name), str(chart.Name), "png"),
            }
        )
    return items


def _require_output(path: Path, producer: str) -> None:
    if not path.is_file() or path.stat().st_size == 0:
        raise AdapterFailure("office_output_missing", f"{producer} did not create the expected output file.")


def _excel_export_pdf(target: Any, output: Path, quality: int, ignore_print_areas: bool) -> None:
    target.ExportAsFixedFormat(
        Type=0,
        Filename=str(output),
        Quality=quality,
        IncludeDocProperties=True,
        IgnorePrintAreas=ignore_print_areas,
        OpenAfterPublish=False,
    )


def _selected_powerpoint_slides(presentation: Any, requested: list[int] | None, include_hidden: bool) -> list[int]:
    slide_count = int(presentation.Slides.Count)
    if requested is not None:
        _validate_number_selection(requested, slide_count, "PowerPoint slide")
    candidates = requested if requested is not None else list(range(1, slide_count + 1))
    selected = [
        index
        for index in candidates
        if include_hidden or not bool(presentation.Slides.Item(index).SlideShowTransition.Hidden)
    ]
    if not selected:
        raise AdapterFailure(
            "no_visible_slides_selected",
            "No visible PowerPoint slides were selected. Use --include-hidden to export hidden slides.",
            category="usage",
        )
    return selected


def _validate_number_selection(numbers: list[int], count: int, label: str) -> None:
    outside = [number for number in numbers if number < 1 or number > count]
    if outside:
        raise AdapterFailure(
            "selection_out_of_range",
            f"Selected {label} {outside[0]} is outside the available range of 1 through {count}.",
            category="usage",
        )


def _float_or_none(getter: Callable[[], Any]) -> float | None:
    try:
        return float(getter())
    except Exception:
        return None


def _int_or_none(getter: Callable[[], Any]) -> int | None:
    try:
        return int(getter())
    except Exception:
        return None
