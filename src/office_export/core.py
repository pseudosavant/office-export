from __future__ import annotations

import json
import shutil
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from PIL import Image, ImageColor

from office_export.errors import OfficeExportError, UsageError
from office_export.naming import (
    IMAGE_SUFFIXES,
    default_output_path,
    ensure_distinct_paths,
    excel_page_image_name,
    powerpoint_image_name,
    resolve_output_path,
    safe_component,
    validate_source,
    word_image_name,
)
from office_export.publishing import publish_file, publish_images, staging_directory
from office_export.rasterizer import (
    DEFAULT_DPI,
    DEFAULT_JPEG_QUALITY,
    DEFAULT_MAX_MEGAPIXELS,
    RasterPage,
    inspect_pdf,
    rasterize_pdf,
    verify_unprotected_pdf,
)
from office_export.results import base_result, file_record, runtime_versions, source_record, write_manifest
from office_export.selectors import is_contiguous, parse_number_selection, parse_sheet_selector
from office_export.worker_protocol import run_worker

SOURCE_FORMATS = {
    ".docx": ("word", "docx"),
    ".doc": ("word", "doc"),
    ".xlsx": ("excel", "xlsx"),
    ".xls": ("excel", "xls"),
    ".pptx": ("powerpoint", "pptx"),
    ".ppt": ("powerpoint", "ppt"),
    ".pdf": ("pdf", "pdf"),
}


@dataclass(slots=True)
class ExportOptions:
    source: Path
    output_format: str
    output: Path | None = None
    force: bool = False
    dpi: int | None = None
    jpeg_quality: int | None = None
    background: str | None = None
    quality: str = "print"
    timeout: float = 120.0
    keep_intermediate: bool = False
    manifest: Path | None = None
    verbose: bool = False
    image_engine: str = "pdfium"
    pages: str | None = None
    include_markup: bool = False
    bookmarks: str = "headings"
    pdf_a: bool = False
    update_toc: bool = True
    slides: str | None = None
    include_hidden: bool = False
    output_type: str = "slides"
    frame_slides: bool = False
    sheets: list[str] = field(default_factory=list)
    range_value: str | None = None
    charts: list[str] = field(default_factory=list)
    charts_all: bool = False
    ignore_print_area: bool = False
    show_formulas: bool = False
    show_headings: bool = False
    recalculate: str = "never"
    update_links: bool = False
    refresh_data: bool = False
    exclude_annotations: bool = False
    max_megapixels: float = DEFAULT_MAX_MEGAPIXELS

    def manifest_options(self) -> dict[str, Any]:
        return {
            "to": self.output_format,
            "output": str(self.output.resolve()) if self.output else None,
            "dpi": self.dpi if self.dpi is not None else DEFAULT_DPI,
            "jpeg_quality": self.jpeg_quality if self.jpeg_quality is not None else DEFAULT_JPEG_QUALITY,
            "background": self.background or "white",
            "quality": self.quality,
            "timeout": self.timeout,
            "keep_intermediate": self.keep_intermediate,
            "image_engine": self.image_engine,
            "pages": self.pages,
            "include_markup": self.include_markup,
            "bookmarks": self.bookmarks,
            "pdf_a": self.pdf_a,
            "update_toc": self.update_toc,
            "slides": self.slides,
            "include_hidden": self.include_hidden,
            "output_type": self.output_type,
            "frame_slides": self.frame_slides,
            "sheets": self.sheets,
            "range": self.range_value,
            "charts": self.charts,
            "charts_all": self.charts_all,
            "ignore_print_area": self.ignore_print_area,
            "show_formulas": self.show_formulas,
            "show_headings": self.show_headings,
            "recalculate": self.recalculate,
            "update_links": self.update_links,
            "refresh_data": self.refresh_data,
            "exclude_annotations": self.exclude_annotations,
            "max_megapixels": self.max_megapixels,
        }


def detect_source(path: Path) -> tuple[str, str]:
    suffix = path.suffix.lower()
    if suffix not in SOURCE_FORMATS:
        supported = ", ".join(sorted(SOURCE_FORMATS))
        raise UsageError(
            f"Unsupported source format '{suffix or '(none)'}'. Supported formats: {supported}.",
            code="unsupported_source_format",
        )
    return SOURCE_FORMATS[suffix]


def inspect_document(source: Path, *, timeout: float = 120.0) -> dict[str, Any]:
    path = validate_source(source)
    application, source_format = detect_source(path)
    started = time.perf_counter()
    payload = base_result(mode="inspect", source=path, source_format=source_format)
    if application == "pdf":
        detail = inspect_pdf(path)
    else:
        response = run_worker(
            {
                "action": "inspect",
                "application": application,
                "source": str(path),
                "options": _safe_worker_options(application),
            },
            timeout,
        )
        detail = response["result"]
        detail["application"] = response["application"]
        detail["warnings"] = response.get("warnings", [])
    payload["inspection"] = detail
    payload["duration_seconds"] = round(time.perf_counter() - started, 3)
    return payload


def export_document(options: ExportOptions) -> dict[str, Any]:
    started = time.perf_counter()
    try:
        return _export_document(options, started)
    except OfficeExportError as exc:
        _write_failure_manifest(options, exc, started)
        raise


def _export_document(options: ExportOptions, started: float) -> dict[str, Any]:
    source = validate_source(options.source)
    application, source_format = detect_source(source)
    _validate_options(application, options)
    destination = resolve_output_path(source, options.output_format, options.output)
    ensure_distinct_paths(source, destination)
    _preflight_destination(destination, options)
    _preflight_manifest(source, destination, options)
    payload = base_result(mode="export", source=source, source_format=source_format)
    payload["options"] = options.manifest_options()
    payload["warnings"] = _option_warnings(application, options)
    payload["office"] = None
    payload["dependencies"] = runtime_versions()
    engine, temporary_pdf_used = _rendering_details(application, options)
    payload["rendering"] = {
        "engine": engine,
        "temporary_pdf_used": temporary_pdf_used,
    }

    parent = destination.parent
    with staging_directory(parent) as staging:
        if application == "pdf":
            outputs, mappings, warnings = _export_pdf_source(source, staging, options)
            published = publish_images(outputs, destination, output_format=options.output_format, force=options.force)
            payload["warnings"].extend(warnings)
        elif application == "excel" and (options.charts or options.charts_all):
            outputs, mappings, response = _export_excel_charts(source, staging, options)
            published = publish_images(outputs, destination, output_format=options.output_format, force=options.force)
            payload["office"] = response["application"]
            payload["warnings"].extend(response.get("warnings", []))
        elif (
            application == "powerpoint"
            and options.output_format in {"png", "jpeg"}
            and options.image_engine == "office"
        ):
            outputs, mappings, response = _export_powerpoint_native_images(source, staging, options)
            published = publish_images(outputs, destination, output_format=options.output_format, force=options.force)
            payload["office"] = response["application"]
            payload["warnings"].extend(response.get("warnings", []))
        else:
            published, mappings, response, intermediate = _export_office_via_pdf(
                source,
                application,
                staging,
                destination,
                options,
            )
            payload["office"] = response["application"]
            payload["warnings"].extend(response.get("warnings", []))
            if intermediate is not None:
                payload["intermediate_pdf"] = file_record(intermediate)

        final_by_name = {path.name: path for path in published}
        for mapping in mappings:
            staging_output = Path(str(mapping["output"]))
            if len(published) == 1 and destination.suffix.lower() in {".png", ".jpg", ".jpeg", ".pdf"}:
                mapping["output"] = str(published[0])
            elif staging_output.name in final_by_name:
                mapping["output"] = str(final_by_name[staging_output.name])
        payload["outputs"] = [file_record(path) for path in published]
        for mapping in mappings:
            mapping.setdefault("engine", engine)
        payload["mapping"] = mappings

    payload["duration_seconds"] = round(time.perf_counter() - started, 3)
    if options.manifest:
        payload["manifest"] = str(options.manifest.expanduser().resolve())
        write_manifest(options.manifest, payload, force=options.force)
    return payload


def format_capabilities() -> dict[str, Any]:
    return {
        "inputs": {
            "docx": {"application": "word", "outputs": ["pdf", "png", "jpeg"]},
            "doc": {"application": "word", "outputs": ["pdf", "png", "jpeg"]},
            "xlsx": {"application": "excel", "outputs": ["pdf", "png", "jpeg"], "chart_images": True},
            "xls": {"application": "excel", "outputs": ["pdf", "png", "jpeg"], "chart_images": True},
            "pptx": {
                "application": "powerpoint",
                "outputs": ["pdf", "png", "jpeg"],
                "image_engines": ["pdfium", "office"],
                "office_image_engine_output_types": ["slides"],
            },
            "ppt": {
                "application": "powerpoint",
                "outputs": ["pdf", "png", "jpeg"],
                "image_engines": ["pdfium", "office"],
                "office_image_engine_output_types": ["slides"],
            },
            "pdf": {"application": None, "outputs": ["png", "jpeg"]},
        },
        "defaults": {
            "image_engine": "pdfium",
            "dpi": DEFAULT_DPI,
            "jpeg_quality": DEFAULT_JPEG_QUALITY,
            "jpeg_background": "white",
            "max_megapixels": DEFAULT_MAX_MEGAPIXELS,
            "bookmarks": "headings",
        },
    }


def _export_pdf_source(
    source: Path,
    staging: Path,
    options: ExportOptions,
) -> tuple[list[Path], list[dict[str, Any]], list[dict[str, str]]]:
    inspection = inspect_pdf(source)
    page_count = int(inspection["page_count"])
    selected = parse_number_selection(options.pages, label="pages", count=page_count) or list(range(1, page_count + 1))
    image_dir = staging / "images"
    image_dir.mkdir()
    pages = [
        RasterPage(
            source_page=page,
            filename=word_image_name(source, page, page_count, options.output_format),
            logical={"page": page, "page_label": inspection["pages"][page - 1]["label"]},
        )
        for page in selected
    ]
    outputs, mappings, warnings = rasterize_pdf(
        source,
        image_dir,
        pages,
        output_format=options.output_format,
        dpi=options.dpi or DEFAULT_DPI,
        jpeg_quality=options.jpeg_quality or DEFAULT_JPEG_QUALITY,
        background=options.background or "white",
        exclude_annotations=options.exclude_annotations,
        max_megapixels=options.max_megapixels,
    )
    return outputs, mappings, warnings


def _export_excel_charts(
    source: Path,
    staging: Path,
    options: ExportOptions,
) -> tuple[list[Path], list[dict[str, Any]], dict[str, Any]]:
    native_dir = staging / "native-charts"
    native_dir.mkdir()
    response = run_worker(
        {
            "action": "export_charts",
            "application": "excel",
            "source": str(source),
            "output_dir": str(native_dir),
            "options": {
                **_safe_worker_options("excel", options),
                "charts": options.charts,
                "charts_all": options.charts_all,
            },
        },
        options.timeout,
    )
    native = [native_dir / item["filename"] for item in response["result"]["images"]]
    if options.output_format == "png":
        outputs = native
    else:
        outputs = _convert_chart_jpegs(native, options)
    mappings = []
    for item, output in zip(response["result"]["images"], outputs, strict=True):
        mappings.append(
            {
                "sheet": item["sheet"],
                "chart": item["chart"],
                "output": str(output),
                "width_points": item["width_points"],
                "height_points": item["height_points"],
            }
        )
    return outputs, mappings, response


def _convert_chart_jpegs(native: list[Path], options: ExportOptions) -> list[Path]:
    quality = options.jpeg_quality or DEFAULT_JPEG_QUALITY
    try:
        background = ImageColor.getcolor(options.background or "white", "RGB")
    except ValueError as exc:
        raise UsageError(f"Invalid background color '{options.background}'.") from exc
    outputs: list[Path] = []
    for source in native:
        target = source.with_suffix(".jpg")
        with Image.open(source) as image:
            rgba = image.convert("RGBA")
            flattened = Image.new("RGB", rgba.size, background)
            flattened.paste(rgba, mask=rgba.getchannel("A"))
            flattened.save(target, format="JPEG", quality=quality)
            flattened.close()
            rgba.close()
        outputs.append(target)
    return outputs


def _export_powerpoint_native_images(
    source: Path,
    staging: Path,
    options: ExportOptions,
) -> tuple[list[Path], list[dict[str, Any]], dict[str, Any]]:
    image_dir = staging / "native-images"
    image_dir.mkdir()
    selected = parse_number_selection(options.slides, label="slides", count=None)
    response = run_worker(
        {
            "action": "export_images",
            "application": "powerpoint",
            "source": str(source),
            "output_dir": str(image_dir),
            "options": {
                **_safe_worker_options("powerpoint", options),
                "slides": selected,
                "source_stem": safe_component(source.stem),
                "output_format": options.output_format,
                "dpi": options.dpi or DEFAULT_DPI,
            },
        },
        options.timeout,
    )
    outputs = [image_dir / item["filename"] for item in response["result"]["images"]]
    mappings = [
        {
            "slide": item["slide"],
            "output": str(image_dir / item["filename"]),
            "width": item["width"],
            "height": item["height"],
            "engine": "office",
        }
        for item in response["result"]["images"]
    ]
    return outputs, mappings, response


def _export_office_via_pdf(
    source: Path,
    application: str,
    staging: Path,
    destination: Path,
    options: ExportOptions,
) -> tuple[list[Path], list[dict[str, Any]], dict[str, Any], Path | None]:
    temporary_pdf = staging / f"{safe_component(source.stem)}.pdf"
    worker_options = _safe_worker_options(application, options)
    logical_selection: list[int] | None = None
    if application == "word":
        logical_selection = parse_number_selection(options.pages, label="pages", count=None)
        worker_options["pages"] = logical_selection if options.output_format == "pdf" else None
    elif application == "powerpoint":
        logical_selection = parse_number_selection(options.slides, label="slides", count=None)
        worker_options["slides"] = logical_selection
    response = run_worker(
        {
            "action": "export_pdf",
            "application": application,
            "source": str(source),
            "output_file": str(temporary_pdf),
            "options": worker_options,
        },
        options.timeout,
    )
    verify_unprotected_pdf(temporary_pdf)
    if options.output_format == "pdf":
        published = [publish_file(temporary_pdf, destination, force=options.force)]
        mappings = _pdf_output_mapping(application, response["result"], published[0])
        return published, mappings, response, None

    pdf_inspection = inspect_pdf(temporary_pdf)
    pages = _office_raster_pages(
        source,
        application,
        response["result"],
        pdf_inspection,
        options,
        logical_selection,
    )
    image_dir = staging / "images"
    image_dir.mkdir()
    outputs, mappings, warnings = rasterize_pdf(
        temporary_pdf,
        image_dir,
        pages,
        output_format=options.output_format,
        dpi=options.dpi or DEFAULT_DPI,
        jpeg_quality=options.jpeg_quality or DEFAULT_JPEG_QUALITY,
        background=options.background or "white",
        exclude_annotations=False,
        max_megapixels=options.max_megapixels,
    )
    response.setdefault("warnings", []).extend(warnings)
    published = publish_images(outputs, destination, output_format=options.output_format, force=options.force)
    intermediate: Path | None = None
    if options.keep_intermediate:
        intermediate_target = _intermediate_target(destination, source)
        staged_copy = staging / f"{safe_component(source.stem)}-intermediate-copy.pdf"
        shutil.copy2(temporary_pdf, staged_copy)
        intermediate = publish_file(staged_copy, intermediate_target, force=options.force)
    return published, mappings, response, intermediate


def _office_raster_pages(
    source: Path,
    application: str,
    result: dict[str, Any],
    inspection: dict[str, Any],
    options: ExportOptions,
    logical_selection: list[int] | None,
) -> list[RasterPage]:
    pdf_count = int(inspection["page_count"])
    if application == "word":
        source_page_count = int(result["page_count"])
        selected = logical_selection or list(range(1, source_page_count + 1))
        return [
            RasterPage(
                source_page=page,
                filename=word_image_name(source, page, source_page_count, options.output_format),
                logical={"page": page},
            )
            for page in selected
        ]
    if application == "powerpoint" and options.output_type in {"slides", "notes"}:
        selected = result["exported_slides"]
        if len(selected) == pdf_count:
            return [
                RasterPage(
                    source_page=pdf_page,
                    filename=powerpoint_image_name(
                        source,
                        slide,
                        int(result["slide_count"]),
                        options.output_format,
                    ),
                    logical={"slide": slide, "output_type": options.output_type},
                )
                for pdf_page, slide in enumerate(selected, start=1)
            ]
    if application == "excel":
        mapped = _excel_pages_from_units(source, result.get("units", []), pdf_count, options.output_format)
        if mapped:
            return mapped
    digits = max(3, len(str(pdf_count)))
    return [
        RasterPage(
            source_page=page,
            filename=f"{safe_component(source.stem)}-page-{page:0{digits}d}{IMAGE_SUFFIXES[options.output_format]}",
            logical={"page": page, "output_type": options.output_type},
        )
        for page in range(1, pdf_count + 1)
    ]


def _excel_pages_from_units(
    source: Path,
    units: list[dict[str, Any]],
    pdf_count: int,
    output_format: str,
) -> list[RasterPage] | None:
    counts = [unit.get("page_count") for unit in units]
    if len(units) == 1:
        unit = units[0]
        return [
            RasterPage(
                source_page=page,
                filename=excel_page_image_name(source, unit["sheet"], page, pdf_count, output_format),
                logical={"sheet": unit["sheet"], "sheet_page": page, **_optional_range(unit)},
            )
            for page in range(1, pdf_count + 1)
        ]
    if not counts or any(not isinstance(count, int) or count < 1 for count in counts) or sum(counts) != pdf_count:
        return None
    pages: list[RasterPage] = []
    pdf_page = 1
    for unit, count in zip(units, counts, strict=True):
        for sheet_page in range(1, count + 1):
            pages.append(
                RasterPage(
                    source_page=pdf_page,
                    filename=excel_page_image_name(source, unit["sheet"], sheet_page, count, output_format),
                    logical={"sheet": unit["sheet"], "sheet_page": sheet_page, **_optional_range(unit)},
                )
            )
            pdf_page += 1
    return pages


def _optional_range(unit: dict[str, Any]) -> dict[str, Any]:
    return {"range": unit["range"]} if unit.get("range") else {}


def _pdf_output_mapping(application: str, result: dict[str, Any], output: Path) -> list[dict[str, Any]]:
    if application == "word":
        return [{"pages": result["exported_pages"], "output": str(output)}]
    if application == "powerpoint":
        return [{"slides": result["exported_slides"], "output": str(output)}]
    return [{"units": result.get("units", []), "output": str(output)}]


def _intermediate_target(destination: Path, source: Path) -> Path:
    if destination.suffix.lower() in {".png", ".jpg", ".jpeg"}:
        return destination.with_name(f"{destination.stem}-intermediate.pdf")
    return destination / f"{safe_component(source.stem)}-intermediate.pdf"


def _safe_worker_options(application: str, options: ExportOptions | None = None) -> dict[str, Any]:
    if options is None:
        return {
            "update_links": False,
            "refresh_data": False,
            "recalculate": "never",
            "update_toc": False,
        }
    common = {"quality": options.quality}
    if application == "word":
        common.update(
            {
                "include_markup": options.include_markup,
                "bookmarks": options.bookmarks,
                "pdf_a": options.pdf_a,
                "update_toc": options.update_toc,
            }
        )
    elif application == "powerpoint":
        common.update(
            {
                "include_hidden": options.include_hidden,
                "output_type": options.output_type,
                "frame_slides": options.frame_slides,
                "pdf_a": options.pdf_a,
            }
        )
    elif application == "excel":
        common.update(
            {
                "sheets": [parse_sheet_selector(value) for value in options.sheets],
                "range": options.range_value,
                "ignore_print_area": options.ignore_print_area,
                "show_formulas": options.show_formulas,
                "show_headings": options.show_headings,
                "recalculate": options.recalculate,
                "update_links": options.update_links,
                "refresh_data": options.refresh_data,
            }
        )
    return common


def _validate_options(application: str, options: ExportOptions) -> None:
    if options.output_format not in {"pdf", "png", "jpeg"}:
        raise UsageError("--to must be pdf, png, or jpeg.")
    if options.timeout <= 0:
        raise UsageError("--timeout must be greater than zero.")
    if options.max_megapixels <= 0:
        raise UsageError("--max-megapixels must be greater than zero.")
    if options.dpi is not None and (options.dpi < 36 or options.dpi > 2400):
        raise UsageError("--dpi must be between 36 and 2400.")
    if options.jpeg_quality is not None and (options.jpeg_quality < 1 or options.jpeg_quality > 100):
        raise UsageError("--jpeg-quality must be between 1 and 100.")
    if options.output_format == "pdf" and options.dpi is not None:
        raise UsageError("--dpi applies only to PNG and JPEG output.")
    if options.output_format != "jpeg" and options.jpeg_quality is not None:
        raise UsageError("--jpeg-quality applies only to JPEG output.")
    if options.output_format != "jpeg" and options.background is not None:
        raise UsageError("--background applies only to JPEG output.")
    if application == "pdf":
        if options.output_format == "pdf":
            raise UsageError("PDF-to-PDF rewriting is not supported.")
        _reject_app_options(options, allowed={"pages", "exclude_annotations"}, application="PDF input")
    elif application == "word":
        _reject_app_options(
            options,
            allowed={"pages", "include_markup", "bookmarks", "pdf_a", "update_toc"},
            application="Word",
        )
        if options.output_format == "pdf" and options.pages:
            pages = parse_number_selection(options.pages, label="pages", count=None) or []
            if not is_contiguous(pages):
                raise UsageError(
                    "Word PDF output supports only a contiguous --pages selection. Use image output for noncontiguous pages."
                )
    elif application == "powerpoint":
        _reject_app_options(
            options,
            allowed={"slides", "include_hidden", "output_type", "frame_slides", "pdf_a"},
            application="PowerPoint",
        )
        if options.image_engine == "office" and options.output_type != "slides":
            raise UsageError("--image-engine office supports PowerPoint slide output only.")
        if options.image_engine == "office" and options.jpeg_quality is not None:
            raise UsageError("--jpeg-quality does not apply to the PowerPoint Office image engine.")
    elif application == "excel":
        _reject_app_options(
            options,
            allowed={
                "sheets",
                "range_value",
                "charts",
                "charts_all",
                "ignore_print_area",
                "show_formulas",
                "show_headings",
                "recalculate",
                "update_links",
                "refresh_data",
            },
            application="Excel",
        )
        if options.range_value and options.sheets:
            raise UsageError("--range cannot be combined with --sheet.")
        if options.charts and options.charts_all:
            raise UsageError("Use repeated --chart selectors or --charts all, not both.")
        if (options.charts or options.charts_all) and options.output_format == "pdf":
            raise UsageError("Excel chart-only selection supports PNG and JPEG output only.")
        if (options.charts or options.charts_all) and options.dpi is not None:
            raise UsageError("--dpi does not apply to Excel's native chart image export.")
        if (options.charts or options.charts_all) and (options.sheets or options.range_value):
            raise UsageError("Chart-only selection cannot be combined with --sheet or --range.")
    if options.image_engine == "office" and not (
        application == "powerpoint" and options.output_format in {"png", "jpeg"}
    ):
        raise UsageError("--image-engine office is valid only for PowerPoint PNG or JPEG slide output.")
    if options.pdf_a and options.output_format != "pdf":
        raise UsageError("--pdf-a applies only to PDF output.")


def _reject_app_options(options: ExportOptions, *, allowed: set[str], application: str) -> None:
    defaults: dict[str, Any] = {
        "pages": None,
        "include_markup": False,
        "bookmarks": "headings",
        "pdf_a": False,
        "update_toc": True,
        "slides": None,
        "include_hidden": False,
        "output_type": "slides",
        "frame_slides": False,
        "sheets": [],
        "range_value": None,
        "charts": [],
        "charts_all": False,
        "ignore_print_area": False,
        "show_formulas": False,
        "show_headings": False,
        "recalculate": "never",
        "update_links": False,
        "refresh_data": False,
        "exclude_annotations": False,
    }
    names = {
        "range_value": "--range",
        "update_toc": "--no-update-toc",
        "sheets": "--sheet",
        "charts": "--chart",
        "charts_all": "--charts",
    }
    invalid = []
    for name, default in defaults.items():
        if name not in allowed and getattr(options, name) != default:
            invalid.append(names.get(name, f"--{name.replace('_', '-')}"))
    if invalid:
        raise UsageError(f"{', '.join(sorted(invalid))} is not valid for {application}.")


def _preflight_destination(destination: Path, options: ExportOptions) -> None:
    if options.output_format == "pdf":
        if destination.suffix.lower() != ".pdf":
            raise UsageError("PDF output path must end in .pdf.")
        if destination.exists() and not destination.is_file():
            raise UsageError(f"PDF output path is not a file: {destination}.", code="output_not_file")
        if destination.exists() and not options.force:
            raise UsageError(f"Output already exists: {destination}. Use --force to replace it.", code="output_exists")
    elif (
        destination.exists()
        and destination.is_file()
        and destination.suffix.lower()
        not in {
            ".png",
            ".jpg",
            ".jpeg",
        }
    ):
        raise UsageError(f"Image output path is not a directory: {destination}", code="output_not_directory")
    elif (
        not destination.exists()
        and destination.suffix.lower() in {".png", ".jpg", ".jpeg"}
        and destination.suffix.lower() not in ({".png"} if options.output_format == "png" else {".jpg", ".jpeg"})
    ):
        expected = ".png" if options.output_format == "png" else ".jpg or .jpeg"
        raise UsageError(f"{options.output_format.upper()} output path must end in {expected}.")


def _preflight_manifest(source: Path, destination: Path, options: ExportOptions) -> None:
    if options.manifest is None:
        return
    manifest = options.manifest.expanduser().resolve()
    ensure_distinct_paths(source, manifest)
    if destination.suffix.lower() in {".pdf", ".png", ".jpg", ".jpeg"}:
        ensure_distinct_paths(destination, manifest)
    if manifest.exists() and manifest.is_dir():
        raise UsageError(f"Manifest path is a directory: {manifest}.", code="manifest_is_directory")
    if manifest.exists() and not options.force:
        raise UsageError(f"Manifest already exists: {manifest}. Use --force to replace it.", code="manifest_exists")


def _write_failure_manifest(options: ExportOptions, exc: OfficeExportError, started: float) -> None:
    if options.manifest is None:
        return
    manifest = options.manifest.expanduser().resolve()
    source = options.source.expanduser().resolve()
    destination = resolve_output_path(source, options.output_format, options.output)
    try:
        ensure_distinct_paths(source, manifest)
        if destination.suffix.lower() in {".pdf", ".png", ".jpg", ".jpeg"}:
            ensure_distinct_paths(destination, manifest)
    except UsageError:
        return
    if manifest.is_dir() or (manifest.exists() and not options.force):
        return
    try:
        if source.is_file():
            try:
                _, source_format = detect_source(source)
            except UsageError:
                source_format = source.suffix.lower().lstrip(".") or "unknown"
            payload = base_result(mode="export")
            payload["source"] = source_record(source, source_format)
        else:
            payload = base_result(mode="export")
        payload.update(
            {
                "ok": False,
                "options": options.manifest_options(),
                "error": exc.context.to_dict(),
                "duration_seconds": round(time.perf_counter() - started, 3),
                "manifest": str(manifest),
            }
        )
        write_manifest(manifest, payload, force=options.force)
    except Exception:
        return


def _option_warnings(application: str, options: ExportOptions) -> list[dict[str, str]]:
    warnings: list[dict[str, str]] = []
    if application == "excel" and options.update_links:
        warnings.append(
            {
                "code": "external_links_enabled",
                "message": "External workbook link updates were explicitly enabled.",
            }
        )
    if application == "excel" and options.refresh_data:
        warnings.append(
            {
                "code": "external_data_refresh_enabled",
                "message": "External data refresh was explicitly enabled.",
            }
        )
    if application == "excel" and options.recalculate != "never":
        warnings.append(
            {
                "code": "recalculation_enabled",
                "message": f"Excel recalculation mode '{options.recalculate}' was explicitly enabled.",
            }
        )
    if options.max_megapixels > DEFAULT_MAX_MEGAPIXELS:
        warnings.append(
            {
                "code": "pixel_limit_raised",
                "message": (
                    f"The image safety limit was raised from {DEFAULT_MAX_MEGAPIXELS:g} "
                    f"to {options.max_megapixels:g} megapixels."
                ),
            }
        )
    return warnings


def _rendering_details(application: str, options: ExportOptions) -> tuple[str, bool]:
    if application == "pdf":
        return "pdfium", False
    if application == "excel" and (options.charts or options.charts_all):
        return "office", False
    if application == "powerpoint" and options.output_format in {"png", "jpeg"} and options.image_engine == "office":
        return "office", False
    if options.output_format == "pdf":
        return "office", False
    return "office+pdfium", True


def json_safe_result(payload: dict[str, Any]) -> str:
    return json.dumps(payload, indent=2, ensure_ascii=False) + "\n"


def default_output_for(source: Path, output_format: str) -> Path:
    return default_output_path(source.resolve(), output_format)


def platform_supports_office() -> bool:
    return sys.platform == "win32"
