from __future__ import annotations

import os
import re
import unicodedata
from pathlib import Path

from office_export.errors import InputError, UsageError

PDF_SUFFIX = ".pdf"
IMAGE_SUFFIXES = {"png": ".png", "jpeg": ".jpg"}
RESERVED_NAMES = {
    "con",
    "prn",
    "aux",
    "nul",
    *(f"com{number}" for number in range(1, 10)),
    *(f"lpt{number}" for number in range(1, 10)),
}


def safe_component(value: str, *, fallback: str = "item") -> str:
    value = unicodedata.normalize("NFKC", value).strip().lower()
    value = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "-", value)
    value = re.sub(r"[^\w\s-]", "-", value, flags=re.UNICODE)
    value = re.sub(r"[\s_-]+", "-", value).strip(" .-")
    if not value:
        value = fallback
    if value.casefold() in RESERVED_NAMES:
        value = f"{value}-item"
    return value[:80].rstrip(" .-") or fallback


def default_output_path(source: Path, output_format: str) -> Path:
    if output_format == "pdf":
        return source.with_suffix(PDF_SUFFIX)
    label = "PNG" if output_format == "png" else "JPEG"
    return source.with_name(f"{source.stem} - {label} export")


def resolve_output_path(source: Path, output_format: str, explicit: Path | None) -> Path:
    raw = explicit if explicit is not None else default_output_path(source, output_format)
    return raw.expanduser().resolve()


def validate_source(source: Path) -> Path:
    path = source.expanduser().resolve()
    if not path.exists():
        raise InputError("source_not_found", f"Source file does not exist: {path}")
    if not path.is_file():
        raise InputError("source_not_file", f"Source path is not a file: {path}")
    return path


def ensure_distinct_paths(source: Path, output: Path) -> None:
    try:
        same = os.path.samefile(source, output)
    except (FileNotFoundError, OSError):
        same = os.path.normcase(str(source)) == os.path.normcase(str(output))
    if same:
        raise UsageError("The output path must be different from the source path.")


def word_image_name(source: Path, page: int, total_pages: int, image_format: str) -> str:
    digits = max(3, len(str(total_pages)))
    return f"{safe_component(source.stem)}-page-{page:0{digits}d}{IMAGE_SUFFIXES[image_format]}"


def powerpoint_image_name(source: Path, slide: int, slide_count: int, image_format: str) -> str:
    digits = max(3, len(str(slide_count)))
    return f"{safe_component(source.stem)}-slide-{slide:0{digits}d}{IMAGE_SUFFIXES[image_format]}"


def excel_page_image_name(
    source: Path,
    sheet_name: str,
    page: int,
    total_pages: int,
    image_format: str,
) -> str:
    digits = max(3, len(str(total_pages)))
    return (
        f"{safe_component(source.stem)}-sheet-{safe_component(sheet_name, fallback='sheet')}"
        f"-page-{page:0{digits}d}{IMAGE_SUFFIXES[image_format]}"
    )


def excel_chart_image_name(source: Path, sheet_name: str, chart_name: str, image_format: str) -> str:
    return (
        f"{safe_component(source.stem)}-sheet-{safe_component(sheet_name, fallback='sheet')}"
        f"-chart-{safe_component(chart_name, fallback='chart')}{IMAGE_SUFFIXES[image_format]}"
    )
