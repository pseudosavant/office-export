from __future__ import annotations

import math
import warnings
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from PIL import Image, ImageColor

from office_export.errors import ConversionError, SecurityError, UsageError
from office_export.results import dependency_version

DEFAULT_DPI = 150
DEFAULT_JPEG_QUALITY = 92
DEFAULT_MAX_MEGAPIXELS = 50.0


@dataclass(slots=True)
class RasterPage:
    source_page: int
    filename: str
    logical: dict[str, Any] = field(default_factory=dict)


def inspect_pdf(path: Path) -> dict[str, Any]:
    pdfium, raw = _load_pdfium()
    document = _open_document(pdfium, raw, path)
    try:
        form_type = int(document.get_formtype())
        form_warning = _initialize_forms(document, raw, form_type)
        pages: list[dict[str, Any]] = []
        for index in range(len(document)):
            page = document.get_page(index)
            try:
                crop = page.get_cropbox()
                media = page.get_mediabox()
                label = document.get_page_label(index)
                pages.append(
                    {
                        "index": index + 1,
                        "label": label or None,
                        "width_points": round(float(page.get_width()), 4),
                        "height_points": round(float(page.get_height()), 4),
                        "rotation": int(page.get_rotation()),
                        "crop_box": [round(float(value), 4) for value in crop],
                        "media_box": [round(float(value), 4) for value in media],
                    }
                )
            finally:
                page.close()
        warning_items = [form_warning] if form_warning else []
        return {
            "application": None,
            "page_count": len(document),
            "pages": pages,
            "metadata": {key: value for key, value in document.get_metadata_dict().items() if value},
            "form_type": _form_type_name(raw, form_type),
            "warnings": warning_items,
            "dependencies": {"pdfium": dependency_version("pypdfium2")},
        }
    finally:
        document.close_forms()
        document.close()


def rasterize_pdf(
    path: Path,
    output_dir: Path,
    pages: list[RasterPage],
    *,
    output_format: str,
    dpi: int = DEFAULT_DPI,
    jpeg_quality: int = DEFAULT_JPEG_QUALITY,
    background: str = "white",
    exclude_annotations: bool = False,
    max_megapixels: float = DEFAULT_MAX_MEGAPIXELS,
) -> tuple[list[Path], list[dict[str, Any]], list[dict[str, str]]]:
    if dpi < 36 or dpi > 2400:
        raise UsageError("--dpi must be between 36 and 2400.")
    if jpeg_quality < 1 or jpeg_quality > 100:
        raise UsageError("--jpeg-quality must be between 1 and 100.")
    if max_megapixels <= 0:
        raise UsageError("--max-megapixels must be greater than zero.")
    if output_format not in {"png", "jpeg"}:
        raise UsageError("PDF rasterization supports only PNG and JPEG output.")
    try:
        background_rgba = ImageColor.getcolor(background, "RGBA")
    except ValueError as exc:
        raise UsageError(f"Invalid background color '{background}'.") from exc

    pdfium, raw = _load_pdfium()
    document = _open_document(pdfium, raw, path)
    output_dir.mkdir(parents=True, exist_ok=True)
    outputs: list[Path] = []
    mappings: list[dict[str, Any]] = []
    warning_items: list[dict[str, str]] = []
    try:
        form_type = int(document.get_formtype())
        form_warning = _initialize_forms(document, raw, form_type)
        if form_warning:
            warning_items.append(form_warning)
        page_count = len(document)
        for selected in pages:
            if selected.source_page < 1 or selected.source_page > page_count:
                raise UsageError(f"Selected page {selected.source_page} is outside the PDF page count of {page_count}.")
            page = document.get_page(selected.source_page - 1)
            bitmap = None
            image = None
            encoded = None
            try:
                scale = dpi / 72.0
                width = math.ceil(float(page.get_width()) * scale)
                height = math.ceil(float(page.get_height()) * scale)
                pixel_count = width * height
                if pixel_count > max_megapixels * 1_000_000:
                    raise ConversionError(
                        "pixel_limit_exceeded",
                        f"Page {selected.source_page} would create {pixel_count / 1_000_000:.1f} megapixels, "
                        f"above the {max_megapixels:g} megapixel limit.",
                        details={"page": selected.source_page, "width": width, "height": height},
                    )

                has_transparency = bool(raw.FPDFPage_HasTransparency(page))
                preserve_alpha = output_format == "png" and has_transparency
                fill = (255, 255, 255, 0) if preserve_alpha else background_rgba
                if not preserve_alpha:
                    fill = (fill[0], fill[1], fill[2], 255)
                bitmap = page.render(
                    scale=scale,
                    may_draw_forms=not exclude_annotations,
                    draw_annots=not exclude_annotations,
                    fill_color=fill,
                    maybe_alpha=preserve_alpha,
                    optimize_mode="print",
                )
                image = bitmap.to_pil()
                encoded = image.copy()
                output_path = output_dir / selected.filename
                if output_format == "jpeg":
                    if encoded.mode not in {"RGB", "L"}:
                        flattened = Image.new("RGB", encoded.size, background_rgba[:3])
                        if "A" in encoded.getbands():
                            flattened.paste(encoded.convert("RGBA"), mask=encoded.getchannel("A"))
                        else:
                            flattened.paste(encoded.convert("RGB"))
                        encoded.close()
                        encoded = flattened
                    elif encoded.mode == "L":
                        encoded = encoded.convert("RGB")
                    encoded.save(output_path, format="JPEG", quality=jpeg_quality, dpi=(dpi, dpi))
                else:
                    encoded.save(output_path, format="PNG", dpi=(dpi, dpi))
                outputs.append(output_path)
                mappings.append(
                    {
                        **selected.logical,
                        "source_page": selected.source_page,
                        "output": str(output_path),
                        "width": width,
                        "height": height,
                        "dpi": dpi,
                    }
                )
            except (UsageError, ConversionError):
                raise
            except Exception as exc:
                raise ConversionError(
                    "pdf_rasterization_failed",
                    f"PDFium could not render page {selected.source_page} ({type(exc).__name__}).",
                    details={"page": selected.source_page},
                ) from exc
            finally:
                if encoded is not None:
                    encoded.close()
                if image is not None:
                    image.close()
                if bitmap is not None:
                    bitmap.close()
                page.close()
        return outputs, mappings, warning_items
    finally:
        document.close_forms()
        document.close()


def verify_unprotected_pdf(path: Path) -> None:
    pdfium, raw = _load_pdfium()
    document = _open_document(pdfium, raw, path)
    try:
        _reject_security_handler(raw, document, path)
        if len(document) < 1:
            raise ConversionError("empty_pdf", "Office created a PDF with no pages.")
    finally:
        document.close()


def _load_pdfium() -> tuple[Any, Any]:
    try:
        import pypdfium2 as pdfium
        from pypdfium2 import raw
    except ImportError as exc:
        raise ConversionError(
            "pdfium_unavailable",
            "PDFium support is unavailable. Reinstall office-export with its required dependencies.",
        ) from exc
    return pdfium, raw


def _open_document(pdfium: Any, raw: Any, path: Path) -> Any:
    try:
        document = pdfium.PdfDocument(path)
    except Exception as exc:
        text = str(exc).lower()
        if "password" in text or "security" in text:
            raise SecurityError(
                "pdf_password_or_protection",
                "Encrypted or password-protected PDFs are not supported.",
            ) from exc
        raise ConversionError(
            "pdf_open_failed",
            f"PDFium could not open '{path}' ({type(exc).__name__}).",
        ) from exc
    try:
        _reject_security_handler(raw, document, path)
    except Exception:
        document.close()
        raise
    return document


def _reject_security_handler(raw: Any, document: Any, path: Path) -> None:
    revision = int(raw.FPDF_GetSecurityHandlerRevision(document))
    if revision >= 0:
        permissions = int(raw.FPDF_GetDocUserPermissions(document))
        raise SecurityError(
            "pdf_restricted",
            f"Encrypted or permission-restricted PDF content is not supported: {path}",
            details={"security_revision": revision, "permissions": permissions},
        )


def _initialize_forms(document: Any, raw: Any, form_type: int) -> dict[str, str] | None:
    if form_type == int(raw.FORMTYPE_NONE):
        return None
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        document.init_forms()
    if form_type in {int(raw.FORMTYPE_XFA_FULL), int(raw.FORMTYPE_XFA_FOREGROUND)}:
        return {
            "code": "xfa_rendering_limited",
            "message": "This PDF contains XFA forms. Dynamic form content may differ from a full PDF viewer.",
        }
    return None


def _form_type_name(raw: Any, form_type: int) -> str:
    names = {
        int(raw.FORMTYPE_NONE): "none",
        int(raw.FORMTYPE_ACRO_FORM): "acroform",
        int(raw.FORMTYPE_XFA_FULL): "xfa_full",
        int(raw.FORMTYPE_XFA_FOREGROUND): "xfa_foreground",
    }
    return names.get(form_type, f"unknown_{form_type}")
