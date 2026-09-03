from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path
from typing import Any, TextIO

from office_export import __version__
from office_export.core import ExportOptions, default_output_for, export_document, format_capabilities, inspect_document
from office_export.doctor import run_doctor
from office_export.errors import EXIT_INTERNAL, OfficeExportError, UsageError
from office_export.naming import ensure_distinct_paths
from office_export.results import SCHEMA_VERSION, base_result, write_manifest
from office_export.skill import install_skill, remove_skill

PROGRAM_NAME = "office-export"


class CliArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        raise UsageError(message)


def build_root_help() -> str:
    return f"""{PROGRAM_NAME} {__version__}

Export Microsoft Office documents through desktop Office, or rasterize PDFs.

Usage:
  {PROGRAM_NAME} INPUT --to pdf|png|jpeg [OPTIONS]
  {PROGRAM_NAME} inspect INPUT [--json]
  {PROGRAM_NAME} doctor [--json]
  {PROGRAM_NAME} formats [--json]
  {PROGRAM_NAME} batch PATH --to pdf|png|jpeg [OPTIONS]
  {PROGRAM_NAME} skill install [--skills-dir DIR] [--force] [--json]
  {PROGRAM_NAME} skill remove [--skills-dir DIR] [--force] [--json]

Examples:
  {PROGRAM_NAME} report.docx --to pdf
  {PROGRAM_NAME} deck.pptx --to png --slides 1,3-5 --dpi 200
  {PROGRAM_NAME} model.xlsx --to jpeg --chart "Dashboard!Revenue"
  {PROGRAM_NAME} document.pdf --to png --pages 1-4

Run `{PROGRAM_NAME} INPUT --to FORMAT --help` for export options.
"""


def build_export_parser() -> CliArgumentParser:
    parser = CliArgumentParser(prog=PROGRAM_NAME, add_help=False)
    parser.add_argument("source", type=Path)
    parser.add_argument("--to", dest="output_format", choices=("pdf", "png", "jpeg"), required=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--dpi", type=int)
    parser.add_argument("--jpeg-quality", type=int)
    parser.add_argument("--background")
    parser.add_argument("--quality", choices=("screen", "print"), default="print")
    parser.add_argument("--timeout", type=float, default=120.0)
    parser.add_argument("--keep-intermediate", action="store_true")
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--image-engine", choices=("pdfium", "office"), default="pdfium")
    parser.add_argument("--max-megapixels", type=float, default=50.0)
    parser.add_argument("--pages")
    parser.add_argument("--include-markup", action="store_true")
    parser.add_argument("--bookmarks", choices=("none", "headings", "word"), default="headings")
    parser.add_argument("--pdf-a", action="store_true")
    parser.add_argument("--no-update-toc", dest="update_toc", action="store_false", default=True)
    parser.add_argument("--slides")
    parser.add_argument("--include-hidden", action="store_true")
    parser.add_argument(
        "--output-type",
        choices=("slides", "notes", "outline", "handout1", "handout2", "handout3", "handout4", "handout6", "handout9"),
        default="slides",
    )
    parser.add_argument("--frame-slides", action="store_true")
    parser.add_argument("--sheet", dest="sheets", action="append", default=[])
    parser.add_argument("--range", dest="range_value")
    parser.add_argument("--chart", dest="charts", action="append", default=[])
    parser.add_argument("--charts", dest="charts_mode", choices=("all",))
    parser.add_argument("--ignore-print-area", action="store_true")
    parser.add_argument("--show-formulas", action="store_true")
    parser.add_argument("--show-headings", action="store_true")
    parser.add_argument("--recalculate", choices=("never", "auto", "full"), default="never")
    parser.add_argument("--update-links", action="store_true")
    parser.add_argument("--refresh-data", action="store_true")
    parser.add_argument("--exclude-annotations", action="store_true")
    return parser


def build_export_help() -> str:
    return f"""Usage: {PROGRAM_NAME} INPUT --to pdf|png|jpeg [OPTIONS]

Common options:
  --output PATH                 Override the deterministic default output path.
  --force                       Replace only planned output collisions.
  --dpi INTEGER                 Image resolution from 36 through 2400. Default: 150.
  --jpeg-quality INTEGER        JPEG quality from 1 through 100. Default: 92.
  --background COLOR            JPEG transparency background. Default: white.
  --quality screen|print        Native Office PDF quality. Default: print.
  --timeout SECONDS             Office worker timeout. Default: 120.
  --keep-intermediate           Keep the Office PDF used for image output.
  --image-engine pdfium|office  Office is valid only for PowerPoint slides.
  --max-megapixels NUMBER       Per-page image safety limit. Default: 50.
  --manifest PATH               Persist the JSON conversion result.
  --json                        Write the JSON conversion result to stdout.
  --verbose                     Include internal exception details for unexpected failures.

Word options:
  --pages LIST                  One-based pages such as 1,3-5.
  --include-markup
  --bookmarks none|headings|word
  --pdf-a
  --no-update-toc

PowerPoint options:
  --slides LIST                 One-based slides such as 1,3-5.
  --include-hidden
  --output-type TYPE            Slides, notes, outline, or a handout layout.
  --frame-slides
  --pdf-a

Excel options:
  --sheet NAME_OR_INDEX         Repeat to select sheets.
  --range SHEET!A1:H40
  --chart SHEET!NAME            Repeat to select charts.
  --charts all
  --ignore-print-area
  --show-formulas
  --show-headings
  --recalculate never|auto|full
  --update-links
  --refresh-data

PDF input options:
  --pages LIST
  --exclude-annotations
"""


def main(
    argv: list[str] | None = None,
    *,
    stdout: TextIO | None = None,
    stderr: TextIO | None = None,
) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    stdout = stdout or sys.stdout
    stderr = stderr or sys.stderr
    json_mode = "--json" in args
    verbose = "--verbose" in args
    try:
        if not args:
            stdout.write(build_root_help())
            return 0
        if args == ["--version"]:
            stdout.write(f"{PROGRAM_NAME} {__version__}\n")
            return 0
        if args == ["--about"]:
            stdout.write(_about_text())
            return 0
        if args[0] in {"-h", "--help"}:
            stdout.write(build_root_help())
            return 0
        if args[0] == "skill":
            return _run_skill(args[1:], stdout=stdout)
        if args[0] == "inspect":
            return _run_inspect(args[1:], stdout=stdout)
        if args[0] == "doctor":
            return _run_doctor(args[1:], stdout=stdout)
        if args[0] == "formats":
            return _run_formats(args[1:], stdout=stdout)
        if args[0] == "batch":
            return _run_batch(args[1:], stdout=stdout)
        if "--help" in args or "-h" in args:
            stdout.write(build_export_help())
            return 0
        parsed = build_export_parser().parse_args(args)
        options = _options_from_args(parsed)
        result = export_document(options)
        _write_export_result(result, json_mode=parsed.json, stdout=stdout)
        return 0
    except OfficeExportError as exc:
        _write_error(exc, json_mode=json_mode, stdout=stdout, stderr=stderr)
        return exc.context.exit_code
    except Exception as exc:
        message = f"Unexpected internal error ({type(exc).__name__})."
        details = {"exception": repr(exc)} if verbose else None
        payload: dict[str, Any] = {
            "ok": False,
            "schema_version": SCHEMA_VERSION,
            "error": {"code": "internal_error", "message": message},
        }
        if details:
            payload["error"]["details"] = details
        if json_mode:
            stdout.write(json.dumps(payload, indent=2) + "\n")
        else:
            stderr.write(f"error: {message}\n")
            if verbose:
                stderr.write(f"details: {exc!r}\n")
        return EXIT_INTERNAL


def _run_skill(args: list[str], *, stdout: TextIO) -> int:
    if not args or args[0] in {"-h", "--help"}:
        stdout.write(
            f"Usage:\n  {PROGRAM_NAME} skill install [--skills-dir DIR] [--force] [--json]\n"
            f"  {PROGRAM_NAME} skill remove [--skills-dir DIR] [--force] [--json]\n"
        )
        return 0
    parser = CliArgumentParser(prog=f"{PROGRAM_NAME} skill", add_help=False)
    parser.add_argument("action", choices=("install", "remove"))
    parser.add_argument("--skills-dir", type=Path)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--json", action="store_true")
    parsed = parser.parse_args(args)
    root = parsed.skills_dir.expanduser().resolve() if parsed.skills_dir else None
    result = (
        install_skill(root, force=parsed.force)
        if parsed.action == "install"
        else remove_skill(root, force=parsed.force)
    )
    payload = {"ok": True, "schema_version": SCHEMA_VERSION, "mode": f"skill_{parsed.action}", **result}
    if parsed.json:
        stdout.write(json.dumps(payload, indent=2, ensure_ascii=False) + "\n")
    else:
        verb = "Installed" if parsed.action == "install" else "Removed"
        if parsed.action == "remove" and not result["removed"]:
            stdout.write(f"Skill is not installed at {result['path']}\n")
        else:
            stdout.write(f"{verb} {result['path']}\n")
    return 0


def _run_inspect(args: list[str], *, stdout: TextIO) -> int:
    if not args or args[0] in {"-h", "--help"}:
        stdout.write(f"Usage: {PROGRAM_NAME} inspect INPUT [--timeout SECONDS] [--json]\n")
        return 0
    parser = CliArgumentParser(prog=f"{PROGRAM_NAME} inspect", add_help=False)
    parser.add_argument("source", type=Path)
    parser.add_argument("--timeout", type=float, default=120.0)
    parser.add_argument("--json", action="store_true")
    parsed = parser.parse_args(args)
    if parsed.timeout <= 0:
        raise UsageError("--timeout must be greater than zero.")
    result = inspect_document(parsed.source, timeout=parsed.timeout)
    if parsed.json:
        stdout.write(json.dumps(result, indent=2, ensure_ascii=False) + "\n")
    else:
        stdout.write(_inspection_text(result))
    return 0


def _run_doctor(args: list[str], *, stdout: TextIO) -> int:
    if args and args[0] in {"-h", "--help"}:
        stdout.write(
            f"Usage: {PROGRAM_NAME} doctor [--timeout SECONDS] [--smoke-word FILE] "
            "[--smoke-excel FILE] [--smoke-powerpoint FILE] [--json]\n"
        )
        return 0
    parser = CliArgumentParser(prog=f"{PROGRAM_NAME} doctor", add_help=False)
    parser.add_argument("--timeout", type=float, default=30.0)
    parser.add_argument("--smoke-word", type=Path)
    parser.add_argument("--smoke-excel", type=Path)
    parser.add_argument("--smoke-powerpoint", type=Path)
    parser.add_argument("--json", action="store_true")
    parsed = parser.parse_args(args)
    if parsed.timeout <= 0:
        raise UsageError("--timeout must be greater than zero.")
    payload = base_result(mode="doctor")
    payload["diagnostics"] = run_doctor(timeout=parsed.timeout)
    smoke_sources = {
        "word": parsed.smoke_word,
        "excel": parsed.smoke_excel,
        "powerpoint": parsed.smoke_powerpoint,
    }
    if any(smoke_sources.values()):
        payload["diagnostics"]["smoke_tests"] = _doctor_smoke(smoke_sources, timeout=parsed.timeout)
    if parsed.json:
        stdout.write(json.dumps(payload, indent=2, ensure_ascii=False) + "\n")
    else:
        stdout.write(_doctor_text(payload["diagnostics"]))
    return 0


def _run_formats(args: list[str], *, stdout: TextIO) -> int:
    if args and args[0] in {"-h", "--help"}:
        stdout.write(f"Usage: {PROGRAM_NAME} formats [--json]\n")
        return 0
    parser = CliArgumentParser(prog=f"{PROGRAM_NAME} formats", add_help=False)
    parser.add_argument("--json", action="store_true")
    parsed = parser.parse_args(args)
    payload = base_result(mode="formats")
    payload["formats"] = format_capabilities()
    if parsed.json:
        stdout.write(json.dumps(payload, indent=2) + "\n")
    else:
        stdout.write("Input  Application  Outputs\n")
        for name, detail in payload["formats"]["inputs"].items():
            stdout.write(f"{name:<6} {(detail['application'] or 'none'):<12} {', '.join(detail['outputs'])}\n")
    return 0


def _run_batch(args: list[str], *, stdout: TextIO) -> int:
    if not args or args[0] in {"-h", "--help"}:
        stdout.write(
            f"Usage: {PROGRAM_NAME} batch PATH --to pdf|png|jpeg "
            "[--output-dir DIR] [--recursive] [--continue-on-error] [--json]\n"
        )
        return 0
    parser = CliArgumentParser(prog=f"{PROGRAM_NAME} batch", add_help=False)
    parser.add_argument("path", type=Path)
    parser.add_argument("--to", dest="output_format", choices=("pdf", "png", "jpeg"), required=True)
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--recursive", action="store_true")
    parser.add_argument("--continue-on-error", action="store_true")
    parser.add_argument("--jobs", type=int, default=1)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--dpi", type=int)
    parser.add_argument("--jpeg-quality", type=int)
    parser.add_argument("--background")
    parser.add_argument("--quality", choices=("screen", "print"), default="print")
    parser.add_argument("--timeout", type=float, default=120.0)
    parser.add_argument("--keep-intermediate", action="store_true")
    parser.add_argument("--image-engine", choices=("pdfium", "office"), default="pdfium")
    parser.add_argument("--max-megapixels", type=float, default=50.0)
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--json", action="store_true")
    parsed = parser.parse_args(args)
    if parsed.jobs != 1:
        raise UsageError("Only --jobs 1 is currently supported.")
    sources = _batch_sources(parsed.path, recursive=parsed.recursive)
    output_root = parsed.output_dir.expanduser().resolve() if parsed.output_dir else None
    if output_root:
        output_root.mkdir(parents=True, exist_ok=True)
    manifest_target = parsed.manifest.expanduser().resolve() if parsed.manifest else None
    if manifest_target is not None:
        if manifest_target.exists() and manifest_target.is_dir():
            raise UsageError(f"Manifest path is a directory: {manifest_target}.", code="manifest_is_directory")
        if manifest_target.exists() and not parsed.force:
            raise UsageError(
                f"Manifest already exists: {manifest_target}. Use --force to replace it.",
                code="manifest_exists",
            )
        for source in sources:
            ensure_distinct_paths(source, manifest_target)
    results: list[dict[str, Any]] = []
    failures: list[dict[str, Any]] = []
    for source in sources:
        destination = None
        if output_root:
            destination = output_root / (
                f"{source.stem}.pdf"
                if parsed.output_format == "pdf"
                else f"{source.stem} - {'PNG' if parsed.output_format == 'png' else 'JPEG'} export"
            )
        if manifest_target is not None:
            ensure_distinct_paths(destination or default_output_for(source, parsed.output_format), manifest_target)
        options = ExportOptions(
            source=source,
            output_format=parsed.output_format,
            output=destination,
            force=parsed.force,
            dpi=parsed.dpi,
            jpeg_quality=parsed.jpeg_quality,
            background=parsed.background,
            quality=parsed.quality,
            timeout=parsed.timeout,
            keep_intermediate=parsed.keep_intermediate,
            image_engine=parsed.image_engine,
            max_megapixels=parsed.max_megapixels,
        )
        try:
            results.append(export_document(options))
        except OfficeExportError as exc:
            failure = {"source": str(source), "error": exc.context.to_dict()}
            failures.append(failure)
            if not parsed.continue_on_error:
                raise
    payload = base_result(mode="batch")
    payload.update(
        {
            "ok": not failures,
            "path": str(parsed.path.expanduser().resolve()),
            "jobs": 1,
            "converted": len(results),
            "failed": len(failures),
            "results": results,
            "failures": failures,
        }
    )
    if manifest_target is not None:
        payload["manifest"] = str(manifest_target)
        write_manifest(manifest_target, payload, force=parsed.force)
    if parsed.json:
        stdout.write(json.dumps(payload, indent=2, ensure_ascii=False) + "\n")
    else:
        stdout.write(f"Converted {len(results)} file(s). Failed: {len(failures)}.\n")
        for result in results:
            for output in result["outputs"]:
                stdout.write(f"{output['path']}\n")
    return 0 if not failures else 6


def _options_from_args(args: argparse.Namespace) -> ExportOptions:
    return ExportOptions(
        source=args.source,
        output_format=args.output_format,
        output=args.output,
        force=args.force,
        dpi=args.dpi,
        jpeg_quality=args.jpeg_quality,
        background=args.background,
        quality=args.quality,
        timeout=args.timeout,
        keep_intermediate=args.keep_intermediate,
        manifest=args.manifest,
        verbose=args.verbose,
        image_engine=args.image_engine,
        pages=args.pages,
        include_markup=args.include_markup,
        bookmarks=args.bookmarks,
        pdf_a=args.pdf_a,
        update_toc=args.update_toc,
        slides=args.slides,
        include_hidden=args.include_hidden,
        output_type=args.output_type,
        frame_slides=args.frame_slides,
        sheets=args.sheets,
        range_value=args.range_value,
        charts=args.charts,
        charts_all=args.charts_mode == "all",
        ignore_print_area=args.ignore_print_area,
        show_formulas=args.show_formulas,
        show_headings=args.show_headings,
        recalculate=args.recalculate,
        update_links=args.update_links,
        refresh_data=args.refresh_data,
        exclude_annotations=args.exclude_annotations,
        max_megapixels=args.max_megapixels,
    )


def _write_export_result(result: dict[str, Any], *, json_mode: bool, stdout: TextIO) -> None:
    if json_mode:
        stdout.write(json.dumps(result, indent=2, ensure_ascii=False) + "\n")
        return
    for output in result["outputs"]:
        stdout.write(f"{output['path']}\n")
    for warning in result.get("warnings", []):
        stdout.write(f"warning: {warning['message']}\n")


def _write_error(exc: OfficeExportError, *, json_mode: bool, stdout: TextIO, stderr: TextIO) -> None:
    if json_mode:
        stdout.write(
            json.dumps(
                {
                    "ok": False,
                    "schema_version": SCHEMA_VERSION,
                    "tool": {"name": PROGRAM_NAME, "version": __version__},
                    "error": exc.context.to_dict(),
                },
                indent=2,
                ensure_ascii=False,
            )
            + "\n"
        )
    else:
        stderr.write(f"error [{exc.context.code}]: {exc.context.message}\n")


def _inspection_text(payload: dict[str, Any]) -> str:
    source = payload["source"]
    detail = payload["inspection"]
    lines = [f"Source: {source['path']}", f"Format: {source['format']}"]
    if "page_count" in detail:
        lines.append(f"Pages: {detail['page_count']}")
    if "slide_count" in detail:
        lines.append(f"Slides: {detail['slide_count']}")
    if "sheets" in detail:
        lines.append(f"Sheets: {len(detail['sheets'])}")
        for sheet in detail["sheets"]:
            lines.append(f"  {sheet['position']}: {sheet['name']} ({sheet['kind']}, {sheet['visibility']})")
    if detail.get("charts"):
        lines.append(f"Charts: {len(detail['charts'])}")
    for warning in detail.get("warnings", []):
        lines.append(f"Warning: {warning['message']}")
    return "\n".join(lines) + "\n"


def _doctor_text(diagnostics: dict[str, Any]) -> str:
    lines = [
        f"office-export {diagnostics['tool_version']}",
        f"Platform: {diagnostics['platform']['system']} {diagnostics['platform']['release']}",
        f"Python: {diagnostics['platform']['python']}",
        f"PDF rasterization: {'available' if diagnostics['pdf_rasterization']['available'] else 'unavailable'}",
    ]
    for name, item in diagnostics["office_export"]["applications"].items():
        if item.get("available"):
            lines.append(f"{name.capitalize()}: {item['application'].get('version') or 'available'}")
        else:
            lines.append(f"{name.capitalize()}: unavailable")
    printing = diagnostics["printing"]
    lines.append(f"Print Spooler: {printing.get('spooler') or 'not applicable'}")
    lines.append(f"Printers: {len(printing.get('printers', []))}")
    for warning in diagnostics["warnings"]:
        lines.append(f"Warning: {warning['message']}")
    for name, result in diagnostics.get("smoke_tests", {}).items():
        lines.append(f"{name.capitalize()} smoke: {'passed' if result['ok'] else 'failed'}")
    return "\n".join(lines) + "\n"


def _doctor_smoke(sources: dict[str, Path | None], *, timeout: float) -> dict[str, Any]:
    results: dict[str, Any] = {}
    with tempfile.TemporaryDirectory(prefix="office-export-smoke-") as directory:
        root = Path(directory)
        for application, source in sources.items():
            if source is None:
                continue
            try:
                conversion = export_document(
                    ExportOptions(
                        source=source,
                        output_format="pdf",
                        output=root / f"{application}.pdf",
                        timeout=timeout,
                    )
                )
                results[application] = {
                    "ok": True,
                    "source": str(source.expanduser().resolve()),
                    "size": conversion["outputs"][0]["size"],
                    "warnings": conversion["warnings"],
                }
            except OfficeExportError as exc:
                results[application] = {
                    "ok": False,
                    "source": str(source.expanduser().resolve()),
                    "error": exc.context.to_dict(),
                }
    return results


def _batch_sources(path: Path, *, recursive: bool) -> list[Path]:
    target = path.expanduser().resolve()
    if target.is_file():
        return [target]
    if not target.exists():
        raise UsageError(f"Batch path does not exist: {target}", code="batch_path_not_found")
    if not target.is_dir():
        raise UsageError(f"Batch path is not a directory or file: {target}")
    iterator = target.rglob("*") if recursive else target.iterdir()
    supported = {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".pdf"}
    sources = sorted(
        (item for item in iterator if item.is_file() and item.suffix.lower() in supported),
        key=lambda p: str(p).casefold(),
    )
    if not sources:
        raise UsageError(f"No supported input files were found in: {target}", code="batch_empty")
    return sources


def _about_text() -> str:
    return f"""{PROGRAM_NAME} {__version__}
Export Office documents through installed desktop Microsoft Office.
Repository: https://github.com/pseudosavant/office-export
License: MIT
"""


if __name__ == "__main__":
    raise SystemExit(main())
