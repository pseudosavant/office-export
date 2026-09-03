# office-export

`office-export` exports Word, Excel, and PowerPoint documents to PDF, PNG, or JPEG through the installed desktop Microsoft Office applications. It also renders PDF pages to PNG or JPEG without requiring Office.

It is a predictable CLI for faithful local rendering by people, scripts, and coding agents. Office source files are opened read-only, macros remain disabled, and source changes are never saved.

## Prerequisite

`office-export` is designed to be used with [`uv`](https://docs.astral.sh/uv/getting-started/installation/). Install `uv` before continuing. The documented workflows and managed agent skill use `uvx` to run the tool without requiring a global installation.

## Quick start with an agent

Install the managed agent skill:

```powershell
uvx office-export skill install
```

Then use `$office-export` in Codex, Claude Code, or another agent harness that supports skills:

> Use $office-export to export this presentation as PNG images, inspect the rendered slides, and report any visual problems.

The skill teaches the agent how to inspect Office and PDF files, choose safe export options, render the requested outputs, and review the results.

## What it exports

| Input | Outputs | Select content by | Required renderer |
| --- | --- | --- | --- |
| Word | PDF, PNG, JPEG | Pages | Desktop Word |
| PowerPoint | PDF, PNG, JPEG | Slides | Desktop PowerPoint |
| Excel | PDF, PNG, JPEG | Sheets, ranges, charts | Desktop Excel |
| PDF | PNG, JPEG | Pages | Bundled PDFium dependency |

Office inputs use the corresponding installed desktop application to preserve native layout and rendering. PDF input does not require Word, Excel, or PowerPoint.

## Use the CLI directly

Export a document without installing the package globally:

```powershell
uvx office-export report.docx --to pdf
```

The PDF is written beside the source as `report.pdf`. Check whether Office applications and supporting dependencies are ready with:

```powershell
uvx office-export doctor
```

To install the command as a persistent tool instead:

```powershell
uv tool install office-export
```

The examples below continue to use `uvx office-export` so they work without a global installation.

## Common recipes

```powershell
uvx office-export report.docx --to pdf
uvx office-export report.docx --to png --pages 1,3-5 --dpi 200
uvx office-export deck.pptx --to jpeg --slides 2-6 --dpi 200
uvx office-export model.xlsx --to pdf --sheet Summary --sheet "Q4 Charts"
uvx office-export model.xlsx --to png --range "Summary!A1:H40"
uvx office-export document.pdf --to jpeg --pages 1,3-5 --dpi 300
```

## Output locations and overwrite safety

Use `--output PATH` to choose a different destination. Existing output is never replaced unless `--force` is supplied. The force option replaces only the files planned for the current conversion.

Default destinations are:

- PDF is written beside the input as `<name>.pdf`.
- PNG files are written to `<name> - PNG export`.
- JPEG files are written to `<name> - JPEG export`.

Image names retain their logical source identity:

```text
report-page-001.png
deck-slide-003.jpg
model-sheet-summary-page-001.png
model-sheet-summary-chart-revenue.png
```

Image output always uses a directory by default. An explicit image filename is accepted when the selection produces exactly one image.

## Requirements

Office document export supports:

- Windows 11
- Microsoft 365 Apps desktop applications
- Office 2024 and Office LTSC 2024
- CPython 3.11 or newer
- `uv` and `uvx`

Office LTSC 2021 and older desktop releases are best effort. The CLI probes capabilities instead of rejecting an application only because of its version number.

Direct PDF rasterization works on any supported Python platform that has a compatible `pypdfium2` wheel. Word, Excel, and PowerPoint are not required for PDF input.

Office automation requires an interactive Windows user profile. It is not a server-side conversion service. First-run setup, modal dialogs, add-ins, Protected View, or an uninitialized Office license can block automation.

## How rendering works

Word, Excel, and PowerPoint create native PDFs through their desktop export APIs. When PNG or JPEG is requested, the default image engine rasterizes that Office-created PDF with PDFium for consistent DPI, encoding, annotations, and selection behavior.

PowerPoint can optionally use its native slide-image exporter. Excel chart selection uses native chart export to create a tightly bounded image. PDF inputs go directly to PDFium and never launch an Office application.

## Word

```powershell
uvx office-export report.docx --to pdf --bookmarks headings
uvx office-export report.docx --to pdf --pages 2-5 --include-markup
uvx office-export report.docx --to png --pages 1,4,7
```

Word PDF bookmarks can come from headings, explicit Word bookmarks, or neither. The default is `headings`. The tool updates each table of contents in memory before export without saving the source. Use `--no-update-toc` to preserve the currently displayed values and pagination.

Contiguous page selection uses Word's native fixed-format range. Noncontiguous Word pages are supported for image output by rasterizing selected pages from a complete temporary PDF. Noncontiguous Word PDF output is rejected because combining PDFs can damage tags, links, destinations, metadata, and bookmarks.

## PowerPoint

```powershell
uvx office-export deck.pptx --to pdf --slides 1,3-5
uvx office-export deck.pptx --to png --output-type notes
uvx office-export deck.pptx --to pdf --output-type handout6
uvx office-export deck.pptx --to png --image-engine office
```

The default PDFium image engine renders the Office-created PDF. It provides consistent DPI, encoding, annotation, selection, notes, and handout behavior. The optional Office image engine calls PowerPoint's native slide export. It supports slide output only.

Inspection reports both one-based slide positions and stable PowerPoint slide IDs.

## Excel

```powershell
uvx office-export model.xlsx --to pdf --sheet Summary --sheet "Q4 Charts"
uvx office-export model.xlsx --to png --range "Summary!A1:H40"
uvx office-export model.xlsx --to png --charts all
uvx office-export model.xlsx --to jpeg --chart "Dashboard!Margin" --jpeg-quality 95
```

Repeated `--sheet` values accept exact names or one-based positions. A range uses `SHEET!ADDRESS` syntax. Excel chart selection creates one tightly bounded native chart image. JPEG charts are converted from Excel's temporary PNG so quality and background handling stay deterministic.

`--dpi` does not apply to native chart export. Chart dimensions come from the workbook. Excel can lose internal workbook links during native PDF conversion. External links emitted by Excel remain subject to Excel's own behavior.

The following options intentionally change how workbook content is evaluated:

- `--recalculate auto|full`
- `--update-links`
- `--refresh-data`

They are disabled by default.

## PDF input

```powershell
uvx office-export document.pdf --to png
uvx office-export document.pdf --to jpeg --pages 1,3-5 --dpi 300
uvx office-export document.pdf --to png --exclude-annotations
```

PDF input supports PNG and JPEG output. PDF-to-PDF rewriting is not supported. Physical page selectors are one-based. `inspect` also reports page labels when the PDF contains them.

Visible annotations and standard form appearances are rendered by default. `--exclude-annotations` also excludes form widget appearances. Dynamic XFA content produces a warning. Password-protected, encrypted, or permission-restricted PDFs are rejected.

Defaults:

- 150 DPI
- JPEG quality 92
- White JPEG background
- 50 megapixels per page

Use `--max-megapixels` only after reviewing the memory cost of the requested page size and DPI.

## Inspect and troubleshoot

Inspect a source before exporting when a script or agent needs to discover its pages, slides, sheets, ranges, charts, or other metadata. Use `doctor` to diagnose the local Office and PDFium environment.

```powershell
uvx office-export inspect report.docx --json
uvx office-export inspect deck.pptx --json
uvx office-export inspect model.xlsx --json
uvx office-export inspect document.pdf --json
uvx office-export doctor --json
uvx office-export doctor --smoke-word report.docx --smoke-excel model.xlsx --smoke-powerpoint deck.pptx --json
uvx office-export formats --json
```

`doctor` reports each Office application separately from PDFium. It also reports the Office version, bitness, active printer when available, Print Spooler state, installed printers, dependency versions, and temporary-directory access.

Smoke exports are opt-in. Supply one or more known local fixtures with `--smoke-word`, `--smoke-excel`, or `--smoke-powerpoint`.

## Automation and batch conversion

### Structured results and manifests

Add `--json` to print a stable result object to stdout. Diagnostics stay on stderr. A successful conversion includes:

- Tool and schema versions
- Source size and SHA-256 hash
- Office application and version
- Active printer when relevant
- Effective options and warnings
- PDFium and Pillow versions
- Output paths, sizes, and SHA-256 hashes
- Logical source-to-output mappings
- Duration

Use `--manifest PATH` to persist the same result. Failed conversions also write a structured failure manifest when the requested manifest path is safe and writable. No sidecar manifest is created by default.

### Batch conversion

```powershell
uvx office-export batch .\incoming --to pdf
uvx office-export batch .\incoming --to png --recursive --continue-on-error --json
```

Batch conversion currently processes files sequentially and accepts only `--jobs 1`. Each Office source gets a fresh worker process and an isolated temporary directory.

## Manage the agent skill

Inspect, update, restore, install to a custom location, or remove the bundled managed skill:

```powershell
uvx office-export skill status --json
uvx office-export skill install
uvx office-export skill install --force
uvx office-export skill install --skills-dir C:\custom\skills
uvx office-export skill remove
```

The default target is `~/.agents/skills/office-export/SKILL.md`. Normally installed CLIs automatically synchronize an already-installed managed skill when its version is older and its content is intact. The running CLI version is the authority. Equal or newer skills are left alone. Missing skills are never installed automatically. This maintenance is local. It does not query package indexes, refresh uv caches, or update the CLI.

Management metadata lives in the supported YAML `metadata` mapping: `managed-by: office-export`, a quoted `managed-version`, and `managed-content-sha256: "sha256:<digest>"`. The hash covers the complete UTF-8 skill with LF line endings and the hash value replaced by `""`. It detects edits, including metadata and formatting changes. It is not a signature. Legacy HTML markers remain recognized. Managed skills with missing or invalid versions receive a fresh replacement as a recovery step.

Modified skills and skills with valid versions but missing or invalid hashes are preserved. To replace managed content explicitly, run `uvx office-export skill install --force`. Installation never overwrites an unmanaged skill, even with `--force`, and never downgrades a newer skill. Removal refuses unmanaged content and extra directory entries unless its existing `--force` override is supplied.

`skill status` reports the path, ownership, CLI and installed versions, version comparison, integrity, and automatic synchronization eligibility. Add `--json` for structured output. Skill commands do not trigger automatic synchronization. Update notices and maintenance warnings go to stderr, so normal JSON stdout stays unchanged.

Only the standard directory participates in automatic synchronization. Custom locations require explicit updates with `skill install --skills-dir PATH`. Local checkouts, local source installations, and editable builds skip automatic synchronization. An installed wheel remains eligible. Explicit installation still works during development with `uvx --from . office-export skill install`.

Updates affect future agent skill loading. They may not change instructions already loaded into a running agent session. The skill continues to teach invocation through `uvx office-export`, direct exports for explicit requests, conditional discovery and visual review, and safe source handling.

## Safety model

- Sources open read-only.
- VBA macros are force-disabled.
- AutoOpen and Auto_Open macros are not executed.
- External links do not update by default.
- Data connections do not refresh by default.
- Workbooks do not recalculate by default.
- Source changes are never saved.
- Passwords, Protected View, IRM, sensitivity restrictions, and PDF permissions are not bypassed.
- Final outputs are published only after conversion validation.
- A timeout can terminate only an Office process proven to have been created by the isolated worker.
- A pre-existing user Office process is never quit or terminated.

Office automation is not a security sandbox. Do not use it to open untrusted documents outside the protections of your Windows and Office environment.

## Known native limitations

- Office rendering can vary with Office updates, installed fonts, document compatibility settings, and printer metrics.
- Excel internal workbook links can be lost in native PDF output.
- PowerPoint-native images can differ slightly from PDFium images.
- Dynamic XFA forms may not match a full PDF viewer.
- Modal Office dialogs and add-ins can still interfere with automation.

## Reference

Useful discovery and metadata commands:

```powershell
uvx office-export --help
uvx office-export --version
uvx office-export formats
uvx office-export inspect report.docx
uvx office-export doctor
```

Exit codes:

| Code | Meaning |
| ---: | --- |
| `0` | Success |
| `2` | Usage or selection error |
| `3` | Input error |
| `4` | Required capability is unavailable |
| `5` | Security restriction |
| `6` | Conversion error |
| `7` | Office worker timeout |
| `8` | Unexpected internal error |

## Development

```powershell
$env:UV_LINK_MODE="copy"
uv sync --all-groups
uv run ruff check .
uv run ruff format --check .
uv run pytest
uv run pytest -m office
uv build
uv run twine check dist/*
```

The default test run is cross-platform and excludes tests that require licensed desktop Office. Run the `office` marker in an interactive Windows session before release.

## License

MIT. See [LICENSE](LICENSE) and [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
