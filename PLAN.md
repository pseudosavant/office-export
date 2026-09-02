# Office Export

The first public testing release is `0.1.0`. The `1.0.0` references below describe the intended stable-release contract. That stable release will follow usage and compatibility testing.

## Project summary

`office-export` is a Windows-first Python CLI for exporting Microsoft Office documents through the installed desktop Office applications.

Version 1.0.0 will support:

- Microsoft Word
- Microsoft Excel
- Microsoft PowerPoint
- PDF output
- PNG output
- JPEG output
- PDF input for PNG and JPEG rasterization

The core promise is faithful local rendering through the same Microsoft Office applications that users rely on in the UI. The tool is intended for humans, scripts, and coding agents.

Markdown and HTML are not part of version 1.0.0. They remain potential future semantic conversion formats for agent document understanding.

## Product positioning

The tool should be described as a local Windows desktop CLI. It should not be described as a server-side Office conversion service or as truly headless Office automation.

Microsoft Office assumes an interactive Windows user profile. Modal dialogs, add-ins, protected documents, and first-run setup can interfere with automation. Microsoft does not support unattended Office automation from services or noninteractive server processes.

The differentiating features are:

- Native Microsoft Office rendering fidelity
- Simple installation and execution through `uvx`
- One CLI for Word, Excel, and PowerPoint
- PDF, PNG, and JPEG outputs
- Page, slide, sheet, range, and chart selection
- Agent-friendly inspection and JSON responses
- Deterministic output naming
- Optional conversion manifests and warnings
- Safe defaults for macros, links, data refresh, and recalculation
- Process isolation, timeouts, and careful Office cleanup
- A bundled agent skill

The project will use the MIT license.

## Project and package name

The project, distribution, and console-command name is `office-export`. The import package is `office_export`. The live PyPI project API returned not found for the normalized distribution name on September 2, 2026, so it appears unclaimed. PyPI availability is not guaranteed until the first genuine package upload succeeds.

## Supported platforms

Office document export in version 1.0.0 targets:

- Windows 11
- Microsoft 365 Apps desktop applications
- Office 2024 and Office LTSC 2024
- CPython 3.11 or newer
- `uv` and `uvx`

Microsoft 365 Apps should be the primary test environment. Office 2024 and Office LTSC 2024 should receive formal compatibility testing before release.

Office LTSC 2021 and older desktop releases may still work because the core Word, Excel, and PowerPoint fixed-format COM APIs have existed for many releases. They should be treated as best effort rather than formally supported. Older releases differ in Microsoft lifecycle status, Windows compatibility, PDF implementation, protected-content behavior, graphics filters, rendering results, and application bugs.

Compatibility should be capability based rather than controlled only by an Office version number. `doctor` should report the detected application version and probe the operations needed by the requested conversion. Versions outside formal support should receive a warning when the required capability is present and a targeted error when it is absent. The CLI must not fail solely because an Office version is outside the formal test matrix.

Office document export is not supported on macOS in version 1.0.0. A macOS Office implementation would require separate AppleScript or JXA adapters with a different capability matrix.

PDF to PNG or JPEG rasterization does not require Office or Windows. It should work on any supported Python platform for which the selected `pypdfium2` wheel is available. The CLI must report Office export and PDF rasterization as separate capabilities.

## Input formats

Version 1.0.0 supports only:

- Word: `.docx`, `.doc`
- Excel: `.xlsx`, `.xls`
- PowerPoint: `.pptx`, `.ppt`
- PDF rasterization: `.pdf` to PNG or JPEG

All other input formats must return an unsupported-format error in version 1.0.0, even when the installed Office application could open them.

Legacy `.doc`, `.xls`, and `.ppt` files are part of the tested 1.0.0 surface. They do not add a new rendering architecture because the installed Office application opens them. They do add compatibility, repair, password, and security test cases. Legacy binary files can contain macros without a macro-specific filename suffix.

All files must open with macros disabled unless a future explicitly unsafe option is added. Version 1.0.0 should not include an option to execute macros.

## Output formats

### PDF

PDF output should use each application's native fixed-format export API.

- Word uses `Document.ExportAsFixedFormat`.
- Excel uses workbook, worksheet, or range `ExportAsFixedFormat` as appropriate.
- PowerPoint uses `Presentation.ExportAsFixedFormat`.

When exporting the complete source without selection or post-processing, the native Office PDF should be returned unchanged.

Native PDF quality is a version 1.0.0 requirement. Default exports should:

- Use the application's print or standard quality setting.
- Include document properties and metadata.
- Include document structure tags for accessibility when Office supports them.
- Preserve hyperlinks emitted by Office.
- Preserve internal destinations emitted by Office.
- Use bitmap text when a font cannot legally be embedded and Office supports that option.
- Avoid PDF/A unless the user requests it.
- Produce an unencrypted PDF without copy, print, or modification restrictions.

Word bookmark export has an Office limitation. Its native API can create PDF bookmarks from headings or from explicit Word bookmarks, but not both in one setting. Expose `--bookmarks headings|word|none` and default to `headings`.

A Word table of contents is document content rather than the PDF bookmark tree. Updating a stale table of contents before export changes the in-memory document and can affect page numbers. The tool will update each Word table of contents in memory before export without saving the source. It will not update arbitrary external fields by default. Expose `--no-update-toc` for cases where the caller needs to preserve the current pagination and displayed TOC values exactly.

Microsoft documents that internal links within an Excel workbook are lost during PDF conversion, while external links remain. Version 1.0.0 must test and document this native Office limitation. Reconstructing internal Excel links would require a separate coordinate-mapping and PDF-annotation feature. That work is deferred until after version 1.0.0.

### PNG and JPEG

Word and Excel do not expose a reliable whole-page PNG or JPEG export API. Their native PDF will be rasterized into page images.

PowerPoint exposes native slide PNG and JPEG export. Version 1.0.0 should use the PDF rasterization path by default so DPI, encoding, naming, page selection, notes pages, and handouts behave consistently across all applications.

Version 1.0.0 will also include the PowerPoint-native image engine:

```text
--image-engine pdfium
--image-engine office
```

`pdfium` is the default. `office` is valid only for PowerPoint slide output. It does not apply to notes, outlines, or handouts. Help and JSON capability output must make that distinction clear.

### PDF source rasterization

Version 1.0.0 will accept PDF input for PNG and JPEG output:

```powershell
uvx office-export document.pdf --to png
uvx office-export document.pdf --to jpeg --pages 1,3-5 --dpi 300
```

Implications:

- No Office application is needed for PDF input.
- PDF rasterization can work cross-platform even though Office export is Windows-only.
- PDF input provides direct fixtures for testing the rasterizer without Office.
- PDF page selection uses one-based physical page indices. `inspect` should also report PDF page labels when present.
- Page rotation and visible crop boxes must be honored.
- Encrypted or permission-restricted PDFs must be rejected in version 1.0.0.
- The tool must not accept passwords or bypass PDF restrictions in version 1.0.0.
- Render visible annotations and standard form field appearances by default so the image matches a typical viewer display.
- Do not execute PDF JavaScript, launch actions, submit actions, or other active document behavior.
- Expose `--exclude-annotations` for callers that need only the underlying page content. This option also excludes form widget appearances because PDF form widgets are annotations.
- Warn when dynamic XFA forms or other unsupported interactive content may not match a full PDF viewer.
- PDF input supports PNG and JPEG output in version 1.0.0. PDF-to-PDF rewriting is out of scope.
- Generic PDFs can contain unusually large pages, malformed structures, complex transparency, forms, and annotations. Existing timeout and pixel limits apply.
- PNG and JPEG DPI metadata should be written consistently.

### PDF rasterizer

Use `pypdfium2` for PDF rendering and Pillow for PNG and JPEG encoding.

Reasons:

- Prebuilt Windows wheels are available.
- PDFium is a mature PDF rendering engine.
- `pypdfium2` integrates with Pillow.
- It does not require a separate Poppler or Ghostscript installation.
- The binding is available under Apache 2.0 or BSD 3-Clause terms.
- PDFium uses a BSD-style license, with additional notices for bundled dependencies.

Tentative dependency bounds for initial development:

```toml
pypdfium2 = ">=5.13,<6"
pillow = ">=11,<13"
```

The `uv.lock` file should lock the exact PDFium wheel used for reproducible development and testing.

Rendering scale is calculated as:

```python
scale = dpi / 72
```

Rasterization requirements:

- Render one page at a time.
- Explicitly close PDFium pages, bitmaps, and documents.
- Do not call PDFium concurrently from multiple threads.
- Use processes instead of threads if raster work is parallelized later.
- Enforce a maximum output pixel count before allocating the bitmap.
- Convert JPEG output to RGB and flatten transparency onto a configurable background.
- Default the JPEG background to white.
- Preserve PNG alpha only when the rendered PDF page actually has meaningful transparency.
- Respect PDF page rotation and page dimensions.
- Use the visible crop box by default for generic PDF input.
- Render visible annotations and standard form widget appearances by default.
- Do not execute JavaScript or other PDF actions while rendering.
- Support `--exclude-annotations` for clean page-content rendering.
- Call `PdfDocument.init_forms()` immediately after opening a PDF when forms are present, before reading page handles or the page count.
- Use `draw_annots=True` and `may_draw_forms=True` for the default viewer-like path.
- Use `draw_annots=False` and `may_draw_forms=False` when `--exclude-annotations` is set.
- Include the PDFium and Pillow versions in the output manifest.

Recommended initial defaults:

- Image resolution: 150 DPI
- JPEG quality: 92
- JPEG background: white
- Maximum image size: 50 megapixels per page

These defaults should be verified against real Word pages, Excel print layouts, standard slides, large-format slides, and unusually large worksheet paper sizes.

## CLI design

The primary command should remain concise:

```powershell
uvx office-export report.docx --to pdf
uvx office-export report.docx --to png --pages 1,3-5
uvx office-export deck.pptx --to jpeg --slides 2-6 --dpi 200
uvx office-export model.xlsx --to pdf --sheet Summary --sheet "Q4 Charts"
uvx office-export model.xlsx --to png --range "Summary!A1:H40"
uvx office-export document.pdf --to jpeg --pages 1,3-5 --dpi 300
```

Supporting commands:

```powershell
uvx office-export inspect INPUT
uvx office-export doctor
uvx office-export formats
uvx office-export batch PATH
```

The package name and console command should both use `office-export` if the PyPI name is available. The import package should be `office_export`.

### Common options

- `--to pdf|png|jpeg`
- `--output PATH`
- `--force`
- `--dpi INTEGER`
- `--jpeg-quality INTEGER`
- `--background COLOR`
- `--quality screen|print`
- `--timeout SECONDS`
- `--keep-intermediate`
- `--json`
- `--manifest PATH`
- `--verbose`
- `--image-engine pdfium|office`

### Word options

- `--pages 1,3-5`
- `--include-markup`
- `--bookmarks none|headings|word`
- `--pdf-a`
- `--no-update-toc`

Word page numbers refer to the pagination produced by the installed Word version, available fonts, document compatibility settings, and active printer configuration.

### PowerPoint options

- `--slides 1,3-5`
- `--include-hidden`
- `--output-type slides|notes|outline|handout1|handout2|handout3|handout4|handout6|handout9`
- `--frame-slides`
- `--pdf-a`

Slides are selected by one-based presentation position. The inspection response should also expose stable PowerPoint slide IDs.

### Excel options

- Repeated `--sheet NAME_OR_INDEX`
- `--range "Sheet!A1:H40"`
- Repeated `--chart "Sheet!Chart Name"`
- `--charts all`
- `--ignore-print-area`
- `--show-formulas`
- `--show-headings`
- `--recalculate never|auto|full`
- `--update-links`
- `--refresh-data`

Version 1.0.0 defaults:

- Do not update external links.
- Do not refresh external data.
- Do not execute macros.
- Avoid recalculation unless the user requests it.
- Open the source read-only.
- Do not save changes back to the source workbook.

An Excel sheet can produce multiple printed pages. Image names and manifests must retain both the sheet identity and the printed-page position.

### Excel chart-only image export

Version 1.0.0 will create one tightly bounded image per selected Excel chart for reuse in Markdown, HTML, presentations, and other documents. It will support both embedded chart objects on worksheets and dedicated chart sheets.

Commands:

```powershell
uvx office-export model.xlsx --to png --charts all
uvx office-export model.xlsx --to png --chart "Summary!Revenue Chart"
uvx office-export model.xlsx --to jpeg --chart "Dashboard!Margin" --jpeg-quality 95
```

Behavior:

- `inspect` lists each chart's worksheet or chart sheet, object name, title, type, visibility, and dimensions.
- Repeated `--chart SHEET!NAME` selects specific charts. `--charts all` selects all visible charts.
- Use Excel's native `Chart.Export` to create a tightly bounded PNG at the chart's current dimensions.
- Produce JPEG by converting the temporary native PNG through Pillow so JPEG quality and background handling remain deterministic.
- Do not use worksheet PDF rasterization for chart-only images because it would include paper size, margins, cells, and surrounding worksheet content.
- Chart-only selection supports PNG and JPEG output only. Using `--chart` or `--charts` with `--to pdf` returns a targeted invalid-option error.
- Treat `--dpi` as inapplicable to the native chart path. Excel's native graphic export does not expose a DPI parameter.
- Consider a later explicit pixel-width or scale option. Any high-resolution technique that temporarily resizes or duplicates a chart must operate only in memory and must never save workbook changes.
- Preserve the existing safe defaults for calculation, external links, data refresh, macros, and read-only opening.
- Return a logical mapping from the selected chart identity to each output image.

Native chart dimensions are the version 1.0.0 contract. Precise high-resolution sizing should not be promised until a visual and dimensional spike verifies Excel behavior across supported Office releases.

### PDF input options

- `--pages 1,3-5`
- `--exclude-annotations`

PDF pages are selected by one-based physical page index. Page labels are informational and are exposed through `inspect`.

## Selection and PDF fidelity

Use native Office selection controls when they can express the user's request without modifying the source file.

The full native Office PDF must not be rewritten. This preserves the metadata, tagged structure, hyperlinks, destinations, and bookmarks that Office emitted.

Selected PDF output should use native Office selection APIs. Post-processing a PDF can drop or damage bookmarks, named destinations, internal hyperlinks, and the tagged-PDF structure.

Application behavior:

- PowerPoint should use native print ranges.
- Word should use native contiguous page ranges.
- Excel should use native workbook, worksheet, or range export APIs.
- Noncontiguous image selection may rasterize selected pages from a complete temporary PDF without rewriting that PDF.
- Noncontiguous Word PDF selection should produce separate native PDFs or return a clear unsupported-selection error in version 1.0.0.
- Combining independently exported PDFs is not permitted unless preservation of metadata, bookmarks, links, destinations, and tags has been verified.

`pypdf` is not required for the initial architecture. Add a PDF manipulation library only when a tested feature requires it.

## Output naming

Default output should never overwrite an existing file or directory without `--force`.

Default placement and names:

- PDF output is written beside the input with the same base filename and a `.pdf` extension.
- PNG output is written into a sibling directory named `<base filename> - PNG export`.
- JPEG output is written into a sibling directory named `<base filename> - JPEG export`.
- An explicit `--output PATH` overrides these defaults.

```text
report.pdf
report - PNG export\report-page-001.png
report - PNG export\report-page-002.png
deck - JPEG export\deck-slide-003.jpg
model - PNG export\model-sheet-summary-page-001.png
model - PNG export\model-sheet-summary-chart-revenue.png
```

Names must be deterministic, filesystem-safe, and independent of the Office UI language.

Default image output always uses an output directory, including when selection produces one image. An explicit `--output` may name a single image file when exactly one image is selected.

All outputs should be written to a temporary path and atomically moved into place only after successful completion.

## Inspection and machine-readable output

`inspect` should open the source using the same safe defaults as export.

Word inspection should include:

- Page count
- Document title
- Section count
- Page dimensions
- Tracked-change presence
- Comment presence
- Protection state

PowerPoint inspection should include:

- Slide count
- Slide index and slide ID
- Slide name and title
- Hidden status
- Slide dimensions
- Notes presence

Excel inspection should include:

- Worksheet and chart-sheet names
- Sheet position and visibility
- Used range
- Print area
- Page orientation and paper size
- Estimated printed-page count
- External link presence
- Connection presence
- Calculation mode

`--json` should return a stable schema for agents and scripts. Human-readable output should be generated from the same underlying result model.

## Conversion result and optional manifest

A manifest is a JSON audit record about a conversion. It is separate from the PDF or image output. It records what was converted, how it was converted, warnings, and which files were created.

Version 1.0.0 behavior:

- `--json` prints the full conversion result to stdout without creating a sidecar file.
- Human-readable mode prints a concise summary.
- `--manifest PATH` opts in to persisting the same result as a JSON file.
- No sidecar manifest is written by default.
- The bundled skill should normally use `--json`.

This provides agents with structured provenance without cluttering ordinary output directories or unexpectedly writing absolute source paths and hashes to disk.

Each conversion should be able to emit a JSON manifest containing:

- Tool version
- Timestamp
- Source path
- Source size
- Source SHA-256 hash
- Detected application
- Source format
- Office application name and version
- Active printer name when relevant
- Export options
- Security and compatibility warnings
- Temporary PDF use
- PDFium version
- Pillow version
- Output file paths
- Output sizes and hashes
- Logical source unit to output-page mapping
- Duration
- Success or failure status

## Office automation architecture

### Main process

The main CLI process should:

- Parse and validate arguments.
- Detect the required application adapter.
- Create a job request.
- Start a dedicated worker subprocess.
- Enforce a timeout.
- Receive structured progress and results.
- Validate completed outputs.
- Move outputs into their final locations.
- Print human-readable or JSON output.

### Worker process

Use a fresh worker subprocess for each input document.

The worker should:

- Initialize COM as a single-threaded apartment.
- Start or acquire an Office application instance in a controlled way.
- Record whether it created the Office instance.
- Disable alerts and active content where the application permits it.
- Open the source read-only.
- Export to an isolated temporary directory.
- Close the document explicitly.
- Quit only an Office application instance owned by the worker.
- Release COM proxies explicitly.
- Call `CoUninitialize`.
- Exit completely after the job.

The implementation must never terminate a pre-existing user Office process. If process termination is needed after a timeout, it must target only a process proven to have been created and owned by the worker.

### Adapter protocol

```python
class OfficeAdapter(Protocol):
    def inspect(self, request: InspectRequest) -> InspectResult: ...
    def export_pdf(self, request: ExportRequest, output: Path) -> ExportResult: ...
    def capabilities(self) -> CapabilityResult: ...
```

Potential future hook:

```python
class NativeImageExporter(Protocol):
    def export_images(self, request: ExportRequest, output_dir: Path) -> ExportResult: ...
```

## Printer handling

An installed printer should not be a universal hard requirement.

`doctor` should check:

- The Windows Print Spooler state
- Installed printer drivers
- The active Word and Excel printer where detectable
- Whether Microsoft Print to PDF is available

Behavior:

- Do not block PowerPoint because no printer is installed.
- Attempt Word and Excel native export normally.
- Return a targeted printer diagnostic if pagination or export fails and no usable printer exists.
- Do not silently set or change the user's printer.
- Record the active printer in manifests because printer metrics can affect Word and Excel pagination.

## Security defaults

Office documents can contain active content and external references. The tool must default to the safest practical behavior.

- Open source documents read-only.
- Force-disable VBA macros.
- Do not execute AutoOpen or Auto_Open macros.
- Do not update external links.
- Do not refresh data connections.
- Do not save changes back to the source.
- Do not bypass Protected View, passwords, IRM, or sensitivity labels.
- Do not strip sensitivity protection as an incidental part of rasterization.
- Refuse anonymous overwrite without `--force`.
- Warn that Office automation is not a security sandbox.

Password support, if added, should accept secrets from standard input or a protected file descriptor. Passwords should not be accepted directly on the command line where they can appear in process listings and shell history.

## Failure handling

Expected structured failure categories:

- Unsupported source format
- Required Office application unavailable
- Source file missing
- Output path invalid
- Output already exists
- Source password required
- Source protected or restricted
- Protected View blocked automation
- Missing or restricted fonts
- External content blocked
- No usable printer for required pagination
- Office modal dialog or timeout
- Office COM error
- PDF export failure
- PDF rasterization failure
- Output exceeds megapixel limit
- Partial output cleanup failure

The default CLI should show a concise error. `--verbose` should add COM HRESULT values, application context, and worker diagnostics without exposing secrets.

## `doctor` command

`doctor` should report:

- Windows version
- Python version
- Tool version
- Installed Word version
- Installed Excel version
- Installed PowerPoint version
- Office bitness where detectable
- Print Spooler state
- Installed printers
- Microsoft Print to PDF availability
- PDFium version
- Pillow version
- Temporary-directory write test
- Optional smoke export for each installed application

Smoke tests should require explicit source fixtures or use small fixtures shipped specifically for diagnostics.

## Batch behavior

Batch conversion should be sequential in version 1.0.0.

- Default `--jobs 1`
- Do not run multiple Office conversions concurrently.
- Do not run PDFium calls concurrently from multiple threads.
- Continue-on-error should be explicit.
- Each document gets its own worker subprocess and temporary directory.

Future process-based parallel rendering can be evaluated independently from Office automation concurrency.

## Bundled agent skill

The skill must ship inside the Python package, following the patterns reviewed in `markdown-pptx`, `imap-agent-cli`, `sql-agent-cli`, and `azwi`.

Implementation pattern:

- Define the managed skill text in `src/office_export/skill.py`.
- Use `SKILL_NAME = "office-export"`.
- Use a marker such as `<!-- managed-by: office-export -->`.
- Install by default to `~/.agents/skills/office-export/SKILL.md`.
- Make installation idempotent and replace stale managed content.
- Refuse to overwrite or remove unmanaged skill content without an explicit force option.
- Keep the skill platform-neutral and agent-tool-neutral.
- Teach the skill to invoke the published CLI as `uvx office-export`.
- Keep diagnostics on stderr and structured payloads on stdout.

CLI commands:

```text
uvx office-export skill install
uvx office-export skill remove
uvx office-export skill install --skills-dir PATH
```

The skill should teach an agent to:

1. Run `doctor` when capabilities are unknown.
2. Run `inspect --json` before making nontrivial selections.
3. Choose page, slide, sheet, or range selectors.
4. Export the requested PDF or images.
5. Inspect representative image outputs or create a contact sheet.
6. Report warnings, selections, and exact output paths.
7. Avoid unsafe macro, refresh, or link behavior.

The skill should call the CLI. It should not contain its own COM automation implementation.

Tests should cover the embedded skill text, managed marker, installation, idempotent updates, guarded removal, force removal, custom skill roots, and JSON command output.

## Testing strategy

### Unit tests

- CLI argument parsing
- Selector grammar
- Capability resolution
- Output naming
- Manifest schema
- Manifest opt-in behavior
- Error mapping
- Pixel-limit calculations
- DPI calculations
- JPEG background and encoding behavior
- Mocked adapter lifecycle
- Worker protocol

### Integration tests

Use real installed Office applications with fixtures covering:

- Standard and custom page sizes
- Portrait and landscape Word sections
- Word comments and tracked changes
- Word table-of-contents updating without saving the source
- Missing and restricted fonts
- PowerPoint standard and widescreen slides
- Hidden slides
- Speaker notes and handouts
- Excel print areas
- Multi-page worksheets
- Hidden and very hidden sheets
- Charts and chart sheets
- Chart-only PNG and JPEG output at native chart dimensions
- Formulas and cached values
- External links and data connections
- Macro-enabled files with macros disabled
- Password and protection failures
- PDF annotations and standard form appearances
- PDF rendering with `--exclude-annotations`
- Dynamic XFA warning behavior
- Long paths, Unicode paths, and OneDrive paths

### Visual regression tests

Compare rendered images perceptually rather than byte for byte.

- Maintain golden fixtures for Word, Excel, and PowerPoint.
- Compare PDFium output across supported DPI settings.
- Test thin lines, gradients, transparency, charts, tables, and font rendering.
- Compare PDFium PowerPoint images with native `Slide.Export` during development.
- Define a review workflow for expected changes after Office or PDFium updates.

Office integration tests will require a Windows machine with licensed Office installed. They should not be assumed to run on ordinary hosted CI workers.

## Release milestones

### Milestone 1: Project foundation

- `uv` project and package metadata
- CLI entry point
- Typed request and result models
- Structured errors
- Worker subprocess protocol
- `doctor`
- Basic JSON output

### Milestone 2: Native PDF export

- Word adapter
- Excel adapter
- PowerPoint adapter
- Native PowerPoint PNG and JPEG engine
- Safe Office lifecycle
- Timeout handling
- Atomic output writes
- PDF integration fixtures

### Milestone 3: Raster image output

- `pypdfium2` integration
- Pillow integration
- PNG and JPEG output
- DPI and JPEG controls
- Pixel and memory limits
- Deterministic naming
- Page selection
- Direct PDF input
- Excel chart-only PNG and JPEG export

### Milestone 4: Inspection and selections

- Word page inspection and selection
- PowerPoint slide inspection and selection
- Excel sheet and range inspection and selection
- Logical source-unit mappings
- Manifest generation

### Milestone 5: Reliability and agent use

- Batch conversion
- Failure recovery
- Visual regression fixtures
- Documentation
- Bundled agent skill
- `uvx` installation verification
- Version 1.0.0 release checklist

## PDF rasterizer concerns and mitigations

`pypdfium2` remains the recommended rasterizer. The concerns are manageable.

### Thread safety

PDFium is not thread-safe. Concurrent PDFium calls across threads can crash or corrupt the process.

Mitigation:

- Keep version 1.0.0 rasterization sequential.
- Use a mutex if PDFium can ever be reached from more than one thread.
- Prefer worker processes for future parallel rendering.

Reference: <https://pypdfium2.readthedocs.io/en/stable/python_api.html#incompatibility-with-threading>

### Native resource lifetime

PDFium objects hold native resources and file handles. Garbage collection is not sufficiently prompt for large batches.

Mitigation:

- Use context managers where available.
- Close documents, pages, and bitmaps explicitly.
- Render and save one page at a time.
- Add batch memory tests.

### Memory use

Large Excel paper sizes or high DPI can create very large bitmaps. A 100-megapixel RGBA bitmap can require roughly 400 MB before additional encoder overhead.

Mitigation:

- Calculate output dimensions before rendering.
- Default to a 50-megapixel page limit.
- Allow an explicit override with a warning.
- Keep only one page bitmap in memory at a time.

### Rendering differences

The PDF is rendered by Office, then rasterized by PDFium. Small pixel differences from PowerPoint's direct slide rasterizer are possible.

Mitigation:

- Use visual golden tests.
- Keep the native PowerPoint image engine as an optional comparison and fallback path.
- Record the PDFium version in manifests.
- Treat PDF output as the layout source of truth for the default image engine.

### Licensing and notices

`pypdfium2` is permissively licensed, but its wheels contain PDFium and third-party dependencies with their own notices.

Mitigation:

- Review the exact wheel licenses before the first public release.
- Include required third-party notices in packaged distributions.
- Lock and audit dependency versions.
- Do not describe the full binary bundle simply as Apache 2.0 without qualification.

Reference: <https://github.com/pypdfium2-team/pypdfium2#licensing>

### API stability

The high-level support API is still described as beta and may change across major releases.

Mitigation:

- Pin to a tested major release.
- Wrap PDFium behind a small internal rasterizer interface.
- Test dependency upgrades separately.

### Protected PDF output and input

Office may create a protected or encrypted PDF when source sensitivity labels or IRM require it. Rasterization may fail or may be prohibited by policy.

Mitigation:

- Version 1.0.0 must not emit protected or encrypted PDF outputs.
- Verify that generated PDFs are unencrypted and have no copy, print, or modification restrictions.
- If Office emits a protected PDF, delete the temporary output and return a clear protection error.
- Do not remove protection from the generated PDF after export.
- Do not bypass protection on an Office source or PDF source.
- Reject encrypted or permission-restricted PDF inputs.

## Deferred features

- Markdown output
- Semantic HTML output
- Static HTML document galleries
- Visio support
- OneNote support
- macOS Office automation
- LibreOffice backend
- Server-side conversion service
- Macro execution
- External data refresh by default
- PowerPoint video or animated GIF output
- SVG output
- OCR
- PDF-to-PDF rewriting
- Excel internal hyperlink reconstruction
- Precise high-resolution Excel chart sizing
- Dynamic XFA form rendering

## References

- Microsoft Word fixed-format export: <https://learn.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat>
- Microsoft Excel worksheet fixed-format export: <https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.exportasfixedformat>
- Microsoft PowerPoint fixed-format export: <https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.exportasfixedformat>
- Microsoft PowerPoint slide export: <https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.export>
- Microsoft Excel chart image export: <https://learn.microsoft.com/en-us/office/vba/api/excel.chart.export>
- Microsoft Office lifecycle matrix: <https://learn.microsoft.com/en-us/lifecycle/office-windows-configuration-matrix>
- Microsoft guidance on server-side Office automation: <https://support.microsoft.com/en-us/office/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2>
- OfficeToPDF: <https://github.com/cognidox/OfficeToPDF>
- `pypdfium2`: <https://github.com/pypdfium2-team/pypdfium2>
- `pypdfium2` Python API: <https://pypdfium2.readthedocs.io/en/stable/python_api.html>
