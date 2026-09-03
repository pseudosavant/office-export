# Changelog

## 0.2.0

- Synchronize older pristine managed skills locally during normal CLI invocations, including help and version output.
- Store the CLI version and normalized content hash in `SKILL.md` front matter. Migrate legacy managed skills and recover missing or invalid version metadata.
- Preserve edited or unverifiable skills and newer installed versions. Limit install-time `--force` to recognized managed skills.
- Add read-only `skill status` with plain and JSON diagnostics. Keep maintenance notices on stderr.
- Exclude local source and editable builds from automatic synchronization. Custom skill directories require explicit updates.

## 0.1.1

- Give explicit exports a one-command skill fast path without routine capability checks, source inspection, or visual review.
- Make discovery, retries, and visual validation conditional on the request or a structured failure.

## 0.1.0

- Export Word, Excel, and PowerPoint documents through desktop Microsoft Office on Windows.
- Produce native PDF output plus PNG and JPEG images rendered through PDFium.
- Rasterize PDF inputs without requiring Office.
- Inspect supported documents and report stable JSON results.
- Select Word pages, PowerPoint slides, Excel sheets, ranges, and charts.
- Apply safe defaults for macros, external links, data refresh, recalculation, and source writes.
- Isolate Office automation in a worker subprocess with timeouts and owned-process cleanup.
- Add sequential batch conversion, capability diagnostics, deterministic output naming, and optional manifests.
- Add managed `skill install` and `skill remove` commands for coding agents.
