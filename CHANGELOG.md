# Changelog

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
