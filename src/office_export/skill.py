from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

from office_export.errors import UsageError

SKILL_NAME = "office-export"
MANAGED_MARKER = "<!-- managed-by: office-export -->"

SKILL_MD = f"""---
name: office-export
description: Export Word, Excel, PowerPoint, and PDF files to faithful PDF, PNG, or JPEG output with `uvx office-export`. Use when an agent needs to inspect or render Office documents, select pages, slides, sheets, ranges, or charts, or create local previews through desktop Microsoft Office.
---

{MANAGED_MARKER}

# Office Export

Use the published CLI through `uvx office-export`. Always invoke the tool as `uvx office-export ...`. Do not assume that a bare command is installed globally.

Office export requires Windows, an interactive user session, and installed desktop Microsoft Office. PDF rasterization works anywhere the packaged PDFium wheel is supported. The tool is not a server-side or truly headless Office conversion service.

## Fast path for explicit exports

When the input, output format, and selection are explicit, run the export directly. Use the user's output path or choose a deterministic output path when one was not provided.

User request: "Get a PNG screenshot of slide 3 from deck.pptx."

```text
uvx office-export deck.pptx --to png --slides 3 --output <planned-output-path> --json
```

This request should normally require one export command. Do not add capability checks, source inspection, duplicate exports, or visual review.

## Use discovery only when needed

- Run `doctor --json` only when diagnosing an environment or capability failure.
- Run `formats --json` only when format support is uncertain.
- Run `inspect <input> --json` only when metadata is needed to resolve the request. Examples include discovering page counts, slide numbers, worksheet names, ranges, or chart identities.
- Do not inspect merely to confirm an explicit one-based selection before exporting.

Discovery commands are available when those conditions apply:

```text
uvx office-export doctor --json
uvx office-export formats --json
uvx office-export inspect report.docx --json
uvx office-export inspect deck.pptx --json
uvx office-export inspect model.xlsx --json
uvx office-export inspect document.pdf --json
```

Selectors are one-based unless a sheet or chart is selected by name.

## Export examples

```text
uvx office-export report.docx --to pdf --json
uvx office-export deck.pptx --to pdf --slides 2-6 --json
uvx office-export model.xlsx --to pdf --sheet Summary --json
uvx office-export report.docx --to png --pages 1,3-5 --dpi 200 --json
uvx office-export deck.pptx --to jpeg --slides 2-6 --jpeg-quality 95 --json
uvx office-export document.pdf --to png --pages 1-4 --exclude-annotations --json
uvx office-export model.xlsx --to png --sheet Summary --json
uvx office-export model.xlsx --to png --range "Summary!A1:H40" --json
uvx office-export model.xlsx --to png --charts all --json
uvx office-export model.xlsx --to jpeg --chart "Dashboard!Revenue" --json
```

The tool preserves complete native Office PDFs unchanged. Noncontiguous Word page selection is not supported for PDF output. PNG and JPEG output defaults to a new sibling directory. PDFium is the default image engine. PowerPoint slide output can instead use `--image-engine office`. Chart-only Excel export uses the chart's native dimensions, so `--dpi` does not apply.

## Review and report results

- Visually inspect output when the user requests visual quality assurance, fidelity validation, layout review, or troubleshooting.
- Consider visual review for complex, high-risk, or multi-file exports where rendering defects are plausible.
- Do not visually inspect a straightforward screenshot or conversion when the export succeeds and returns a valid output file.
- Do not re-export or copy an artifact solely because a preview tool cannot open it.
- Treat successful structured output as sufficient for an ordinary conversion. Report the selected pages, slides, sheets, or charts, the exact absolute output path, and every warning.
- On failure, use the structured error code and message to choose the next action. Do not repeat the same command unchanged. Keep retries proportional to the request.

## Preserve safety defaults

The tool opens Office sources read-only. It disables macros and does not update links, refresh data, or save source changes by default. Do not add `--update-links`, `--refresh-data`, or recalculation options unless the user explicitly needs the resulting content change and understands the external access involved.

Never attempt to bypass passwords, Protected View, IRM, sensitivity labels, or PDF permissions. `--force` authorizes replacement only at the planned output paths.
"""


def default_skills_dir() -> Path:
    return Path.home() / ".agents" / "skills"


def skill_dir(skills_dir: Path | None = None) -> Path:
    return (skills_dir or default_skills_dir()) / SKILL_NAME


def install_skill(skills_dir: Path | None = None, *, force: bool = False) -> dict[str, Any]:
    target = skill_dir(skills_dir)
    skill_path = target / "SKILL.md"
    if target.exists() and not target.is_dir():
        raise UsageError(f"Skill target is not a directory: {target}")
    if target.exists() and not skill_path.exists() and not force:
        raise UsageError(f"Refusing to install into '{target}' because it contains no managed SKILL.md.")
    if skill_path.exists() and MANAGED_MARKER not in skill_path.read_text(encoding="utf-8") and not force:
        raise UsageError(f"Refusing to overwrite unmanaged skill file '{skill_path}'.")
    target.mkdir(parents=True, exist_ok=True)
    existed = skill_path.exists()
    previous = skill_path.read_text(encoding="utf-8") if existed else ""
    updated = existed and previous != SKILL_MD
    skill_path.write_text(SKILL_MD, encoding="utf-8", newline="\n")
    return {
        "installed": True,
        "created": not existed,
        "updated": updated,
        "skill": SKILL_NAME,
        "path": str(skill_path.resolve()),
    }


def remove_skill(skills_dir: Path | None = None, *, force: bool = False) -> dict[str, Any]:
    target = skill_dir(skills_dir)
    skill_path = target / "SKILL.md"
    if not target.exists():
        return {"removed": False, "skill": SKILL_NAME, "path": str(target), "reason": "not_installed"}
    if not target.is_dir():
        raise UsageError(f"Skill target is not a directory: {target}")
    if not skill_path.exists() and not force:
        raise UsageError(f"Refusing to remove '{target}' because SKILL.md is missing.")
    if skill_path.exists():
        content = skill_path.read_text(encoding="utf-8")
        if MANAGED_MARKER not in content and not force:
            raise UsageError(
                f"Refusing to remove '{target}' because it is not marked as managed by office-export. "
                "Use --force to override."
            )
    extra_paths = [path for path in target.iterdir() if path.name != "SKILL.md"]
    if extra_paths and not force:
        names = ", ".join(sorted(path.name for path in extra_paths))
        raise UsageError(
            f"Refusing to remove '{target}' because it contains unmanaged entries: {names}. Use --force to override."
        )
    shutil.rmtree(target)
    return {"removed": True, "skill": SKILL_NAME, "path": str(target)}
