# AGENTS.md

This repository contains `office-export`, a Windows-first Python CLI that exports Word, Excel, PowerPoint, and PDF files to PDF, PNG, or JPEG.

## Project identity

- Published package name: `office-export`
- CLI command: `office-export`
- Python package import: `office_export`
- Current release: `0.1.0`

## Working conventions

- Preserve native Office PDF output when no post-processing is required.
- Never save changes back to an Office source document.
- Keep macros disabled. Do not update links or refresh data unless the caller explicitly requests it.
- Use one isolated subprocess per Office input.
- Never terminate or quit a pre-existing user Office process.
- Keep PDFium work sequential.
- Keep human output concise and JSON output stable.
- Do not use em dashes or semicolons in copy or documentation.
- Add tests for CLI, selection, output safety, worker lifecycle, and security behavior.

## Common commands

```powershell
$env:UV_LINK_MODE="copy"
uv sync --all-groups
uv run ruff check .
uv run ruff format --check .
uv run pytest
uv build
uv run twine check dist/*
uvx --refresh --from . office-export --version
```

Office integration tests require Windows, an interactive desktop session, and licensed desktop Office applications. They are excluded from the default test run.

## Release rules

- Keep `pyproject.toml` and `src/office_export/__init__.py` versions in sync.
- Run the locked unit suite on Windows and Linux.
- Run real Word, Excel, and PowerPoint smoke tests on Windows before release.
- Verify the wheel in a clean virtual environment.
