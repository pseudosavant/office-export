from __future__ import annotations

import io
import json
from pathlib import Path

from office_export import __version__
from office_export.cli import build_export_help, build_root_help, main


def test_no_args_and_version() -> None:
    stdout = io.StringIO()
    stderr = io.StringIO()
    assert main([], stdout=stdout, stderr=stderr) == 0
    assert "skill install" in stdout.getvalue()
    assert "document.pdf" in stdout.getvalue()
    assert stderr.getvalue() == ""

    stdout = io.StringIO()
    assert main(["--version"], stdout=stdout, stderr=stderr) == 0
    assert stdout.getvalue() == f"office-export {__version__}\n"


def test_help_describes_full_selection_surface() -> None:
    root_help = build_root_help()
    export_help = build_export_help()
    assert "doctor" in root_help
    assert "batch" in root_help
    assert "--pages" in export_help
    assert "--slides" in export_help
    assert "--sheet" in export_help
    assert "--chart" in export_help
    assert "--exclude-annotations" in export_help


def test_formats_json_has_stable_shape() -> None:
    stdout = io.StringIO()
    stderr = io.StringIO()
    assert main(["formats", "--json"], stdout=stdout, stderr=stderr) == 0
    payload = json.loads(stdout.getvalue())
    assert payload["ok"] is True
    assert payload["schema_version"] == "1.0"
    assert payload["formats"]["inputs"]["pdf"]["outputs"] == ["png", "jpeg"]
    assert payload["formats"]["defaults"]["dpi"] == 150
    assert stderr.getvalue() == ""


def test_skill_cli_round_trip(tmp_path: Path) -> None:
    root = tmp_path / "skills"
    stdout = io.StringIO()
    stderr = io.StringIO()
    assert main(["skill", "install", "--skills-dir", str(root), "--json"], stdout=stdout, stderr=stderr) == 0
    payload = json.loads(stdout.getvalue())
    assert payload["created"] is True
    assert Path(payload["path"]).is_file()

    stdout = io.StringIO()
    assert main(["skill", "remove", "--skills-dir", str(root), "--json"], stdout=stdout, stderr=stderr) == 0
    assert json.loads(stdout.getvalue())["removed"] is True


def test_missing_input_and_unsupported_format_have_json_errors(tmp_path: Path) -> None:
    stdout = io.StringIO()
    stderr = io.StringIO()
    code = main([str(tmp_path / "missing.docx"), "--to", "pdf", "--json"], stdout=stdout, stderr=stderr)
    assert code == 3
    assert json.loads(stdout.getvalue())["error"]["code"] == "source_not_found"
    assert stderr.getvalue() == ""

    source = tmp_path / "notes.txt"
    source.write_text("text", encoding="utf-8")
    stdout = io.StringIO()
    code = main([str(source), "--to", "pdf", "--json"], stdout=stdout, stderr=stderr)
    assert code == 2
    assert json.loads(stdout.getvalue())["error"]["code"] == "unsupported_source_format"


def test_inapplicable_option_errors_are_targeted(tmp_path: Path) -> None:
    source = tmp_path / "doc.pdf"
    source.write_bytes(b"not needed because validation runs first")
    stdout = io.StringIO()
    stderr = io.StringIO()
    code = main([str(source), "--to", "pdf"], stdout=stdout, stderr=stderr)
    assert code == 2
    assert "PDF-to-PDF" in stderr.getvalue()
