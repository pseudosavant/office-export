from __future__ import annotations

import tomllib
from pathlib import Path

from office_export import __version__
from office_export.skill import MANAGED_MARKER, SKILL_MD, SKILL_NAME


def test_package_metadata_is_consistent() -> None:
    root = Path(__file__).resolve().parents[1]
    project = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))["project"]

    assert project["name"] == "office-export"
    assert project["version"] == __version__ == "1.0.0"
    assert project["scripts"]["office-export"] == "office_export.cli:main"
    assert project["license"] == "MIT"


def test_embedded_skill_identity_and_invocations() -> None:
    assert SKILL_NAME == "office-export"
    assert MANAGED_MARKER in SKILL_MD
    assert SKILL_MD.startswith("---\nname: office-export\n")
    command_lines = [
        line.strip()
        for line in SKILL_MD.splitlines()
        if line.strip().startswith(("office-export", "uvx office-export"))
    ]
    assert command_lines
    assert all(line.startswith("uvx office-export") for line in command_lines)
