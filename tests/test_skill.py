from __future__ import annotations

from pathlib import Path

import pytest

from office_export.errors import UsageError
from office_export.skill import SKILL_MD, install_skill, remove_skill


def test_skill_fast_path_exports_explicit_slide_in_one_command() -> None:
    expected_command = "uvx office-export deck.pptx --to png --slides 3 --output <planned-output-path> --json"
    fast_path = SKILL_MD.split("## Fast path for explicit exports", 1)[1].split("## Use discovery only when needed", 1)[
        0
    ]

    assert expected_command in fast_path
    assert "normally require one export command" in fast_path
    assert "doctor --json" not in fast_path
    assert "formats --json" not in fast_path
    assert "inspect deck.pptx" not in fast_path
    assert "visual review" in fast_path


@pytest.mark.parametrize(
    ("command", "condition"),
    [
        ("doctor --json", "only when diagnosing an environment or capability failure"),
        ("formats --json", "only when format support is uncertain"),
        ("inspect <input> --json", "only when metadata is needed to resolve the request"),
    ],
)
def test_skill_discovery_commands_are_conditional(command: str, condition: str) -> None:
    assert command in SKILL_MD
    assert condition in SKILL_MD


def test_skill_visual_review_is_conditional() -> None:
    assert "when the user requests visual quality assurance" in SKILL_MD
    assert "Do not visually inspect a straightforward screenshot or conversion" in SKILL_MD
    assert "Do not re-export or copy an artifact solely because a preview tool cannot open it" in SKILL_MD


def test_skill_install_update_idempotence_and_remove(tmp_path: Path) -> None:
    root = tmp_path / "skills"
    created = install_skill(root)
    skill_file = root / "office-export" / "SKILL.md"
    assert created["created"] is True
    assert created["updated"] is False
    assert skill_file.read_text(encoding="utf-8") == SKILL_MD

    unchanged = install_skill(root)
    assert unchanged["created"] is False
    assert unchanged["updated"] is False

    skill_file.write_text("<!-- managed-by: office-export -->\noutdated\n", encoding="utf-8")
    updated = install_skill(root)
    assert updated["updated"] is True
    assert skill_file.read_text(encoding="utf-8") == SKILL_MD

    removed = remove_skill(root)
    assert removed["removed"] is True
    assert not skill_file.parent.exists()
    assert remove_skill(root)["reason"] == "not_installed"


def test_skill_guards_unmanaged_content_and_force_is_explicit(tmp_path: Path) -> None:
    root = tmp_path / "skills"
    target = root / "office-export"
    target.mkdir(parents=True)
    skill_file = target / "SKILL.md"
    skill_file.write_text("user content\n", encoding="utf-8")

    with pytest.raises(UsageError, match="unmanaged"):
        install_skill(root)
    install_skill(root, force=True)
    (target / "notes.txt").write_text("keep", encoding="utf-8")
    with pytest.raises(UsageError, match="unmanaged entries"):
        remove_skill(root)
    remove_skill(root, force=True)
    assert not target.exists()
