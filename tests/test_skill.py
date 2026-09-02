from __future__ import annotations

from pathlib import Path

import pytest

from office_export.errors import UsageError
from office_export.skill import SKILL_MD, install_skill, remove_skill


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
