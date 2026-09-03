from __future__ import annotations

import hashlib
import io
import json
import re
from pathlib import Path
from types import SimpleNamespace

import pytest
import yaml

import office_export
from office_export import cli, skill
from office_export.errors import UsageError


@pytest.fixture
def released(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(skill, "is_local_development", lambda: False)


@pytest.fixture
def rendered(monkeypatch: pytest.MonkeyPatch):
    def render(version: str) -> str:
        with monkeypatch.context() as context:
            context.setattr(office_export, "__version__", version)
            return skill.render_skill()

    return render


def install_text(text: str, root: Path | None = None) -> Path:
    path = skill.skill_dir(root) / "SKILL.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(text.encode("utf-8"))
    return path


def synchronize() -> str:
    stderr = io.StringIO()
    skill.synchronize_skill(stderr=stderr)
    return stderr.getvalue()


def test_canonical_metadata_and_hash(tmp_path: Path) -> None:
    result = skill.install_skill(tmp_path / "skills")
    content = Path(result["path"]).read_bytes()
    text = content.decode("utf-8")
    front = yaml.safe_load(text.split("---", 2)[1])
    fields = front["metadata"]
    assert front["name"] == skill.SKILL_NAME
    assert "version" not in front
    assert fields["managed-by"] == "office-export"
    assert fields["managed-version"] == office_export.__version__
    assert isinstance(fields["managed-version"], str)
    assert f'managed-version: "{office_export.__version__}"' in text
    assert re.fullmatch(r"sha256:[0-9a-f]{64}", fields["managed-content-sha256"])
    blanked = re.sub(r'(managed-content-sha256: )"sha256:[0-9a-f]{64}"', r'\1""', text, count=1)
    assert fields["managed-content-sha256"] == "sha256:" + hashlib.sha256(blanked.encode("utf-8")).hexdigest()
    assert not content.startswith(b"\xef\xbb\xbf")
    assert b"\r" not in content
    assert content.endswith(b"\n")
    assert skill.MANAGED_MARKER not in text
    assert "uvx office-export" in text
    assert sorted(p.name for p in Path(result["path"]).parent.iterdir()) == ["SKILL.md"]


@pytest.mark.parametrize("ending", ["\n", "\r\n", "\r"])
def test_integrity_normalizes_line_endings(ending: str) -> None:
    text = skill.render_skill().replace("\n", ending)
    assert skill.inspect_skill_content(text.encode("utf-8")).integrity == "valid"


@pytest.mark.parametrize(
    ("old", "new"),
    [
        ("# Office Export", "# My Office Export"),
        ("name: office-export", "name: changed"),
        ("description: Export", "description: Modified"),
        ("metadata:\n", "metadata:\n  short-description: Personal notes\n"),
        ("## Export examples", "##  Export examples"),
    ],
)
def test_body_front_matter_and_formatting_changes_are_detected(old: str, new: str) -> None:
    text = skill.render_skill().replace(old, new)
    assert skill.inspect_skill_content(text.encode("utf-8")).integrity == "altered"


def test_hash_replaces_only_metadata_value_without_reserializing_yaml() -> None:
    text = skill.render_skill()
    text = re.sub(r"sha256:[0-9a-f]{64}", "", text, count=1)
    text = text.replace('managed-content-sha256: ""', '"managed-content-sha256":   "" # retained comment')
    text = text.replace("metadata:\n", 'metadata:\n  short-description: "Keep this formatting"\n')
    text += '\nExample: managed-content-sha256: "sha256:' + "a" * 64 + '"\n'
    digest = "sha256:" + hashlib.sha256(text.encode("utf-8")).hexdigest()
    text = text.replace('"managed-content-sha256":   ""', f'"managed-content-sha256":   "{digest}"', 1)
    assert skill.inspect_skill_content(text.encode("utf-8")).integrity == "valid"


@pytest.mark.usefixtures("released")
def test_absent_skill_is_not_installed() -> None:
    assert synchronize() == ""
    assert not skill.default_skills_dir().exists()
    skill.skill_dir().mkdir(parents=True)
    assert synchronize() == ""
    assert not (skill.skill_dir() / "SKILL.md").exists()


@pytest.mark.usefixtures("released")
@pytest.mark.parametrize("owner", [None, "another-tool", "null"])
def test_unmanaged_and_conflicting_ownership_are_preserved(owner: str | None) -> None:
    text = (
        "Personal instructions\n"
        if owner is None
        else (f"---\nmetadata:\n  managed-by: {owner}\n---\n{skill.MANAGED_MARKER}\n")
    )
    path = install_text(text)
    assert synchronize() == ""
    assert path.read_text(encoding="utf-8") == text
    for force in (False, True):
        with pytest.raises(UsageError, match="unmanaged"):
            skill.install_skill(force=force)


@pytest.mark.usefixtures("released")
def test_pristine_older_skill_updates_and_notice_includes_versions_and_path(rendered) -> None:
    path = install_text(rendered("0.1.0"))
    notice = synchronize()
    assert "0.1.0 -> " + office_export.__version__ in notice
    assert str(path) in notice
    assert notice.count("\n") == 1
    assert path.read_bytes() == skill.render_skill().encode("utf-8")


@pytest.mark.usefixtures("released")
@pytest.mark.parametrize("damage", ["body", "missing", "malformed", "uppercase", "mismatch"])
def test_older_altered_or_unverifiable_skill_requires_force(rendered, damage: str) -> None:
    text = rendered("0.1.0")
    if damage == "body":
        text += "Personal instructions\n"
    elif damage == "missing":
        text = re.sub(r"^  managed-content-sha256:.*\n", "", text, flags=re.MULTILINE)
    elif damage == "malformed":
        text = re.sub(r"sha256:[0-9a-f]{64}", "invalid", text, count=1)
    elif damage == "uppercase":
        text = re.sub(r"sha256:[0-9a-f]{64}", "sha256:" + "A" * 64, text, count=1)
    else:
        text = re.sub(r"sha256:[0-9a-f]{64}", "sha256:" + "0" * 64, text, count=1)
    path = install_text(text)
    assert skill.FORCE_COMMAND in synchronize()
    assert path.read_bytes() == text.encode("utf-8")
    status = skill.skill_status()
    assert status["auto_sync_eligible"] is False
    assert status["force_command"] == skill.FORCE_COMMAND
    assert (
        status["integrity"]
        == {
            "body": "altered",
            "missing": "missing",
            "malformed": "malformed",
            "uppercase": "malformed",
            "mismatch": "altered",
        }[damage]
    )
    with pytest.raises(UsageError, match="uvx office-export skill install --force"):
        skill.install_skill()
    assert skill.install_skill(force=True)["updated"]
    assert path.read_bytes() == skill.render_skill().encode("utf-8")


@pytest.mark.usefixtures("released")
@pytest.mark.parametrize("version", [office_export.__version__, "9.0"])
@pytest.mark.parametrize("altered", [False, True])
def test_equal_or_newer_auto_skill_is_not_rewritten(rendered, version: str, altered: bool) -> None:
    text = rendered(version) + ("Personal edit\n" if altered else "")
    path = install_text(text)
    before = path.stat().st_mtime_ns
    assert synchronize() == ""
    assert path.read_bytes() == text.encode("utf-8")
    assert path.stat().st_mtime_ns == before


@pytest.mark.parametrize("force", [False, True])
def test_explicit_install_never_downgrades_newer_skill(rendered, force: bool) -> None:
    path = install_text(rendered("9.0") + "Personal edit\n")
    before = path.read_bytes()
    assert skill.install_skill(force=force)["updated"] is False
    assert path.read_bytes() == before


def test_canonical_explicit_install_is_a_true_noop(monkeypatch: pytest.MonkeyPatch) -> None:
    skill.install_skill()
    monkeypatch.setattr(skill.os, "replace", lambda *args: pytest.fail("Canonical skill must not be rewritten"))
    assert skill.install_skill()["updated"] is False
    assert skill.install_skill(force=True)["updated"] is False


def test_equal_altered_explicit_install_requires_force() -> None:
    path = install_text(skill.render_skill() + "Personal edit\n")
    with pytest.raises(UsageError, match="--force"):
        skill.install_skill()
    assert skill.install_skill(force=True)["updated"] is True
    assert path.read_bytes() == skill.render_skill().encode("utf-8")


@pytest.mark.usefixtures("released")
@pytest.mark.parametrize(
    "version_line", ["", '  managed-version: "garbage"\n', "  managed-version: []\n", "  managed-version: 1.0\n"]
)
def test_invalid_or_missing_version_recovers_before_hash_check(version_line: str) -> None:
    path = install_text(f"---\nmetadata:\n  managed-by: office-export\n{version_line}---\nOld skill\n")
    assert "Updated managed skill" in synchronize()
    assert path.read_bytes() == skill.render_skill().encode("utf-8")


@pytest.mark.usefixtures("released")
def test_legacy_migrates_from_zero_and_is_still_removable() -> None:
    text = skill.MANAGED_MARKER + "\nOld instructions\n"
    path = install_text(text)
    status = skill.skill_status()
    assert status["installed_version"] == "0"
    assert status["integrity"] == "legacy"
    assert status["auto_sync_eligible"] is True
    assert "0 -> " + office_export.__version__ in synchronize()
    assert path.read_bytes() == skill.render_skill().encode("utf-8")
    install_text(text)
    assert skill.remove_skill()["removed"] is True


@pytest.mark.usefixtures("released")
@pytest.mark.parametrize(
    ("installed", "running", "expected"),
    [
        ("1.9", "1.10", True),
        ("1.0rc1", "1.0", True),
        ("1.0", "1.0.post1", True),
        ("1.0.dev1", "1.0a1", True),
        ("2!0.1", "1!9.0", False),
        ("1.0", "1.0.0", False),
        ("1.0+local", "1.0", False),
    ],
)
def test_pep440_comparison(
    rendered, monkeypatch: pytest.MonkeyPatch, installed: str, running: str, expected: bool
) -> None:
    path = install_text(rendered(installed))
    before = path.read_bytes()
    monkeypatch.setattr(office_export, "__version__", running)
    synchronize()
    assert (path.read_bytes() != before) is expected
    if expected:
        assert skill.inspect_skill_content(path.read_bytes()).installed_version == running


@pytest.mark.usefixtures("released")
def test_invalid_running_version_skips_silently(rendered, monkeypatch: pytest.MonkeyPatch) -> None:
    path = install_text(rendered("0.1.0"))
    before = path.read_bytes()
    monkeypatch.setattr(office_export, "__version__", "invalid")
    assert synchronize() == ""
    assert path.read_bytes() == before


@pytest.mark.parametrize(
    ("direct", "local"),
    [
        (None, False),
        ({"url": "file:///project", "dir_info": {}}, True),
        ({"url": "file:///project", "dir_info": {"editable": True}}, True),
        ({"url": "file:///project", "dir_info": {"editable": False}}, True),
        ({"url": "file:///package.tar.gz", "archive_info": {}}, True),
        ({"url": "file:///package.whl", "archive_info": {}}, False),
        ({"url": "https://example.test/package.whl", "archive_info": {}}, False),
        ({"url": "file:///project"}, True),
    ],
)
def test_distribution_provenance(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, direct: dict | None, local: bool
) -> None:
    monkeypatch.setattr(office_export, "__file__", str(tmp_path / "site-packages" / "office_export" / "__init__.py"))
    distribution = SimpleNamespace(
        read_text=lambda name: json.dumps(direct) if direct is not None else None,
        locate_file=lambda name: Path(office_export.__file__),
    )
    monkeypatch.setattr(skill.metadata, "distribution", lambda name: distribution)
    assert skill.is_local_development() is local


@pytest.mark.parametrize("provenance", ["missing", "malformed", "shadowed"])
def test_unknown_or_shadowed_runtime_is_excluded(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, provenance: str
) -> None:
    monkeypatch.setattr(office_export, "__file__", str(tmp_path / "site-packages" / "office_export" / "__init__.py"))

    def distribution(name: str):
        if provenance == "missing":
            raise skill.metadata.PackageNotFoundError(name)
        return SimpleNamespace(
            read_text=lambda name: "{" if provenance == "malformed" else None,
            locate_file=lambda name: tmp_path / name,
        )

    monkeypatch.setattr(skill.metadata, "distribution", distribution)
    assert skill.is_local_development() is True


@pytest.mark.parametrize("direct", ["[]", "null", '{"url": null}', '{"url": ""}', '{"url": "relative-path"}'])
def test_invalid_direct_url_metadata_is_excluded(monkeypatch: pytest.MonkeyPatch, tmp_path: Path, direct: str) -> None:
    monkeypatch.setattr(office_export, "__file__", str(tmp_path / "site-packages" / "office_export" / "__init__.py"))
    distribution = SimpleNamespace(read_text=lambda name: direct, locate_file=lambda name: Path(office_export.__file__))
    monkeypatch.setattr(skill.metadata, "distribution", lambda name: distribution)
    assert skill.is_local_development() is True


@pytest.mark.parametrize("layout", ["flat", "src"])
def test_legacy_editable_checkout_without_direct_url_is_excluded(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, layout: str
) -> None:
    (tmp_path / "pyproject.toml").write_text('[project]\nname = "office-export"\n', encoding="utf-8")
    source_root = tmp_path / "src" if layout == "src" else tmp_path
    runtime = source_root / "office_export" / "__init__.py"
    monkeypatch.setattr(office_export, "__file__", str(runtime))
    distribution = SimpleNamespace(read_text=lambda name: None, locate_file=lambda name: runtime)
    monkeypatch.setattr(skill.metadata, "distribution", lambda name: distribution)
    assert skill.is_local_development() is True


def test_local_build_skips_automatic_but_allows_explicit_install(rendered, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(skill, "is_local_development", lambda: True)
    path = install_text(rendered("0.1.0"))
    before = path.read_bytes()
    assert synchronize() == ""
    assert path.read_bytes() == before
    status = skill.skill_status()
    assert status["local_development"] is True
    assert status["auto_sync_skip_reason"] == "local_development"
    assert skill.install_skill()["updated"] is True


@pytest.mark.usefixtures("released")
def test_custom_directory_requires_explicit_update(rendered, tmp_path: Path) -> None:
    custom = tmp_path / "custom skills"
    path = install_text(rendered("0.1.0"), custom)
    before = path.read_bytes()
    assert synchronize() == ""
    assert path.read_bytes() == before
    assert not skill.default_skills_dir().exists()
    status = skill.skill_status(custom)
    assert status["standard_location"] is False
    assert status["auto_sync_skip_reason"] == "custom_directory"
    assert status["auto_sync_eligible"] is False
    assert skill.install_skill(custom)["updated"] is True
    path.write_text(skill.render_skill() + "Personal edit\n", encoding="utf-8")
    assert f'--skills-dir "{custom}"' in skill.skill_status(custom)["force_command"]


@pytest.mark.usefixtures("released")
def test_atomic_boundary_has_complete_closed_file_and_no_sidecars(rendered, monkeypatch: pytest.MonkeyPatch) -> None:
    path = install_text(rendered("0.1.0"))
    before = path.read_bytes()
    replace = skill.os.replace
    observed = []

    def check_replace(source: Path, destination: Path) -> None:
        assert source.parent == path.parent
        assert destination == path
        assert path.read_bytes() == before
        assert source.read_bytes() == skill.render_skill().encode("utf-8")
        # Rename also verifies that the temporary writer is closed on Windows.
        replace(source, destination)
        observed.append(destination.read_bytes())

    monkeypatch.setattr(skill.os, "replace", check_replace)
    assert "Updated" in synchronize()
    assert observed == [skill.render_skill().encode("utf-8")]
    assert list(path.parent.iterdir()) == [path]


@pytest.mark.usefixtures("released")
@pytest.mark.parametrize("failure", ["replace", "flush"])
def test_atomic_failure_preserves_original_and_cleans_temporary(
    rendered, monkeypatch: pytest.MonkeyPatch, failure: str
) -> None:
    path = install_text(rendered("0.1.0"))
    before = path.read_bytes()

    def fail(*args):
        raise PermissionError("Test failure")

    monkeypatch.setattr(skill.os, "replace" if failure == "replace" else "fsync", fail)
    assert "warning:" in synchronize()
    assert path.read_bytes() == before
    assert list(path.parent.iterdir()) == [path]


@pytest.mark.usefixtures("released")
@pytest.mark.parametrize("change", ["newer", "edited", "removed", "unmanaged"])
def test_revalidation_preserves_concurrent_changes(rendered, monkeypatch: pytest.MonkeyPatch, change: str) -> None:
    path = install_text(rendered("0.1.0"))
    replacement = (
        None
        if change == "removed"
        else (
            rendered("9.0")
            if change == "newer"
            else rendered("0.1.0") + "Edited\n"
            if change == "edited"
            else "Unmanaged\n"
        )
    )
    read = skill._read_content
    reads = 0

    def concurrent_read(target: Path):
        nonlocal reads
        reads += 1
        if reads == 2:
            if replacement is None:
                path.unlink()
            else:
                path.write_bytes(replacement.encode("utf-8"))
        return read(target)

    monkeypatch.setattr(skill, "_read_content", concurrent_read)
    assert synchronize() == ""
    assert (path.read_bytes() if path.exists() else None) == (replacement.encode("utf-8") if replacement else None)
    assert not list(path.parent.glob(".SKILL-*.tmp"))


@pytest.mark.parametrize("entry", ["target_file", "empty_dir", "unrelated_dir", "skill_dir"])
@pytest.mark.parametrize("force", [False, True])
def test_install_refuses_unexpected_paths(entry: str, force: bool) -> None:
    target = skill.skill_dir()
    target.parent.mkdir(parents=True)
    if entry == "target_file":
        target.write_text("keep", encoding="utf-8")
    else:
        target.mkdir()
        if entry == "unrelated_dir":
            (target / "notes.txt").write_text("keep", encoding="utf-8")
        elif entry == "skill_dir":
            (target / "SKILL.md").mkdir()
    with pytest.raises(UsageError):
        skill.install_skill(force=force)


def test_install_preserves_extra_files_and_removal_force_semantics() -> None:
    skill.install_skill()
    note = skill.skill_dir() / "notes.txt"
    note.write_text("keep", encoding="utf-8")
    skill.install_skill(force=True)
    assert note.read_text(encoding="utf-8") == "keep"
    with pytest.raises(UsageError, match="unmanaged entries"):
        skill.remove_skill()
    skill.remove_skill(force=True)
    assert not skill.skill_dir().exists()
    install_text("Unmanaged\n")
    with pytest.raises(UsageError, match="not marked as managed"):
        skill.remove_skill()
    assert skill.remove_skill(force=True)["removed"] is True


@pytest.mark.usefixtures("released")
@pytest.mark.parametrize(
    "text", ["---\nmetadata: [\n---\n", "---\nmetadata:\n  managed-by: office-export\n  managed-by: other\n---\n"]
)
def test_parse_failure_is_preserved_and_reported_without_command_failure(text: str) -> None:
    path = install_text(text)
    stdout, stderr = io.StringIO(), io.StringIO()
    assert cli.main(["formats", "--json"], stdout=stdout, stderr=stderr) == 0
    assert json.loads(stdout.getvalue())["ok"] is True
    assert "warning:" in stderr.getvalue()
    assert path.read_text(encoding="utf-8") == text


@pytest.mark.usefixtures("released")
@pytest.mark.parametrize("problem", ["update", "altered", "filesystem"])
def test_notices_preserve_json_and_primary_exit_status(rendered, monkeypatch: pytest.MonkeyPatch, problem: str) -> None:
    install_text(rendered("0.1.0") + ("Edit\n" if problem == "altered" else ""))
    if problem == "filesystem":

        def fail(*args):
            raise PermissionError("Test")

        monkeypatch.setattr(skill.os, "replace", fail)
    stdout, stderr = io.StringIO(), io.StringIO()
    assert cli.main(["formats", "--json"], stdout=stdout, stderr=stderr) == 0
    assert json.loads(stdout.getvalue())["ok"] is True
    assert stderr.getvalue().count("\n") == 1
    stdout, stderr = io.StringIO(), io.StringIO()
    assert cli.main(["unknown-option", "--json"], stdout=stdout, stderr=stderr) == 2
    assert json.loads(stdout.getvalue())["ok"] is False


@pytest.mark.parametrize(
    "args",
    [
        [],
        ["--help"],
        ["-h"],
        ["--version"],
        ["--about"],
        ["inspect", "--help"],
        ["doctor", "--help"],
        ["formats", "--json"],
        ["batch", "--help"],
        ["deck.pptx", "--to", "png", "--help"],
    ],
)
def test_all_normal_cli_entry_points_synchronize(monkeypatch: pytest.MonkeyPatch, args: list[str]) -> None:
    calls = []
    monkeypatch.setattr(cli, "synchronize_skill", lambda **kwargs: calls.append(kwargs["stderr"]))
    stdout, stderr = io.StringIO(), io.StringIO()
    assert cli.main(args, stdout=stdout, stderr=stderr) == 0
    assert calls == [stderr]


@pytest.mark.parametrize("action", ["install", "remove", "status", "--help", "install --help", "status --force"])
def test_skill_commands_never_synchronize(monkeypatch: pytest.MonkeyPatch, action: str) -> None:
    calls = []
    monkeypatch.setattr(cli, "synchronize_skill", lambda **kwargs: calls.append(True))
    cli.main(["skill", *action.split()], stdout=io.StringIO(), stderr=io.StringIO())
    assert calls == []


@pytest.mark.usefixtures("released")
def test_status_plain_and_json_are_read_only(rendered) -> None:
    path = install_text(rendered("0.1.0"))
    before = path.stat().st_mtime_ns
    stdout, stderr = io.StringIO(), io.StringIO()
    assert cli.main(["skill", "status", "--json"], stdout=stdout, stderr=stderr) == 0
    status = json.loads(stdout.getvalue())
    assert status["ok"] is True
    assert status["schema_version"] == "1.0"
    assert status["mode"] == "skill_status"
    assert status["installed"] is status["managed"] is status["auto_sync_eligible"] is True
    assert status["version_relation"] == "older"
    assert status["integrity"] == "valid"
    assert status["cli_version"] == office_export.__version__
    assert status["installed_version"] == "0.1.0"
    assert Path(status["path"]) == path
    stdout = io.StringIO()
    assert cli.main(["skill", "status"], stdout=stdout, stderr=stderr) == 0
    assert str(path) in stdout.getvalue()
    assert "older" in stdout.getvalue()
    assert "Automatic synchronization eligible: True" in stdout.getvalue()
    assert path.stat().st_mtime_ns == before
    assert stderr.getvalue() == ""


@pytest.mark.usefixtures("released")
@pytest.mark.parametrize("opening", ["--- \n", "\ufeff---\n", "\ufeff--- \n"])
def test_front_matter_delimiters_cannot_hide_conflicting_ownership(opening: str) -> None:
    text = opening + "metadata:\n  managed-by: other-tool\n---\n" + skill.MANAGED_MARKER
    path = install_text(text)
    assert synchronize() == ""
    assert path.read_bytes() == text.encode("utf-8")
    with pytest.raises(UsageError, match="unmanaged"):
        skill.install_skill(force=True)


def test_skill_help_documents_install_and_remove_force() -> None:
    stdout = io.StringIO()
    assert cli.main(["skill", "install", "--help"], stdout=stdout) == 0
    assert "Install --force replaces altered managed content" in stdout.getvalue()
    assert "Remove --force permits removal of unmanaged content" in stdout.getvalue()
    assert "skill status" in cli.build_root_help()
