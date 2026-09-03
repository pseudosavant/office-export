from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import tempfile
from dataclasses import dataclass
from importlib import metadata
from pathlib import Path
from typing import Any, TextIO
from urllib.parse import urlsplit

import yaml
from packaging.version import InvalidVersion, Version
from yaml.nodes import MappingNode, Node, ScalarNode

import office_export
from office_export.errors import UsageError

SKILL_NAME = "office-export"
MANAGED_BY = "office-export"
MANAGED_MARKER = "<!-- managed-by: office-export -->"
FORCE_COMMAND = "uvx office-export skill install --force"

_SKILL_TEMPLATE = """---
name: office-export
description: Export Word, Excel, PowerPoint, and PDF files to faithful PDF, PNG, or JPEG output with `uvx office-export`. Use when an agent needs to inspect or render Office documents, select pages, slides, sheets, ranges, or charts, or create local previews through desktop Microsoft Office.
metadata:
  managed-by: office-export
  managed-version: {version}
  managed-content-sha256: ""
---

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


def _normalize(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _digest(text: str) -> str:
    return "sha256:" + hashlib.sha256(_normalize(text).encode("utf-8")).hexdigest()


def render_skill() -> str:
    """Render the single bundled template with the exact CLI runtime version."""
    text = _normalize(_SKILL_TEMPLATE.format(version=json.dumps(office_export.__version__)))
    return text.replace('managed-content-sha256: ""', f'managed-content-sha256: "{_digest(text)}"', 1)


# Kept for callers that inspect the bundled instructions. Writes use render_skill().
SKILL_MD = render_skill()


def default_skills_dir() -> Path:
    return Path.home() / ".agents" / "skills"


def skill_dir(skills_dir: Path | None = None) -> Path:
    return (skills_dir or default_skills_dir()) / SKILL_NAME


def _mapping(node: Node | None) -> dict[str, Node]:
    if not isinstance(node, MappingNode):
        raise UsageError("Skill front matter and metadata must be YAML mappings.")
    result = {}
    for key, value in node.value:
        name = _string(key)
        if name is None or name in result:
            raise UsageError("Skill metadata contains ambiguous or duplicate YAML keys.")
        result[name] = value
    return result


def _string(node: Node | None) -> str | None:
    if isinstance(node, ScalarNode) and node.tag == "tag:yaml.org,2002:str":
        return node.value
    return None


def _front_matter(text: str) -> tuple[dict[str, Node], int]:
    opening = re.match(r"\A\ufeff?---[ \t]*\n", text)
    if opening is None:
        return {}, 0
    offset = opening.end()
    end = re.search(r"^---[ \t]*(?:\n|$)", text[offset:], re.MULTILINE)
    if end is None:
        raise UsageError("Skill front matter is missing its closing delimiter.")
    try:
        fields = _mapping(yaml.compose(text[offset : offset + end.start()], Loader=yaml.SafeLoader))
        return (_mapping(fields["metadata"]) if "metadata" in fields else {}), offset
    except yaml.YAMLError as exc:
        raise UsageError("Cannot parse skill YAML front matter.") from exc


def _version(value: str | None) -> Version | None:
    if value is None:
        return None
    try:
        return Version(value)
    except InvalidVersion:
        return None


@dataclass(frozen=True)
class SkillState:
    content: bytes | None
    managed: bool = False
    installed_version: str | None = None
    version: Version | None = None
    integrity: str = "not_applicable"
    legacy: bool = False

    def relation(self, current: Version | None) -> str | None:
        if self.version is None or current is None:
            return None
        return "older" if self.version < current else "newer" if self.version > current else "equal"

    def can_update(self, current: Version | None) -> bool:
        return bool(
            self.managed
            and current is not None
            and (self.version is None or (self.version < current and (self.legacy or self.integrity == "valid")))
        )


def _check_paths(skill_path: Path) -> None:
    for path in (skill_path.parent, skill_path):
        if path.is_symlink() or getattr(path, "is_junction", lambda: False)():
            raise UsageError(f"Refusing to manage linked skill path '{path}'.")
    if skill_path.parent.exists() and not skill_path.parent.is_dir():
        raise UsageError(f"Skill target is not a directory: {skill_path.parent}")
    if skill_path.exists() and not skill_path.is_file():
        raise UsageError(f"Skill file is not a regular file: {skill_path}")


def _read_content(skill_path: Path) -> bytes | None:
    _check_paths(skill_path)
    try:
        return skill_path.read_bytes()
    except FileNotFoundError:
        return None


def inspect_skill_content(content: bytes | None) -> SkillState:
    if content is None:
        return SkillState(None)
    text = _normalize(content.decode("utf-8"))
    fields, offset = _front_matter(text)
    owner = fields.get("managed-by")
    legacy_marker = MANAGED_MARKER in text
    managed = _string(owner) == MANAGED_BY if owner is not None else legacy_marker
    if not managed:
        return SkillState(content)
    version_text = _string(fields.get("managed-version"))
    version = _version(version_text)
    legacy = legacy_marker and "managed-version" not in fields
    if legacy:
        return SkillState(content, True, "0", Version("0"), "legacy", True)
    hash_node = fields.get("managed-content-sha256")
    stored_hash = _string(hash_node)
    if hash_node is None:
        integrity = "missing"
    elif stored_hash is None or re.fullmatch(r"sha256:[0-9a-f]{64}", stored_hash) is None:
        integrity = "malformed"
    else:
        # Node marks locate just the scalar, including its quotes, in the original
        # front matter. Never reserialize YAML when checking an installed file.
        start, end = offset + hash_node.start_mark.index, offset + hash_node.end_mark.index
        lexeme = text[start:end]
        if lexeme not in (stored_hash, f'"{stored_hash}"', f"'{stored_hash}'"):
            integrity = "malformed"
        else:
            calculated = _digest(text[:start] + '""' + text[end:])
            integrity = "valid" if calculated == stored_hash else "altered"
    return SkillState(content, True, version_text if version is not None else None, version, integrity)


def is_local_development() -> bool:
    """Use distribution provenance, never the launcher or the current directory."""
    try:
        runtime_init = Path(office_export.__file__).resolve()
        # Older editable installs may have egg-info but no PEP 610 record.
        # Inspect the flat-layout and src-layout project roots in that case too.
        for root in (runtime_init.parents[1], runtime_init.parents[2]):
            if any((root / marker).exists() for marker in ("pyproject.toml", "setup.py", "setup.cfg", ".git")):
                return True
        distribution = metadata.distribution(MANAGED_BY)
        direct_text = distribution.read_text("direct_url.json")
        if direct_text is not None:
            direct = json.loads(direct_text)
            if not isinstance(direct, dict) or not isinstance(direct.get("url"), str):
                return True
            if "dir_info" in direct:
                return True
            source = urlsplit(direct["url"])
            if not source.scheme:
                return True
            # A built wheel installed from disk is still a normal installation.
            # Local source directories and local source archives are development.
            if source.scheme == "file" and not (source.path.lower().endswith(".whl") and "archive_info" in direct):
                return True
        # Also catch a checkout imported ahead of an unrelated installed copy.
        installed_init = Path(distribution.locate_file("office_export/__init__.py")).resolve()
        return installed_init != runtime_init
    except (metadata.PackageNotFoundError, OSError, ValueError, KeyError, TypeError):
        # Unknown provenance is not sufficient authorization to maintain a skill.
        return True


def _atomic_write(skill_path: Path, text: str, *, previous: bytes | None) -> bool:
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            dir=skill_path.parent,
            prefix=".SKILL-",
            suffix=".tmp",
            delete=False,
        ) as stream:
            temporary = Path(stream.name)
            stream.write(text)
            stream.flush()
            os.fsync(stream.fileno())
        # A concurrent update, edit, or removal invalidates this snapshot. In
        # particular, never deliberately overwrite a newly observed newer skill.
        if _read_content(skill_path) != previous:
            return False
        os.replace(temporary, skill_path)
        return True
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def _force_recommendation(skills_dir: Path | None = None) -> str:
    if skills_dir is None:
        return FORCE_COMMAND
    return f'{FORCE_COMMAND} --skills-dir "{skills_dir}"'


def install_skill(skills_dir: Path | None = None, *, force: bool = False) -> dict[str, Any]:
    target = skill_dir(skills_dir)
    skill_path = target / "SKILL.md"
    state = inspect_skill_content(_read_content(skill_path))
    existed = state.content is not None
    if target.exists() and not existed:
        raise UsageError(f"Refusing to install into '{target}' because it contains no managed SKILL.md.")
    if existed and not state.managed:
        raise UsageError(f"Refusing to overwrite unmanaged skill file '{skill_path}', even with --force.")
    current = _version(office_export.__version__)
    if current is None:
        raise UsageError("The running CLI version is invalid. Cannot install its managed skill.")
    text = render_skill()
    write = True
    if state.relation(current) == "newer":
        write = False
    elif existed and state.version is not None and not state.legacy:
        if state.integrity != "valid" and not force:
            raise UsageError(
                f"Managed skill '{skill_path}' is altered or unverifiable. Use {_force_recommendation(skills_dir)}."
            )
        if state.content == text.encode("utf-8") or (state.version == current and not force):
            write = False
    if write:
        target.mkdir(parents=True, exist_ok=True)
        if not _atomic_write(skill_path, text, previous=state.content):
            raise UsageError(f"Skill '{skill_path}' changed during installation. Run the command again.")
    return {
        "installed": True,
        "created": not existed,
        "updated": existed and write,
        "skill": SKILL_NAME,
        "path": str(skill_path.resolve()),
    }


def skill_status(skills_dir: Path | None = None) -> dict[str, Any]:
    path = skill_dir(skills_dir) / "SKILL.md"
    state = inspect_skill_content(_read_content(path))
    current = _version(office_export.__version__)
    local = is_local_development()
    standard = path.absolute() == (skill_dir() / "SKILL.md").absolute()
    recommendation = None
    if state.managed and state.version is not None and not state.legacy and state.integrity != "valid":
        if state.relation(current) != "newer":
            recommendation = _force_recommendation(skills_dir)
    relation = state.relation(current)
    if local:
        reason = "local_development"
    elif not standard:
        reason = "custom_directory"
    elif current is None:
        reason = "invalid_cli_version"
    elif state.content is None:
        reason = "not_installed"
    elif not state.managed:
        reason = "unmanaged"
    elif relation in {"equal", "newer"}:
        reason = relation
    elif not state.can_update(current):
        reason = "altered_or_unverifiable"
    else:
        reason = None
    return {
        "skill": SKILL_NAME,
        "path": str(path.absolute()),
        "standard_location": standard,
        "installed": state.content is not None,
        "managed": state.managed,
        "cli_version": office_export.__version__,
        "installed_version": state.installed_version,
        "version_relation": relation,
        "integrity": state.integrity,
        "auto_sync_eligible": reason is None,
        "local_development": local,
        "auto_sync_skip_reason": reason,
        "force_command": recommendation,
    }


def synchronize_skill(*, stderr: TextIO) -> None:
    """Best-effort local maintenance. Never alter the primary command's result."""
    try:
        current = _version(office_export.__version__)
        if current is None or is_local_development():
            return
        path = skill_dir() / "SKILL.md"
        state = inspect_skill_content(_read_content(path))
        if not state.managed or state.relation(current) in {"equal", "newer"}:
            return
        if not state.can_update(current):
            stderr.write(f"Managed skill at {path} is altered or unverifiable. Use {FORCE_COMMAND}.\n")
            return
        if _atomic_write(path, render_skill(), previous=state.content):
            old = state.installed_version or "unknown"
            stderr.write(f"Updated managed skill {old} -> {office_export.__version__} at {path}.\n")
    except Exception:
        # Maintenance, including its notices, must not fail the user's operation.
        # Keep optional warnings generic so malformed file contents stay out of logs.
        try:
            stderr.write("warning: Could not synchronize the managed office-export skill.\n")
        except Exception:
            pass


def remove_skill(skills_dir: Path | None = None, *, force: bool = False) -> dict[str, Any]:
    target = skill_dir(skills_dir)
    skill_path = target / "SKILL.md"
    _check_paths(skill_path)
    if not target.exists():
        return {"removed": False, "skill": SKILL_NAME, "path": str(target), "reason": "not_installed"}
    if not target.is_dir():
        raise UsageError(f"Skill target is not a directory: {target}")
    if not skill_path.exists() and not force:
        raise UsageError(f"Refusing to remove '{target}' because SKILL.md is missing.")
    if skill_path.exists():
        if not force and not inspect_skill_content(skill_path.read_bytes()).managed:
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
