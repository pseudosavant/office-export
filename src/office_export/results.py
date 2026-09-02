from __future__ import annotations

import hashlib
import json
import platform
from datetime import UTC, datetime
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from typing import Any

from office_export import __version__

SCHEMA_VERSION = "1.0"


def utc_timestamp() -> str:
    return datetime.now(UTC).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_record(path: Path, **logical: Any) -> dict[str, Any]:
    record: dict[str, Any] = {
        "path": str(path.resolve()),
        "size": path.stat().st_size,
        "sha256": sha256_file(path),
    }
    record.update(logical)
    return record


def source_record(path: Path, source_format: str) -> dict[str, Any]:
    return {
        "path": str(path.resolve()),
        "format": source_format,
        "size": path.stat().st_size,
        "sha256": sha256_file(path),
    }


def dependency_version(distribution: str) -> str | None:
    try:
        return version(distribution)
    except PackageNotFoundError:
        return None


def base_result(*, mode: str, source: Path | None = None, source_format: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": True,
        "schema_version": SCHEMA_VERSION,
        "mode": mode,
        "tool": {"name": "office-export", "version": __version__},
        "timestamp": utc_timestamp(),
    }
    if source is not None and source_format is not None:
        payload["source"] = source_record(source, source_format)
    return payload


def runtime_versions() -> dict[str, str | None]:
    return {
        "python": platform.python_version(),
        "pdfium": dependency_version("pypdfium2"),
        "pillow": dependency_version("Pillow"),
        "pywin32": dependency_version("pywin32"),
    }


def write_manifest(path: Path, payload: dict[str, Any], *, force: bool) -> None:
    target = path.expanduser().resolve()
    if target.exists() and not force:
        from office_export.errors import UsageError

        raise UsageError(f"Manifest already exists: {target}. Use --force to replace it.", code="manifest_exists")
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(f".{target.name}.office-export.tmp")
    temporary.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")
    temporary.replace(target)
