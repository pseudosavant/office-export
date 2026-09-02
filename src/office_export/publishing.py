from __future__ import annotations

import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

from office_export.errors import ConversionError, UsageError


@contextmanager
def staging_directory(parent: Path) -> Iterator[Path]:
    parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".office-export-", dir=parent) as value:
        yield Path(value)


def publish_file(source: Path, destination: Path, *, force: bool) -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists() and not destination.is_file():
        raise UsageError(f"File output path is not a file: {destination}", code="output_not_file")
    if destination.exists() and not force:
        raise UsageError(f"Output already exists: {destination}. Use --force to replace it.", code="output_exists")
    backup = source.parent / f".{destination.name}.backup"
    had_existing = destination.exists()
    try:
        if had_existing:
            destination.replace(backup)
        source.replace(destination)
        backup.unlink(missing_ok=True)
    except OSError as exc:
        if backup.exists():
            if destination.exists():
                destination.unlink(missing_ok=True)
            backup.replace(destination)
        raise ConversionError(
            "output_write_failed",
            f"Could not publish output to '{destination}': {exc}",
        ) from exc
    return destination


def publish_images(
    sources: list[Path],
    destination: Path,
    *,
    output_format: str,
    force: bool,
) -> list[Path]:
    if not sources:
        raise ConversionError("no_images_created", "The conversion did not create any images.")
    file_suffixes = {"png": {".png"}, "jpeg": {".jpg", ".jpeg"}}[output_format]
    explicit_file = destination.suffix.lower() in file_suffixes and not destination.is_dir()
    if explicit_file:
        if len(sources) != 1:
            raise UsageError(
                f"A single image output path was provided, but the selection creates {len(sources)} images. "
                "Use an output directory instead."
            )
        return [publish_file(sources[0], destination, force=force)]

    if destination.exists() and not destination.is_dir():
        raise UsageError(f"Image output path is not a directory: {destination}", code="output_not_directory")
    collisions = [destination / source.name for source in sources if (destination / source.name).exists()]
    if collisions and not force:
        preview = ", ".join(str(path) for path in collisions[:3])
        if len(collisions) > 3:
            preview += f", and {len(collisions) - 3} more"
        raise UsageError(f"Image output already exists: {preview}. Use --force to replace it.", code="output_exists")

    destination_existed = destination.exists()
    backup_dir = sources[0].parent / ".publish-backups"
    moved: list[Path] = []
    backups: list[tuple[Path, Path]] = []
    try:
        destination.mkdir(parents=True, exist_ok=True)
        backup_dir.mkdir(exist_ok=True)
        for source in sources:
            target = destination / source.name
            if target.exists():
                backup = backup_dir / target.name
                target.replace(backup)
                backups.append((backup, target))
            source.replace(target)
            moved.append(target)
        for backup, _ in backups:
            backup.unlink(missing_ok=True)
    except OSError as exc:
        for path in reversed(moved):
            path.unlink(missing_ok=True)
        for backup, target in reversed(backups):
            if backup.exists():
                backup.replace(target)
        if not destination_existed:
            try:
                destination.rmdir()
            except OSError:
                pass
        raise ConversionError(
            "output_write_failed",
            f"Could not publish images to '{destination}': {exc}",
        ) from exc
    return moved
