from __future__ import annotations

from pathlib import Path

import pytest

from office_export.errors import UsageError
from office_export.publishing import publish_file, publish_images


def test_publish_file_never_replaces_a_directory(tmp_path: Path) -> None:
    staging = tmp_path / "staging"
    staging.mkdir()
    source = staging / "output.pdf"
    source.write_bytes(b"pdf")
    destination = tmp_path / "existing.pdf"
    destination.mkdir()

    with pytest.raises(UsageError) as raised:
        publish_file(source, destination, force=True)

    assert raised.value.context.code == "output_not_file"
    assert source.read_bytes() == b"pdf"
    assert destination.is_dir()


def test_existing_image_suffix_directory_is_a_directory_target(tmp_path: Path) -> None:
    staging = tmp_path / "staging"
    staging.mkdir()
    source = staging / "page-001.png"
    source.write_bytes(b"png")
    destination = tmp_path / "images.png"
    destination.mkdir()

    published = publish_images([source], destination, output_format="png", force=False)

    assert published == [destination / "page-001.png"]
    assert published[0].read_bytes() == b"png"


def test_image_collision_preserves_existing_and_unrelated_files(tmp_path: Path) -> None:
    staging = tmp_path / "staging"
    staging.mkdir()
    source = staging / "page-001.png"
    source.write_bytes(b"new")
    destination = tmp_path / "images"
    destination.mkdir()
    collision = destination / source.name
    collision.write_bytes(b"old")
    unrelated = destination / "notes.txt"
    unrelated.write_text("keep", encoding="utf-8")

    with pytest.raises(UsageError) as raised:
        publish_images([source], destination, output_format="png", force=False)

    assert raised.value.context.code == "output_exists"
    assert collision.read_bytes() == b"old"
    assert unrelated.read_text(encoding="utf-8") == "keep"
    assert source.read_bytes() == b"new"
