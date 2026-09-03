"""Run with a clean wheel environment's Python, without pytest or source imports."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from importlib import metadata
from pathlib import Path
from unittest.mock import patch

import office_export
from office_export import skill


def main() -> None:
    version = office_export.__version__
    assert version == metadata.version("office-export")
    assert not skill.is_local_development()
    assert skill.inspect_skill_content(skill.render_skill().encode("utf-8")).integrity == "valid"
    with tempfile.TemporaryDirectory(prefix="office-export-wheel-") as temporary:
        root = Path(temporary)
        env = {**os.environ, "HOME": str(root), "USERPROFILE": str(root)}
        env.pop("PYTHONPATH", None)
        executable = Path(sys.executable).parent / ("office-export.exe" if os.name == "nt" else "office-export")

        def run(*args: str) -> subprocess.CompletedProcess[str]:
            return subprocess.run(
                [str(executable), *args], cwd=root, env=env, capture_output=True, text=True, check=True
            )

        result = run("--version")
        assert result.stdout.strip() == f"office-export {version}"
        assert result.stderr == ""
        path = root / ".agents" / "skills" / "office-export" / "SKILL.md"
        assert not path.exists()
        installed = json.loads(run("skill", "install", "--json").stdout)
        assert Path(installed["path"]) == path
        assert skill.inspect_skill_content(path.read_bytes()).installed_version == version
        with patch.object(office_export, "__version__", "0.dev0"):
            older = skill.render_skill()
        path.write_bytes(older.encode("utf-8"))
        status = json.loads(run("skill", "status", "--json").stdout)
        assert status["auto_sync_eligible"] is True
        assert status["local_development"] is False
        assert path.read_text(encoding="utf-8") == older
        result = run("formats", "--json")
        assert json.loads(result.stdout)["ok"] is True
        assert f"0.dev0 -> {version}" in result.stderr
        assert path.read_bytes() == skill.render_skill().encode("utf-8")
        assert list(path.parent.iterdir()) == [path]
        assert json.loads(run("skill", "remove", "--json").stdout)["removed"] is True
    print("Wheel version, canonical skill, CLI entry point, and automatic synchronization passed.")


if __name__ == "__main__":
    main()
