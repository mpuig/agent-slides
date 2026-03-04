from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


def test_visual_diff_detects_mismatch(tmp_path: Path) -> None:
    baseline = tmp_path / "baseline"
    current = tmp_path / "current"
    baseline.mkdir()
    current.mkdir()
    (baseline / "slide1.pdf").write_bytes(b"A")
    (current / "slide1.pdf").write_bytes(b"B")

    report = tmp_path / "report.json"
    cmd = [
        sys.executable,
        "scripts/visual_diff.py",
        str(baseline),
        str(current),
        "--report",
        str(report),
    ]
    result = subprocess.run(cmd, cwd=Path(__file__).resolve().parents[1], check=False)
    assert result.returncode == 1
    payload = json.loads(report.read_text(encoding="utf-8"))
    assert payload["failed"] == 1


def test_visual_diff_passes_on_equal_files(tmp_path: Path) -> None:
    baseline = tmp_path / "baseline"
    current = tmp_path / "current"
    baseline.mkdir()
    current.mkdir()
    (baseline / "slide1.pdf").write_bytes(b"SAME")
    (current / "slide1.pdf").write_bytes(b"SAME")

    cmd = [
        sys.executable,
        "scripts/visual_diff.py",
        str(baseline),
        str(current),
    ]
    result = subprocess.run(cmd, cwd=Path(__file__).resolve().parents[1], check=False)
    assert result.returncode == 0
