from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


def test_benchmark_script_writes_json(tmp_path: Path) -> None:
    out = tmp_path / "bench.json"
    cmd = [
        sys.executable,
        "scripts/benchmark.py",
        "--runs",
        "1",
        "--slides",
        "1",
        "--out",
        str(out),
    ]
    result = subprocess.run(cmd, cwd=Path(__file__).resolve().parents[1], check=False)
    assert result.returncode == 0
    payload = json.loads(out.read_text(encoding="utf-8"))
    assert payload["runs"] == 1
    assert payload["slides"] == 1
    assert isinstance(payload["mean_sec"], float)
