from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

from slides_cli import Presentation


def _make_warning_deck(path: Path) -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_text(
        slide_index=slide.index,
        text="Hello {{name}}",
        left=0.5,
        top=0.5,
        width=4.0,
        height=1.0,
    )
    pres.save(path)


def test_interop_matrix_profiles(tmp_path: Path) -> None:
    corpus = tmp_path / "corpus"
    corpus.mkdir(parents=True, exist_ok=True)
    _make_warning_deck(corpus / "warn.pptx")

    out = tmp_path / "interop.json"
    cmd = [
        sys.executable,
        "scripts/interop_matrix.py",
        str(corpus),
        "--profiles",
        "desktop",
        "strict",
        "--out",
        str(out),
    ]
    result = subprocess.run(cmd, cwd=Path(__file__).resolve().parents[1], check=False)
    assert result.returncode == 1
    payload = json.loads(out.read_text(encoding="utf-8"))
    assert payload["profiles"]["desktop"]["ok"] == 1
    assert payload["profiles"]["strict"]["failed"] == 1
