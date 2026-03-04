from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

from slides_cli import Presentation


def test_render_office_hook_with_copy_command(tmp_path: Path) -> None:
    corpus = tmp_path / "corpus"
    outdir = tmp_path / "rendered"
    corpus.mkdir(parents=True, exist_ok=True)
    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    src = corpus / "deck.pptx"
    pres.save(src)

    report = tmp_path / "render.json"
    cmd = [
        sys.executable,
        "scripts/render_office.py",
        str(corpus),
        str(outdir),
        "--command-template",
        'cp "{pptx}" "{outdir}/{name}"',
        "--report",
        str(report),
    ]
    result = subprocess.run(cmd, cwd=Path(__file__).resolve().parents[1], check=False)
    assert result.returncode == 0
    payload = json.loads(report.read_text(encoding="utf-8"))
    assert payload["failed"] == 0
    assert (outdir / "deck.pptx").exists()
