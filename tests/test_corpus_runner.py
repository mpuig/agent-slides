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


def test_run_corpus_profile_semantics(tmp_path: Path) -> None:
    corpus = tmp_path / "corpus"
    corpus.mkdir(parents=True, exist_ok=True)
    _make_warning_deck(corpus / "warn.pptx")

    report_desktop = tmp_path / "desktop.json"
    cmd_desktop = [
        sys.executable,
        "scripts/run_corpus.py",
        str(corpus),
        "--profile",
        "desktop",
        "--out",
        str(report_desktop),
    ]
    desktop = subprocess.run(cmd_desktop, cwd=Path(__file__).resolve().parents[1], check=False)
    assert desktop.returncode == 0
    desktop_payload = json.loads(report_desktop.read_text(encoding="utf-8"))
    assert desktop_payload["profile"] == "desktop"
    assert desktop_payload["profile_failed"] == 0

    report_web = tmp_path / "web.json"
    cmd_web = [
        sys.executable,
        "scripts/run_corpus.py",
        str(corpus),
        "--profile",
        "web",
        "--out",
        str(report_web),
    ]
    web = subprocess.run(cmd_web, cwd=Path(__file__).resolve().parents[1], check=False)
    assert web.returncode == 1
    web_payload = json.loads(report_web.read_text(encoding="utf-8"))
    assert web_payload["profile"] == "web"
    assert web_payload["profile_failed"] == 1
