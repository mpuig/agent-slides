# agent-slides

`agent-slides` is an agent skill for generating professional, brand-compliant PowerPoint (`.pptx`) presentations. It ships 7 composable skills that work across agent harnesses, powered by a purpose-built CLI and Python library designed for agent DX.

**Skills** encode workflow knowledge — when to dry-run, how to chain stages, how to recover from errors. **The CLI** provides the execution layer — declarative JSON payloads, runtime schema discovery, Pydantic validation, and transactional rollback. **The library** wraps `python-pptx` into a deterministic, agent-safe API.

## Goals

- Agent skill for AI-driven deck generation, usable across harnesses
- Agent-ready CLI with JSON payloads, schema introspection, and context-window discipline
- Declarative operation engine with dry-run + transactional rollback
- Deterministic save path for reproducible automation
- Input hardening: control-char rejection, path-traversal checks, validation-before-mutation
- Design-contract enforcement via design profiles and lint engine
- Python 3.12+ and modern packaging (`uv`)

## Install

```bash
uv sync --all-groups
```

## Quick Example

```bash
# apply operations to create a deck from scratch
slides apply --ops-json @ops.json --output out.pptx
slides apply --ops-json @ops.json --icon-pack-dir /path/to/private/icons --output out.pptx

# open existing deck, apply ops, and save
slides apply in.pptx --ops-json @ops.json --output out.pptx

# repair a deck
slides repair in.pptx --output out.pptx

# inspect and validate
slides inspect in.pptx --summary
slides validate in.pptx
slides validate in.pptx --deep
slides validate in.pptx --deep --xsd-dir ./ooxml-xsd --require-xsd
slides validate in.pptx --fail-on-warning
slides inspect in.pptx --placeholders 0

# discovery and reproducibility
slides docs json
slides docs markdown
slides inspect in.pptx --fingerprint

# inspect and semantic search (read-only)
slides inspect in.pptx --out slides-index.json
slides find in.pptx --query "pricing margin" --out find.json

# prompt-like edits from CLI
slides edit in.pptx --query "legacy plan" --replacement "target-state plan" --output edited.pptx
slides inspect edited.pptx --fields slides.slide_uid,slides.shapes.shape_uid --out ids.json
slides edit edited.pptx --query "target-state" --replacement "approved target-state" --slide-uid "<slide_uid>" --shape-uid "<shape_uid>" --output edited2.pptx
slides transform edited.pptx --slide 7 --to timeline --output styled.pptx

# contract version
slides version

# runtime discovery (methods + workflows + schemas)
slides docs json
slides docs method:render
slides docs method:inspect:markdown
slides docs schema:slides-document
slides docs schema:template-layout:markdown

# extract template into agent contracts
slides extract "examples/Sample Slides_16x9.pptx" --output-dir extracted

# context-window controls
slides inspect in.pptx --out slides.ndjson --ndjson
slides inspect in.pptx --summary --fields slide_count,shape_count
slides plan-inspect --slides-json @slides.json --content-only --page-size 5 --out plan-page1.json

# reduce output verbosity (agent-friendly)
slides inspect in.pptx --compact                                # strip nulls + single-line JSON
slides qa in.pptx --out qa.json --compact                       # default is quiet when writing outputs
slides lint in.pptx --profile dp.json --out lint.json --compact
slides apply --ops-json @ops.json --output out.pptx --verbose   # opt-in full stdout payloads

# paginated list access for agents
slides inspect in.pptx --out page1.json                          # default page-size=25
slides inspect in.pptx --page-size 20 --page-token 0 --out page1.json
slides find in.pptx --query "pricing" --limit 200 --page-size 25 --page-all --ndjson --out find-pages.ndjson

# design contract enforcement (runtime)
slides qa out.pptx --slides-json @slides.json --profile design-profile.json --out qa.json
```

## Agent Workflow (Render -> Lint)

```bash
# 1) Render deck from slides document
slides render --slides-json @slides.json --profile design-profile.json --output out.pptx

# 2) Validate + lint design rules
slides validate out.pptx
slides lint out.pptx --profile design-profile.json --out lint.json
```

`slides.json` contains `plan` and `ops`, where `ops` uses discriminated operation objects:

```json
{
  "operations": [
    {"op": "add_slide", "layout_index": 6},
    {
      "op": "add_text",
      "slide_index": 0,
      "text": "Hello",
      "left": 0.5,
      "top": 0.5,
      "width": 4,
      "height": 1
    },
    {
      "op": "add_table",
      "slide_index": 0,
      "rows": [["KPI", "Value"], ["Revenue", "$10M"]],
      "left": 0.5,
      "top": 1.3,
      "width": 5.5,
      "height": 1.5
    },
    {"op": "set_semantic_text", "slide_index": 0, "role": "body", "text": "Key takeaway"},
    {"op": "set_title_subtitle", "slide_index": 0, "title": "Board Update"},
    {"op": "set_slide_background", "slide_index": 0, "color_hex": "F4F7F9"},
    {"op": "set_chart_title", "slide_index": 0, "chart_index": 0, "text": "Revenue Trend"},
    {"op": "set_chart_axis_titles", "slide_index": 0, "chart_index": 0, "category_title": "Quarter", "value_title": "M EUR"},
    {"op": "set_chart_data_labels", "slide_index": 0, "chart_index": 0, "enabled": true, "show_value": true},
    {"op": "set_chart_axis_scale", "slide_index": 0, "chart_index": 0, "minimum": 0, "maximum": 20, "major_unit": 5},
    {"op": "set_chart_plot_style", "slide_index": 0, "chart_index": 0, "gap_width": 120, "overlap": 0},
    {"op": "set_chart_series_style", "slide_index": 0, "chart_index": 0, "series_index": 0, "fill_color_hex": "0A4280"},
    {"op": "set_line_series_marker", "slide_index": 0, "chart_index": 0, "series_index": 0, "style": "diamond", "size": 8},
    {"op": "set_chart_series_trendline", "slide_index": 0, "chart_index": 0, "series_index": 0, "trend_type": "linear"},
    {"op": "set_chart_series_error_bars", "slide_index": 0, "chart_index": 0, "series_index": 0, "value": 0.2},
    {"op": "set_chart_legend", "slide_index": 0, "chart_index": 0, "position": "bottom"},
    {"op": "add_area_chart", "slide_index": 0, "categories": ["Q1", "Q2"], "series": [["Area", [10,12]]], "style": "percent_stacked", "left": 0.5, "top": 4.0, "width": 4, "height": 2},
    {"op": "add_doughnut_chart", "slide_index": 0, "categories": ["A", "B"], "series": [["Mix", [60,40]]], "style": "exploded", "left": 5.0, "top": 4.0, "width": 4, "height": 2},
    {"op": "add_scatter_chart", "slide_index": 0, "series": [["S1", [[1,2],[2,3]]]], "style": "smooth_no_markers", "left": 0.5, "top": 6.2, "width": 4, "height": 2},
    {"op": "add_radar_chart", "slide_index": 0, "categories": ["N", "S", "E", "W"], "series": [["Coverage", [80,65,72,90]]], "style": "markers", "left": 5.0, "top": 6.2, "width": 4, "height": 2},
    {"op": "add_bubble_chart", "slide_index": 0, "series": [["Portfolio", [[1,2,3],[2,3,5]]]], "style": "bubble_3d", "left": 0.5, "top": 8.4, "width": 4, "height": 2},
    {"op": "add_combo_chart_overlay", "slide_index": 0, "categories": ["Q1", "Q2"], "bar_series": [["Rev", [1,2]]], "line_series": [["Margin", [30,35]]], "left": 0.5, "top": 0.5, "width": 8, "height": 3},
    {"op": "set_chart_secondary_axis", "slide_index": 0, "chart_index": 1, "enable": true},
    {"op": "set_chart_combo_secondary_mapping", "slide_index": 0, "chart_index": 0, "series_indices": [1]},
    {"op": "set_image_crop", "slide_index": 0, "image_index": 0, "crop_left": 0.1, "crop_right": 0.1},
    {"op": "add_media", "slide_index": 0, "path": "demo.mp4", "left": 1, "top": 3, "width": 3, "height": 2, "mime_type": "video/mp4"}
  ]
}
```

## Agent Skills

agent-slides includes 7 composable skills for end-to-end deck workflows. Skills are self-contained Markdown files that encode workflow knowledge agents can't learn from `--help`. They work across any harness that supports skill files (Claude Code, etc.).

Install by symlinking into your harness's skill directory, e.g.: `ln -s skills .claude/skills`

| Skill | Command | Phase | Description |
|---|---|---|---|
| [slides-extract](skills/slides-extract/SKILL.md) | `/slides-extract` | Setup | Extract template contracts from a `.pptx` file |
| [slides-build](skills/slides-build/SKILL.md) | `/slides-build` | Create | Build a complete deck from a brief |
| [slides-edit](skills/slides-edit/SKILL.md) | `/slides-edit` | Modify | Text edits, transforms, ops patches |
| [slides-audit](skills/slides-audit/SKILL.md) | `/slides-audit` | Post-build | Technical lint: fonts, overlap, contrast |
| [slides-critique](skills/slides-critique/SKILL.md) | `/slides-critique` | Post-build | Storytelling: action titles, MECE, hierarchy |
| [slides-polish](skills/slides-polish/SKILL.md) | `/slides-polish` | Post-build | Final pass: notes, metadata, sources |
| [slides-full](skills/slides-full/SKILL.md) | `/slides-full` | End-to-end | Full pipeline: extract → build → audit → critique → polish |

**Typical workflow:**

```
/slides-extract  →  /slides-build  →  /slides-audit  →  /slides-critique  →  /slides-polish
                                        ↕               ↕
                                    /slides-edit  ←  (targeted fixes)

Or use /slides-full to run the entire pipeline in one command.
```

Each skill is self-contained with its own reference files. See `skills/*/SKILL.md` for details.

## Development

```bash
uv run ruff check .
uv run ruff format .
uv run ty check
uv run pytest
uv run python scripts/benchmark.py --runs 5 --slides 25
```

## Examples

```bash
# from project root
slides apply --ops-json @examples/ops.json --output examples/out.pptx
uv run python scripts/benchmark.py
uv run python scripts/run_corpus.py ./corpus --out corpus-report.json --resave-dir ./corpus-resaved
uv run python scripts/run_corpus.py ./corpus --baseline baseline-report.json --out current-report.json
uv run python scripts/run_corpus.py ./corpus --profile web --out web-report.json
uv run python scripts/run_corpus.py ./corpus --deep-validate --xsd-dir ./ooxml-xsd --require-xsd --out corpus-report.json
uv run python scripts/interop_matrix.py ./corpus --out interop-report.json
uv run python scripts/interop_matrix.py ./corpus --xsd-dir ./ooxml-xsd --require-xsd --out interop-report.json
uv run python scripts/roundtrip_compare.py ./corpus/MyDeck.pptx
uv run python scripts/rebuild_compare.py ./corpus/MyDeck.pptx --rebuild-dir rebuilt/
uv run python scripts/render_office.py ./corpus ./render-current --command-template 'soffice --headless --convert-to pdf --outdir "{outdir}" "{pptx}"'
uv run python scripts/visual_diff.py ./render-baseline ./render-current --report visual-diff-report.json
uv run python scripts/rebuild_visual_gate.py ./corpus --max-diff-ratio 0.01
```

Recommended CI order:
1. `run_corpus.py` (fast semantic/compatibility gate)
2. `interop_matrix.py` (multi-profile interop gate)
3. `rebuild_visual_gate.py` (end-to-end rebuild+render+visual gate)

`strict` profile is intentionally aggressive and may fail on heterogeneous real-world corpora; use it as a hardening target.

Testing script details and exit-code semantics: `scripts/README.md`.

Corpus profiles:

- `desktop`: fail on validation errors only
- `web`: fail on errors plus selected render-risk warnings (`UNRESOLVED_TEMPLATE_TOKEN`, `TABLE_VALIDATION_FAILED`, `CHART_VALIDATION_FAILED`)
- `strict`: fail on any warning or error

Contract details: see `docs/CONTRACT.md`.
Documentation:
- `docs/GETTING_STARTED.md`
- `docs/CONTRACT.md`
