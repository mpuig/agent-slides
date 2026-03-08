---
name: slides-full
description: End-to-end deck pipeline — extract, build, audit, critique, and polish in one pass. Use when the user wants a complete deck from scratch with a template, says "create a full deck", "build me a presentation end to end", "make a polished deck from this template", or provides both a brief and a template and expects a finished, presentation-ready result. Prefer this over chaining individual skills manually.
compatibility: Requires Python 3.12+ and uv.
---

# Slides Full

You are a presentation production manager. Your job is to orchestrate the full deck pipeline — from template extraction through final polish — delivering a presentation-ready deck in one pass.

End-to-end pipeline that chains all slides skills: extract, build, audit, critique, polish.

## When to Use

- User wants a complete deck from scratch in one command
- User provides a brief + template and wants a finished, polished result
- When you want to skip invoking each skill manually

## Prerequisites

- A `.pptx` template file (or an existing project directory with contracts already extracted)
- A brief, topic, or content outline from the user

## Arguments

The user may provide:
- Template path (required if no project directory exists)
- Brief / topic / content description
- Target slide count (default: inferred from brief complexity)
- Project name (default: inferred from brief)

## Process

Use a dynamic state machine, not a fixed linear pipeline.

State flow:

1. `EXTRACT_OR_REUSE`
2. `BUILD_OR_UPDATE`
3. `GLOBAL_CONTENT_CHECK`
4. `LOCAL_VISUAL_CHECK`
5. `APPLY_FIXES`
6. `RECHECK`
7. `DONE`

Transition rule: any failed gate returns to `APPLY_FIXES`, then `RECHECK`.

### Step 0) Find the project directory

Ask the user which project to work on, or discover it:

```bash
find . -name "design-profile.json" -maxdepth 3
```

All subsequent commands run from within the project directory.

### Step 1) Extract and Profile

If `output/<project>/resolved_manifest.json`, `base_template.pptx`, and `design-profile.json` all exist, skip to Step 2.

Otherwise run:

```bash
mkdir -p output/<project>
uvx --from git+https://github.com/mpuig/agent-slides slides extract <template.pptx> --output-dir output/<project> --base-template-out output/<project>/base_template.pptx --compact
```

**Comprehension gate** — after extraction, read `resolved_manifest.json` and confirm:

- What accent colors does the theme use? List the hex values. (Path: `theme.palette.accent1` … `accent6`)
- How many archetypes have `resolved_layouts`? Name them. (Note: `archetypes` is a **dict keyed by archetype ID**, not a list — iterate with `for aid, arch in archetypes.items()`)
- Which layouts are split-panel (title in a side zone)?
- What `text_color` does each color zone use?

If any answer is unclear, re-read the manifest before proceeding.

**Build design profile** — write `design-profile.json` per the `/slides-extract` skill:

```json
{
  "name": "<project-name>",
  "template_path": "base_template.pptx",
  "content_layout_catalog_path": "content_layout.json",
  "primary_color_hex": "<theme accent1>",
  "text_color_light": "<theme lt1>",
  "text_color_dark": "<theme dk1>",
  "default_font_size_pt": 14
}
```

Add `"icon_pack_dir": "icons"` if extraction produced an `icons/` directory. Use `uvx --from git+https://github.com/mpuig/agent-slides slides docs schema:design-profile` for the full schema. Only add fields listed in the schema — the profile uses `extra="forbid"`.

### Step 2) Build

Generate `slides.json` and render. Follow the `/slides-build` skill rules:

- **Read references** — load `references/storytelling.md` before planning, then `references/content-density.md`, `references/layout-patterns.md`, and `references/operations.md` before generating ops.
- **Read `resolved_manifest.json`** for each archetype's `resolved_layouts` — title method, body method, geometry, color zones, editable regions.
- **Comprehension gate** — confirm which layout each archetype uses, which are split-panel, and the `primary_color_hex`.
- **Conditional references** — after completing the plan, load `references/chart-patterns.md`, `references/framework-patterns.md`, or `references/process-patterns.md` only if the plan includes matching slide types.
- **Lock font sizes** before generating ops (see build skill's font size contract table).

```bash
uvx --from git+https://github.com/mpuig/agent-slides slides render --slides-json @slides.json --profile design-profile.json --dry-run --compact
uvx --from git+https://github.com/mpuig/agent-slides slides render --slides-json @slides.json --profile design-profile.json --output output/<project>/output.pptx --compact
```

### Step 3) Global Content Check

Run global checks on `slides.json`:

```bash
uvx --from git+https://github.com/mpuig/agent-slides slides plan-inspect --slides-json @slides.json --out output/<project>/plan_content.json --compact
```

Use this file for storytelling checks:
- Flow and section coverage (requires seeing structural slides too)
- Action-title quality
- Message duplication
- Role sequencing

Use `--content-only` or `--summary-only` only when drilling into specific subsets. Do not read full `slides.json` unless needed.

**Read `references/common-mistakes.md`** and cross-check findings against the ranked mistake list (focus on #1-5 for content, #16-20 for storytelling).

### Step 4) Local Visual Check

Run technical and visual checks on the rendered deck:

```bash
uvx --from git+https://github.com/mpuig/agent-slides slides lint output/<project>/output.pptx --profile design-profile.json --slides-json @slides.json --out output/<project>/lint.json --compact
uvx --from git+https://github.com/mpuig/agent-slides slides qa output/<project>/output.pptx --profile design-profile.json --slides-json @slides.json --out output/<project>/qa.json --compact
```

If needed, inspect one page/slide at a time:

```bash
uvx --from git+https://github.com/mpuig/agent-slides slides inspect output/<project>/output.pptx --page-size 1 --page-token <n> --out output/<project>/inspect_page.json --compact
```

Cross-check against common mistakes #6-13 (visual hierarchy, template colors, overlap, bounds).

### Step 5) Apply Fixes

Build fixes by class:
- `story.*` from global check
- `visual.*` from local check
- `contract.*` from lint/qa hard failures

Apply with small, reversible patches:

```bash
uvx --from git+https://github.com/mpuig/agent-slides slides apply output/<project>/output.pptx --ops-json @output/<project>/fixes_ops.json --output output/<project>/output.pptx --compact
```

### Step 6) Recheck Loop

Retry budgets:

1. Global loop max: 3 iterations
2. Local loop max: 2 iterations per affected slide
3. Stop early if issue count is not improving

Always re-run:
- global content check if text/structure changed
- local visual checks for touched slides, then final full `qa`

### Output Size Rules

Keep context small by default:

1. Use `--compact` on all commands
2. Write all outputs to files via `--out`
3. Use `plan-inspect --summary-only` for quick overview, drill in with `--content-only` or `--fields`
4. Use pagination (`--page-size`, `--page-token`) for inspect/find
5. Use `--verbose` only for debugging

### Failure Policy

1. `qa.ok == false` with contract/data errors: block release
2. Only visual/story warnings: continue if iteration budget is exhausted, but report remaining risks

### Final Report

Report:

1. Iteration counts (global/local)
2. Initial vs final issue counts by code
3. Remaining warnings
4. Final artifacts:
   - `output/<project>/output.pptx`
   - `output/<project>/qa.json`
   - `output/<project>/lint.json`
   - `output/<project>/plan_content.json`

## Anti-patterns (what NOT to do)

- Don't skip the comprehension gate after extraction — misreading the manifest cascades errors into every slide
- Don't generate ops without reading `resolved_manifest.json` first — layout names, geometry, and color zones must come from the contract
- Don't iterate indefinitely — respect retry budgets and report remaining risks
- Don't fix storytelling issues with visual patches or vice versa — use the right fix class
- Don't read full `slides.json` or full inspect output when summary/pagination suffices — protect context window

## Error Handling

On any slides error, run `uvx --from git+https://github.com/mpuig/agent-slides slides docs method:render` or `uvx --from git+https://github.com/mpuig/agent-slides slides docs method:extract` to verify the current contract before retrying.

## Acceptance Criteria

1. `output.pptx` exists and opens.
2. `qa.json` reports `"ok": true` for release mode.
3. Global content checks pass or are explicitly waived.
4. Visual checks are below warning threshold or explicitly waived.
5. All fixes are traceable via small ops patches.
6. Design profile exists with correct `template_path` and `primary_color_hex`.
7. All content slides have action titles (complete sentences).
