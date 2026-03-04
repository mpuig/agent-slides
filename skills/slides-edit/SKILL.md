---
name: slides-edit
description: Edit an existing presentation deck. Supports text edits, layout transforms, and ops-based patches. Use when the user wants to change text on a slide, update a title, fix a typo, swap a layout, replace chart data, add or remove slides, or says things like "change slide 3 title to X", "update the numbers on the chart", "move the agenda slide".
compatibility: Requires Python 3.12+, uv, and agent-slides in the current workspace.
---

# Slides Edit

You are a slide editor. Your job is to make precise, targeted changes to existing decks without breaking what already works.

Modify an existing deck — text changes, layout transforms, and structural edits.

## When to Use

- Fix typos or update text in an existing deck
- Apply archetype transforms to restyle slides
- Add/remove/move slides via ops patches
- Any targeted modification that doesn't require a full rebuild

## Prerequisites

- An existing `output.pptx` in a project directory
- Optionally: `slides.json` (for context) and `design-profile.json` (for QA)

## Step 0) Find the project directory

Ask the user which project to edit, or discover it:

```bash
find . -name "design-profile.json" -maxdepth 3
```

All subsequent commands run from within the project directory.

## Process

### Step 1) Locate targets

Inspect the deck to find slide/shape UIDs:

```bash
uv run slides inspect output.pptx \
  --fields slides.slide_uid,slides.shapes.shape_uid,slides.title \
  --out ids.json --compact
```

Search for specific text:

```bash
uv run slides find output.pptx --query "<search text>" --limit 10 \
  --out find.json --compact
```

Pagination for large decks:

```bash
uv run slides inspect output.pptx --page-size 5 --page-token 0 --compact
```

Other inspection:

```bash
uv run slides inspect output.pptx --placeholders 0 --compact   # placeholders on slide 0
uv run slides inspect output.pptx --summary --compact           # deck summary
```

### Step 2) Pre-edit assessment

Before making changes, briefly note:

- **What works well** in the current deck (preserve these strengths)
- **Scope of change** — which slides are affected and which are untouched
- **Risk** — could this edit break layout, contrast, or narrative flow?

This prevents over-editing and protects existing quality.

### Step 3) Apply changes

**Text edits** (find-and-replace scoped by UID):

```bash
uv run slides edit output.pptx --query "old text" \
  --replacement "new text" --slide-uid "<slide_uid>" \
  --shape-uid "<shape_uid>" --output output.pptx --compact
```

Alternative selectors: `--slide <index>`, `--slide-id <slide-N>`, `--shape-id <shape_id>`.

**Archetype transforms** (restyle a slide):

```bash
uv run slides transform output.pptx --slide-uid "<slide_uid>" \
  --to timeline --output output.pptx --compact
```

**Ops-based patches** (apply additional operations):

```bash
uv run slides apply output.pptx --ops-json @patch_ops.json --output output.pptx --compact
```

Write `patch_ops.json` as:
```json
{
  "operations": [
    {"op": "replace_text", "slide_index": 3, "old": "Draft", "new": "Final"},
    {"op": "add_text", "slide_index": 5, "text": "New insight", "left": 1.0, "top": 5.0, "width": 4.0, "height": 0.5, "font_size": 16}
  ]
}
```

### Step 4) Verify changes

```bash
uv run slides find output.pptx --query "new text" --compact
```

### Step 5) Re-run QA

```bash
uv run slides qa output.pptx --profile design-profile.json \
  --slides-json @slides.json --out qa.json --compact
```

### Step 6) Repair (if needed)

```bash
uv run slides repair output.pptx --output output.pptx
```

## Placeholder Rules

- **Never use `set_placeholder_text` with guessed indices.** Use `set_semantic_text` with `role` (`title`, `subtitle`, `body`) for standard placeholders.
- Only use `set_placeholder_text` with exact `idx` from `inspect` or `template_layout.json`.
- `--query` only replaces text content, not formatting (font size, color, bold). To change formatting, use ops patches with new `add_text` ops.

## Anti-patterns (what NOT to do)

- Don't use `--query` to fix formatting (font size, color, bold) — it only replaces text content. Use ops patches instead.
- Don't guess placeholder indices — use `set_semantic_text` with `role` or exact `idx` from inspect
- Don't edit without inspecting first — always locate targets with `inspect` or `find`
- Don't overwrite the input file without verifying the edit worked (`find` or `inspect` after)
- Don't restructure narrative or rewrite content (that's `/slides-critique` or `/slides-build`)

## Error Handling

On any slides error, run `uv run slides docs method:edit` to verify the current contract before retrying.

## Acceptance Criteria

1. Edits are verifiable via `find` or `inspect` subcommands.
2. `qa.json` reports `"ok": true`.
3. No unresolved-token or contract-critical issues.
