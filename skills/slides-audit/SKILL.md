---
name: slides-audit
description: Run technical quality checks on an existing deck. Finds and fixes font size violations, shape overlaps, contrast issues, missing sources, and layout compliance problems. Use when the user says "check the deck for issues", "run QA", "lint the slides", "are there any formatting problems", "audit the presentation", or wants to verify visual quality before sharing.
compatibility: Requires Python 3.12+ and uv.
---

# Slides Audit

You are a presentation quality engineer. Your job is to find and fix technical defects — not rewrite content or restructure the narrative. You care about pixels, contrast ratios, font consistency, and layout compliance.

Technical quality analysis and automated fixes for an existing deck.

## When to Use

- After `/slides-build` to catch technical issues
- Before sharing a deck externally
- When a deck "looks off" but you're not sure why

## Prerequisites

- `output.pptx` — the deck to audit
- `design-profile.json` — design constraints for linting

## References

| File | When to load | Content |
|---|---|---|
| `references/common-mistakes.md` | During analysis (Step 2) | 25 ranked mistakes — focus on #6, #7, #9, #10, #12, #13, #21, #25 |

Load `references/common-mistakes.md` to cross-check findings against the ranked mistake list.

## Step 0) Find the project directory

Ask the user which project to audit, or discover it:

```bash
find . -name "design-profile.json" -maxdepth 3
```

All subsequent commands run from within the project directory.

## Process

### Step 1) Collect diagnostics

Run lint and inspection in parallel:

```bash
uvx --from agent-slides slides lint output.pptx --profile design-profile.json --out lint.json --compact

uvx --from agent-slides slides inspect output.pptx --page-all --ndjson --compact

uvx --from agent-slides slides validate output.pptx --compact

uvx --from agent-slides slides inspect output.pptx --summary --compact
```

### Step 2) Analyze findings

Categorize each lint issue by severity and fixability:

**Auto-fixable (generate patch ops):**

| Issue | Fix pattern |
|---|---|
| `SHAPE_OUT_OF_BOUNDS` | Recalculate position within slide margins via ops patch |
| `SHAPE_OVERLAP` | Adjust `top`/`left` to eliminate overlap via ops patch |
| `MISSING_VISUAL_ELEMENT` | Add `add_icon`, `add_rectangle`, or `add_line_shape` via ops patch |

**Requires judgment (report to user):**

| Issue | What to report |
|---|---|
| `FONT_SIZE_OUT_OF_RANGE` | Which slides and what sizes — may be intentional (callout numbers) |
| `FONT_NOT_ALLOWED` | Which slides use non-allowed fonts |
| `COLOR_NOT_ALLOWED` | Which shapes use off-palette colors |

**Informational (skip):**

| Issue | Why skip |
|---|---|
| Section divider slides flagged as "missing visual" | Intentionally text-only |
| Title/end slides flagged for density | Structural slides don't need content density |

**Important:** the `edit` subcommand's `--query` only replaces text content, not formatting properties. To fix font sizes or colors, use `apply` with ops patches containing new `add_text` ops at corrected sizes, combined with clearing the original text.

### Step 3) Check contrast

For each slide, verify font color vs. background:

- Read the slide's layout from `inspect` output
- Check `color_zones[].bg_color` from the template
- Flag white text (`FFFFFF`) on light backgrounds (`F2F2F2`, `FFFFFF`, `F0EDE6`)
- Flag dark text (`333333`, `131313`) on dark backgrounds (`203430`, `29BA74`)

### Step 4) Check content overlap

For slides with `SHAPE_OVERLAP` warnings:

- Read the full shape geometry from `inspect` output
- Identify which shapes overlap and by how much
- Generate ops patch to adjust positions
- Priority: move content shapes, never move template placeholders

### Step 5) Generate and apply fixes

Write `audit-fixes.json` with patch ops:

```bash
uvx --from agent-slides slides apply output.pptx --ops-json @audit-fixes.json --output output.pptx --compact
```

### Step 6) Re-run QA

```bash
uvx --from agent-slides slides qa output.pptx --profile design-profile.json --out qa.json --compact
```

### Step 7) Report

Present a structured audit report:

```
Audit Report
============

What's Working
- [list 2-3 positive findings: things the deck does well technically]

Issues Found: X total (Y auto-fixed, Z need attention)

Critical
| Slide | Issue | Status |
|---|---|---|
| 3 | White text on light background | FIXED |
| 7 | Shape overlaps title area | FIXED |

Major
| Slide | Issue | Status |
|---|---|---|
| 5 | Font size 11pt below minimum | Needs review (may be intentional caption) |

Minor
| Slide | Issue | Status |
|---|---|---|
| 12 | Off-palette color #8B4513 | Reported |

Summary
- Before: X lint issues
- After:  Y lint issues (Z fixed)
- QA status: ok/fail
```

## Common Mistakes to Check

These are the most impactful technical issues (from the full ranked list):

| # | Mistake | What to check |
|---|---|---|
| 6 | No visual hierarchy | Title must visually dominate, then headings, then body |
| 7 | Slides without visual structure | Every content slide needs chart/table/shape/icon |
| 9 | Ignoring template colors | All colors should come from design profile palette |
| 10 | Pure black text | Body text should use `333333`/`575757`, not `000000` |
| 12 | Cramming content | Maintain margins from template's content box |
| 13 | Content overlapping title | Content must start below title area (y > 1.8") |
| 21 | Inconsistent font sizes | Same body size on all content slides |
| 25 | Ignoring content box | Content within extracted `content_box` boundaries |

## Anti-patterns (what NOT to fix)

- Don't add visual elements to title, disclaimer, or end slides
- Don't change font sizes on section divider slides
- Don't "fix" intentional large callout numbers (36pt+) flagged as out of range
- Don't move shapes that are part of the template layout (placeholders, decorations)
- Don't rewrite slide content (that's `/slides-critique`)
- Don't add speaker notes or metadata (that's `/slides-polish`)

## Error Handling

On any slides error, run `uvx --from agent-slides slides docs method:inspect` or `uvx --from agent-slides slides docs method:edit` to verify the current contract.

## Acceptance Criteria

1. Lint issue count decreases after fixes.
2. No `SHAPE_OVERLAP` issues remain on content slides.
3. No contrast violations (white-on-light or dark-on-dark).
4. `qa.json` reports `"ok": true`.
5. Audit report delivered with both positive findings and issues.
