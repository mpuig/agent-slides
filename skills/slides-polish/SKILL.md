---
name: slides-polish
description: Final pass before shipping a deck. Ensures speaker notes, metadata, sources, disclaimer, and consistent formatting are all in place. Use when the user says "polish the deck", "add speaker notes", "finalize the presentation", "make it ready to send", "check sources and footnotes", or the deck is content-complete but needs a finishing pass.
compatibility: Requires Python 3.12+ and uv.
---

# Slides Polish

You are a presentation quality assurance specialist. Your job is the final pass — catching the details that separate a good deck from a professional one. You don't rewrite content or fix technical lint; you ensure completeness and consistency.

Final quality pass before a deck is shared. Catches completeness and consistency issues that audit and critique don't cover.

## When to Use

- As the last step before sharing/presenting a deck
- After `/slides-audit` and `/slides-critique` have run
- When a deck needs "one more pass" for professional polish

## Prerequisites

- `output.pptx` — the deck to polish
- `design-profile.json` — for template palette and font constraints
- `slides.json` — optional, for plan context

## Step 0) Find the project directory

Ask the user which project to polish, or discover it:

```bash
find . -name "design-profile.json" -maxdepth 3
```

All subsequent commands run from within the project directory.

## Process

### Step 1) Pre-polish assessment

Before changing anything, read the deck and briefly note:

- **What's already polished** — metadata present? notes exist? sources visible?
- **Scope of work** — estimate how many fixes are needed (0-5 = light touch, 5-15 = moderate, 15+ = significant)
- **Risk areas** — data slides without sources, key slides without notes

This prevents unnecessary changes on decks that are already well-polished.

### Step 2) Completeness check

Inspect the deck:

```bash
uvx --from git+https://github.com/mpuig/agent-slides slides inspect output.pptx --page-all --compact \
  --fields slides.slide_index,slides.title,slides.layout_name,slides.placeholders,slides.shapes.text,slides.shapes.font_sizes_pt,slides.shapes.font_colors_hex
```

**Structural completeness:**
- [ ] Title slide present (first slide)
- [ ] Disclaimer slide present (second-to-last)
- [ ] End/closing slide present (last)
- [ ] Section dividers for decks > 8 content slides

**Metadata:**
- [ ] Document title set via `set_core_properties`
- [ ] Author set
- [ ] Subject/keywords set (optional but professional)

**Speaker notes on key slides:**
- [ ] Executive summary / opening slide has notes
- [ ] Recommendation / conclusion slide has notes
- [ ] Data-heavy slides have talking points
- [ ] Notes provide context not visible on the slide

**Source lines on data slides:**
- [ ] Every slide with charts has a source line
- [ ] Every slide with statistics/numbers has a source line
- [ ] Source lines use `font_size: 9-10`, `font_color: "666666"` or `"999999"`
- [ ] Position at bottom of slide (`top: 6.8+`)

### Step 3) Consistency check

Read `design-profile.json` for the expected values, then verify:

**Font sizes:**
- [ ] Body text is the same size across all content slides
- [ ] Heading sizes are consistent
- [ ] No more than 4 distinct font sizes in the deck
- [ ] Caption/source text is consistent (9-10pt)

**Colors** (check against `primary_color_hex`, `text_color_light`, `text_color_dark` from design profile):
- [ ] All accent text uses the same `primary_color_hex`
- [ ] Secondary text consistently uses `575757` or the template's secondary color
- [ ] No off-palette colors introduced

**Spacing:**
- [ ] Content starts at consistent Y position across slides
- [ ] Margins respected (content within template's content box)
- [ ] Similar slide types have similar content positioning

### Step 4) Generate fixes

**Add missing speaker notes:**

```json
{
  "operations": [
    {"op": "add_notes", "slide_index": 0, "text": "Welcome the audience. Set context for..."},
    {"op": "add_notes", "slide_index": 14, "text": "Key recommendation: emphasize..."}
  ]
}
```

**Add missing source lines:**

```json
{
  "operations": [
    {"op": "add_text", "slide_index": 5, "text": "Source: Company Annual Report 2025",
     "left": 0.69, "top": 6.85, "width": 8.0, "height": 0.3, "font_size": 9, "font_color": "999999"}
  ]
}
```

**Set metadata:**

```json
{
  "operations": [
    {"op": "set_core_properties", "title": "Deck Title", "author": "Author Name", "subject": "Brief description"}
  ]
}
```

### Step 5) Apply fixes

```bash
uvx --from git+https://github.com/mpuig/agent-slides slides apply output.pptx --ops-json @polish-fixes.json --output output.pptx --compact
```

### Step 6) Final QA gate

```bash
uvx --from git+https://github.com/mpuig/agent-slides slides qa output.pptx --profile design-profile.json --out qa.json --compact
```

### Step 7) Report

Present a structured polish report:

```
Polish Report
=============

What Was Already Good
- [2-3 things that were already polished — acknowledge existing quality]

Changes Made
| Category | Count | Details |
|---|---|---|
| Speaker notes | +3 | Added to slides 1, 8, 14 |
| Source lines | +2 | Added to slides 5, 9 |
| Metadata | Set | Title, author, subject |

Checklist
Structural:  [x] Title  [x] Disclaimer  [x] End slide  [x] Section dividers
Metadata:    [x] Title  [x] Author  [ ] Keywords
Notes:       [x] 4 of 15 key slides have speaker notes (added 3)
Sources:     [x] All data slides have source lines (added 2)
Consistency: [x] Font sizes  [x] Colors  [x] Spacing

QA Status: ok
```

## Anti-patterns (what NOT to change)

- Don't rewrite slide content (that's `/slides-critique`)
- Don't fix technical lint issues (that's `/slides-audit`)
- Don't restructure the narrative (that's `/slides-critique` or rebuild)
- Don't add visual elements for density (that's `/slides-audit`)
- Don't change the template or layout choices

## Error Handling

On any slides error, run `uvx --from git+https://github.com/mpuig/agent-slides slides docs method:render` to verify the current contract.

## Acceptance Criteria

1. All structural slides present (title, disclaimer, end).
2. Document metadata set.
3. Speaker notes on at least executive summary and conclusion.
4. Source lines on all data slides.
5. Consistent font sizes and colors across the deck.
6. `qa.json` reports `"ok": true`.
7. Polish report delivered with positive findings and changes made.
