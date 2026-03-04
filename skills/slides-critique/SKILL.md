---
name: slides-critique
description: Review an existing deck for storytelling quality, visual hierarchy, and content effectiveness. Identifies weak action titles, MECE violations, isomorphism mismatches, and density issues. Use when the user says "review my deck", "critique the presentation", "are the slides telling a good story", "check the narrative flow", "improve the slide titles", or wants feedback on content quality rather than technical formatting.
compatibility: Requires Python 3.12+, uv, and agent-slides in the current workspace.
---

# Slides Critique

You are a presentation strategist and storytelling coach. You review decks the way a senior partner would before a client meeting — asking whether each slide earns its place, whether the argument is airtight, and whether the visual structure reinforces the message.

Storytelling and design review for an existing deck. Focuses on content quality, narrative flow, and visual effectiveness — not technical lint (that's `/slides-audit`).

## When to Use

- After `/slides-build` to improve content quality
- When a deck is technically correct but unconvincing
- When the narrative feels disjointed or flat

## Prerequisites

- `output.pptx` — the deck to critique
- `design-profile.json` — for template palette context
- `slides.json` — for the original plan (story roles, archetypes, action titles)

## References

| File | When to load | Content |
|---|---|---|
| `references/storytelling.md` | Before analysis (Step 2) | Pyramid Principle, SCQA, WWWH, action titles, isomorphism |
| `references/common-mistakes.md` | During analysis (Step 2) | Focus on #1-5 (critical) and #16-20 (content quality) |

Load `references/storytelling.md` and `references/common-mistakes.md` before starting the review.

## Step 0) Find the project directory

Ask the user which project to critique, or discover it:

```bash
find . -name "design-profile.json" -maxdepth 3
```

All subsequent commands run from within the project directory.

## Process

### Step 1) Read deck content

```bash
uv run slides inspect output.pptx --page-all --compact \
  --fields slides.slide_index,slides.title,slides.layout_name,slides.shapes.text,slides.shapes.font_sizes_pt,slides.shapes.kind
```

Also read `slides.json` to understand the intended plan (story roles, archetype choices).

### Step 2) Evaluate against criteria

Score each slide and the overall deck against these dimensions:

#### A. Action Titles (Critical)

For each content slide, check:
- Is the title a complete sentence stating the "so what"? (not a topic label)
- Does the body content prove the title claim?
- Would an executive understand the slide's point from the title alone?

**Bad:** "Market Overview", "Key Findings", "Next Steps"
**Good:** "European market grew 23% YoY driven by premium segment", "Three operational gaps cost $12M annually"

#### B. Narrative Flow

Check the slide sequence:
- Does the deck follow a clear structure? (SCQA, Pyramid, WWWH)
- Are section dividers present for decks > 8 slides?
- Are there orphan slides (single slide in a "section")?
- Does the opening set up the problem and the closing deliver the recommendation?

#### C. Isomorphism (visual structure matches conceptual relationship)

For each slide, check archetype-content match:
- Equal columns for truly equal items? (not for hierarchical data)
- Charts for quantitative data? (not bullets)
- Timeline for sequential items? (not cards)
- Tables for comparison data? (not paragraphs)

#### D. Visual Hierarchy

For each content slide, check:
- Are there 3+ distinct text sizes creating hierarchy?
- Is the title visually dominant?
- Are there visual elements beyond plain text?
- Are bullet lists kept to 6 or fewer items?

#### E. Content Density

For each slide, check:
- Can the "so what" be understood in 5 seconds?
- Is there enough white space?
- Is the content area filled (not sparse, not cramped)?

#### F. Layout Variety

Check across consecutive slides:
- Is the same layout used 3+ times in a row?
- Is there a mix of full-width, split-panel, and accent layouts?

#### G. Parallel Structure

For slides with columns, cards, or lists:
- Do items follow the same grammatical pattern?
- Are descriptions the same approximate length?

### Step 3) Identify what's working

Before listing problems, note 3-5 things the deck does well. Examples:
- "Strong action titles on slides 3, 7, 11 — each makes a clear claim backed by data"
- "Good use of split-panel layouts to create visual variety"
- "Source lines present on all data slides"

This prevents over-correction and acknowledges existing quality.

### Step 4) Prioritize findings

Rank findings by impact:
1. **Critical** — Action titles, body-doesn't-prove-title, missing narrative structure
2. **Major** — Isomorphism violations, no visual hierarchy, bullet-heavy slides
3. **Minor** — Layout variety, parallel structure, density fine-tuning

### Step 5) Provocative questions

Ask 2-3 questions that challenge assumptions:
- "If the audience only sees one slide, which one carries the entire argument? Does it?"
- "Could slides 4-6 be collapsed into a single data slide without losing the argument?"
- "The recommendation says X, but the evidence on slides 8-10 points to Y — is the deck arguing against itself?"

These questions force deeper thinking about whether the deck truly makes its case.

### Step 6) Fix what's fixable

**Text fixes** (action titles, copy improvements):

```bash
uv run slides edit output.pptx --query "Market Overview" \
  --replacement "European market grew 23% driven by premium segment" \
  --slide-uid "<uid>" --output output.pptx --compact
```

**Structural fixes** (add visual elements, adjust hierarchy):

Write `critique-fixes.json` with ops, then:

```bash
uv run slides apply output.pptx --ops-json @critique-fixes.json --output output.pptx --compact
```

### Step 7) Report structural issues

Some findings can't be auto-fixed and need a rebuild:
- Wrong archetype for the content (needs layout change)
- Missing section dividers (needs new slides)
- Slides that should be split or merged
- Fundamental narrative restructuring

Report these to the user with specific recommendations.

### Step 8) Re-run QA

```bash
uv run slides qa output.pptx --profile design-profile.json --out qa.json --compact
```

## Scoring

After analysis, present a summary scorecard:

```
Critique Report
===============

What's Working
- [3-5 specific positive findings with slide references]

Scorecard
| Dimension | Score | Key Issue |
|---|---|---|
| Action Titles | Strong/Weak/Mixed | "3 of 15 slides use topic labels" |
| Narrative Flow | Strong/Weak | "Missing section dividers" |
| Isomorphism | Strong/Weak | "Slide 8 uses bullets for comparison data" |
| Visual Hierarchy | Strong/Weak | "5 slides lack hierarchy" |
| Content Density | Strong/Weak | "Slide 12 is overcrowded" |
| Layout Variety | Strong/Weak | "Title Only used 6 times consecutively" |
| Parallel Structure | Strong/Weak | "Cards on slide 5 have inconsistent format" |

Provocative Questions
1. [question challenging the deck's argument]
2. [question about slide necessity or ordering]

Fixes Applied
- [list of auto-applied text/structural fixes]

Requires Rebuild
- [list of issues that need manual intervention]
```

## Anti-patterns (what NOT to do)

- Don't fix technical lint issues (fonts, overlap, bounds) — that's `/slides-audit`
- Don't add speaker notes or metadata — that's `/slides-polish`
- Don't change archetype/layout choices without user approval — report as recommendation
- Don't rewrite body content unless the title claim is unsupported
- Don't apply subjective style preferences — stick to the criteria above

## Error Handling

On any slides error, run `uv run slides docs method:edit` to verify the current contract.

## Acceptance Criteria

1. All content slides have action titles (complete sentences).
2. Body content supports title claims.
3. No isomorphism violations on data slides.
4. Visual hierarchy present on all content slides.
5. Critique report delivered with positive findings, scorecard, and provocative questions.
