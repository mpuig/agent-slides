#!/usr/bin/env bash
# Package all skills into a single agent-slides.zip for the Claude app.
#
# Structure (required by Claude):
#   agent-slides.zip
#     └── agent-slides/
#         ├── SKILL.md              (orchestrator)
#         ├── references/
#         │   ├── slides-extract.md
#         │   ├── slides-build.md
#         │   ├── slides-edit.md
#         │   ├── slides-audit.md
#         │   ├── slides-critique.md
#         │   ├── slides-polish.md
#         │   ├── slides-full.md
#         │   └── (all sub-references)
#
# Usage: bash scripts/package-skills.sh [output-dir]
#   output-dir defaults to website/downloads

set -euo pipefail

SKILLS_DIR="skills"
OUTPUT_DIR="${1:-website/downloads}"
STAGING_DIR="$(mktemp -d)"
PACKAGE_DIR="$STAGING_DIR/agent-slides"
REFS_DIR="$PACKAGE_DIR/references"

trap 'rm -rf "$STAGING_DIR"' EXIT

if [ ! -d "$SKILLS_DIR" ]; then
  echo "Error: skills directory not found at $SKILLS_DIR" >&2
  exit 1
fi

mkdir -p "$REFS_DIR"

# Build the orchestrator SKILL.md
cat > "$PACKAGE_DIR/SKILL.md" << 'SKILL_EOF'
---
name: agent-slides
description: Generate professional PowerPoint decks from a single prompt. Extract templates, build slides, edit, audit, critique, and polish presentations. Use when the user wants to create, edit, or review a PowerPoint deck.
compatibility: Requires Python 3.12+ and uv.
---

# agent-slides

You are a presentation expert. You generate professional, brand-compliant PowerPoint decks using the `slides` CLI tool, installed via `uvx`.

## CLI Tool

All commands use:
```bash
uvx --from agent-slides slides <subcommand> [args]
```

Key subcommands: `extract`, `render`, `apply`, `inspect`, `validate`, `lint`, `qa`, `find`, `edit`, `transform`, `docs`.

Run `uvx --from agent-slides slides docs json` for the full schema and operation reference.

## Workflows

Choose the right workflow based on what the user needs:

### Full pipeline (most common)
When the user wants a complete deck from a template and a brief, follow [references/slides-full.md](references/slides-full.md).

### Step-by-step
1. **Extract template** → [references/slides-extract.md](references/slides-extract.md)
2. **Build deck** → [references/slides-build.md](references/slides-build.md)
3. **Audit** → [references/slides-audit.md](references/slides-audit.md)
4. **Critique** → [references/slides-critique.md](references/slides-critique.md)
5. **Polish** → [references/slides-polish.md](references/slides-polish.md)

### Edit existing deck
When the user wants to modify an existing deck → [references/slides-edit.md](references/slides-edit.md)

## Typical Flow

```
Extract template → Build deck → Audit → Critique → Polish
                                  ↕         ↕
                               Edit ← (targeted fixes)
```

## Key Rules

- Always extract a template before building — never generate slides without template contracts.
- Every deck needs: title slide + content slides + disclaimer + end slide.
- Use `--dry-run` before rendering to catch validation errors early.
- Use `--compact` on all CLI output to save context window space.
- Run QA after every render to verify design contract compliance.
SKILL_EOF

# Copy each skill's SKILL.md as a reference (strip frontmatter name/description)
for skill_dir in "$SKILLS_DIR"/*/; do
  skill_name="$(basename "$skill_dir")"
  if [ ! -f "$skill_dir/SKILL.md" ]; then
    continue
  fi

  # Copy the full SKILL.md content as a reference file
  cp "$skill_dir/SKILL.md" "$REFS_DIR/$skill_name.md"

  # Copy sub-references if they exist
  if [ -d "$skill_dir/references" ]; then
    for ref_file in "$skill_dir"/references/*.md; do
      [ -f "$ref_file" ] || continue
      ref_basename="$(basename "$ref_file")"
      cp "$ref_file" "$REFS_DIR/${skill_name}--${ref_basename}"
    done
  fi
done

# Fix reference paths in copied files (references/ → skill-name--filename)
for ref_file in "$REFS_DIR"/*.md; do
  [ -f "$ref_file" ] || continue
  skill_prefix="$(basename "$ref_file" .md)"
  # Only fix files that have sub-references (the skill files, not the sub-refs)
  if [[ "$skill_prefix" != *"--"* ]]; then
    sed -i.bak "s|references/\([^)]*\)|references/${skill_prefix}--\1|g" "$ref_file"
    rm -f "$ref_file.bak"
  fi
done

# Create the ZIP
mkdir -p "$OUTPUT_DIR"
rm -f "$OUTPUT_DIR/agent-slides.zip"
(cd "$STAGING_DIR" && zip -rq - agent-slides) > "$OUTPUT_DIR/agent-slides.zip"

file_count=$(unzip -l "$OUTPUT_DIR/agent-slides.zip" | tail -1 | awk '{print $2}')
echo "Packaged agent-slides.zip ($file_count files) to $OUTPUT_DIR"
