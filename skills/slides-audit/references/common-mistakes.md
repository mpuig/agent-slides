# Common Mistakes — Technical Audit Subset

Technical issues relevant to deck audit, extracted from the full 25-item ranked list. Focus on pixels, fonts, colors, overlap, and layout compliance.

---

## Visual Quality

6. **No visual hierarchy** — Title must visually dominate (largest, brand color), then headings (bold), then body (regular). Without size and color contrast, everything looks flat.

7. **Slides without visual structure** — Every content slide needs at least one visual element: chart, table, structured layout, or colored accent. Text-only slides are forgettable.

9. **Ignoring template colors** — Use the extracted palette from the design profile and color zones. Don't introduce arbitrary colors that clash with the template's visual identity.

10. **Pure black text** — Most professional templates use dark gray (e.g., `333333`, `575757`) for body text, not `000000`. Check the template's extracted text colors.

## Layout Compliance

12. **Cramming content** — Leave breathing room. White space is a design feature. Maintain consistent margins and spacing from the template's content box.

13. **Content overlapping the title** — Content must start below the title area. Never place shapes or text in the title zone (typically y > 1.8" means safe; y < 1.8" overlaps the title).

25. **Ignoring the template's content box** — Place content within the extracted `content_box` boundaries. Content outside these bounds may overlap with template decorations (logos, page numbers, borders).

## Formatting Consistency

21. **Inconsistent font sizes** — Pick size constants at planning time and reuse them throughout. Body text should be the same size on every content slide.

22. **Centered body text** — Body text should be left-aligned. Center alignment only for titles, card labels, and big callout numbers.

23. **Too many font sizes** — Use at most 4 distinct sizes in a deck: title, heading, body, caption. More than that creates visual noise.
