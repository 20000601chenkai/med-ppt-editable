# Review and Acceptance

1. **Content:** Check source-to-slide claims, units, denominators, dates, and missing-data labels. Confirm references correspond to supplied or verified sources. Review patient information before any public sharing.
2. **Structure:** Run `scripts/audit_pptx.py` with the build manifest. Resolve errors. Review warnings individually; a large clinical image may be intentional. The script exits 1 for audit errors, 2 for invalid/unreadable input, and 0 if no structural errors were found.
3. **Rendering:** Render the actual PPTX, inspect all pages for clipping, font substitution, blank media, label collisions, and chart readability. If no office renderer is available, disclose that only structural validation ran.
4. **Editing:** Change one title, one table cell, and chart data in the target application, save a copy, and reopen it when application automation or manual review is available. Package structure alone does not establish that the target application's editing UI works.

Record target application and version, date, slide count, audited native tables/charts, image exceptions, and known limitations. WPS, Keynote, PowerPoint and LibreOffice do not render every object identically. Do not infer compatibility with an untested application.

For a supplied template, inspect its rendered slides as well as object geometry, fonts, image cropping and transparency. Reproduce the relevant visual system with editable objects; keep the original file intact. A full-slide photograph used as a faint background can be intentional and should be recorded as such when the audit flags it. Do not add user-supplied template assets to a public skill package without authorization.

Review footnotes and table rows at actual slide size. A text box can remain within slide bounds while its rendered text wraps into the footer; a bounds-only audit will not detect this. Re-render changed slides after layout fixes. Retain full references in notes when the visible source line must be shortened.

Review composition separately from clipping. At presentation scale and in a contact sheet, check for an accidentally empty quadrant, undersized section numbers, repeated text-only layouts, and diagrams too small to function as visual anchors. Follow [visual-composition.md](visual-composition.md) for remedies. Passing an object-bounds audit does not establish that a page has an effective visual hierarchy. For a visual revision, retain a page-level list of changes and inspect every affected page in the rendered output.

## Existing Deck Changes

Work on a copy. `python-pptx` does not round-trip every Office feature, so do not use a load/save pass as a universal preservation guarantee. For animation, embedded objects, equations or custom XML, prefer the originating Office application or a narrowly scoped OOXML patch and validate it. Keep the latest user-edited deck and update authoring data deliberately.

## Cost

The included generator is local and does not request paid APIs. The surrounding AI host and optional external image/research services may cost money. A smaller entrypoint and structured renderer are intended to reduce repeated authoring work; token savings have not been benchmarked.
