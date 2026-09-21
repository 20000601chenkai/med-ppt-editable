---
name: med-ppt-editable
description: Create editable PowerPoint presentations for medical case discussions, journal clubs, and academic reports. Use for medical PPT creation with native text, tables, charts, and evidence traceability; not for clinical diagnosis or treatment decisions.
---

# Medical Editable PPT

Produce a presentation a clinician or researcher can revise in PowerPoint. Prioritize readable medical content, traceable evidence, and native objects over decorative effects. Match the user's language. This package is a focused v0.1 authoring workflow, not a general converter for arbitrary existing decks.

## Start From the Materials

Read the provided materials and identify audience, talk duration, intended slide count, and whether this is a case discussion, journal club, or research report. Infer routine style choices; ask only for information that changes the content or deliverable. A supplied template or requested outline takes priority over the example layouts.

Read [references/medical-content.md](references/medical-content.md) when organizing medical material. Preserve original files; copy inputs into a working folder if needed. Do not move them or overwrite the input presentation.

Separate source facts, the author's interpretation, missing information, and synthetic demonstrations. Never fill a missing patient detail, citation, result, dose, unit, effect size, or confidence interval with a plausible invention. Mark missing items in the draft and notes. Public examples must use synthetic or explicitly authorized de-identified material; hidden notes and cropped source images can also contain identifiers.

## Build With Native Objects

Before choosing layouts or revising a sparse, text-heavy deck, read [references/visual-composition.md](references/visual-composition.md). Plan a meaningful visual relationship for each content slide: process, mechanism, comparison, location, evidence, or collaboration. Preserve the requested density and template language; avoid repeating the same text columns throughout a treatment section. Section pages need a balanced visual anchor and readable numbered labels, not a large accidental empty quadrant.

- Text: editable text frames containing paragraphs; avoid one box per character or forced line.
- Tables: native PowerPoint tables, with units and explicit missing-data labels.
- Supported data charts: native charts with an embedded workbook. A collection of bars and labels is not a data-editable chart.
- Timelines and simple diagrams: native shapes and text. Complex scientific plots may remain vector or bitmap figures with their data and plotting source, and a clear limitation recorded.
- Clinical photographs and paper figures: individual replaceable image objects with source and caption. Preserve aspect ratio; no diagnostic image enhancement or AI replacement of clinical evidence.
- Equations, SmartArt, animation, and arbitrary PPTX template preservation are outside the bundled builder's v0.1 scope. Use an appropriate tool for a user-requested extension and report the actual result, not a blanket claim that everything is editable.

For custom follow-up line charts, mixed units, overlapping values, or appending chart slides, read [references/native-charts.md](references/native-charts.md). It covers percentage conventions, point-label formatting, independent embedded workbooks and preserving previously edited slides. These are custom-authoring guidelines, not additional JSON fields supported by the bundled builder.

For the bundled layouts, read [references/deck-format.md](references/deck-format.md). Write a UTF-8 deck JSON in the work folder and use the relative script paths below, resolved from this skill folder:

```sh
python -m pip install -r requirements.txt
python scripts/build_deck.py /path/to/deck.json --output /path/to/deck.pptx
python scripts/audit_pptx.py /path/to/deck.pptx --manifest /path/to/deck.manifest.json --output /path/to/deck.audit.json
```

Reuse already-installed dependencies or an isolated environment. Generation itself makes no network calls and requires no paid image service. The host AI service may have its own cost. A custom layout may use `python-pptx` directly; retain the same native-object contract and audit it.

Use `examples/case-demo.json` for the schema and a synthetic case demonstration. Use `examples/journal-club.json` for a paper-review scaffold. These files contain demonstration content or missing-information labels, not medical evidence.

## Review the Actual PPTX

Read [references/quality.md](references/quality.md) before delivery. Run the structural audit, then render the exported PPTX with PowerPoint or LibreOffice when available and inspect each slide. An HTML/SVG source preview does not establish that the exported PowerPoint looks correct.

Fix clipped text by adjusting layout or splitting content before shrinking it excessively. The builder uses a configurable CJK font and conservative space estimates; those estimates are not a rendering guarantee. For PowerPoint/WPS interoperability, state exactly which application and version was tested.

The audit checks native object counts, expected text, embedded chart workbooks, table cells, out-of-bounds objects, and large picture surfaces. It cannot prove visual fidelity, chart semantics, medical accuracy, privacy, or the editability of every pixel inside an image. Do not call its result a universal editability percentage.

For revisions after manual editing in PowerPoint, treat the edited PPTX as the source. Do not regenerate from an older JSON and silently erase those changes. The bundled builder generates a new deck; it does not merge manual PowerPoint changes.

## Deliver

Provide the `.pptx`, authoring JSON or script, available source data, and audit report. Add previews when rendered. Summarize concrete editable object types, unresolved content, image-only exceptions, and rendering coverage. Keep identifiable patient data out of public repositories and shared examples.
