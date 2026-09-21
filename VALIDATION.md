# v0.1.1 Validation Record

## Release Checks — 2026-09-21

- All 8 existing behavior tests passed in the release copy.
- The bundled skill-creator `quick_validate.py` passed with PyYAML 6.0.3 installed in a separate local tooling directory. PyYAML is not required by the skill's PPTX generator.
- The bundled synthetic sample passed its manifest audit: 7 slides, 1 native table, 1 native chart and 1 embedded workbook; 0 errors and 0 warnings.
- Code and sample PPTX are unchanged from the previously validated generator and rendered sample. This release adds authoring guidance and aligned release documentation, plus a GitHub Actions workflow to run the existing tests and sample audit. The hosted workflow has not yet run at preparation time.
- PowerPoint rendering coverage remains the dated sample inspection below; no new Office rendering or interactive editing test is claimed for this publication.

## Original Sample Validation

Date: 2026-09-07.

## Automated Behavior Checks

Runtime: Python with `python-pptx 1.0.2` and Pillow `12.3.0` on Windows.

Command: `python -m unittest discover -s tests -v`.

Result: 8 tests passed after the final generator revision.

- Native table cells, chart data, embedded workbook, and missing-value gaps survive generation and reload.
- Title text, a table cell, and a chart series can be changed through `python-pptx`, saved, and reopened; the old manifest detects the changes.
- Invalid references, mismatched data, booleans/non-finite chart values, unknown fields, and duplicate IDs are rejected.
- Estimated text overflow and accidental output overwrite are rejected.
- A PNG is stored as a separate picture with unchanged source bytes, aspect ratio, and alternative text.
- Negative bar data and missing values are retained.
- Out-of-bounds shapes are reported as errors; large image surfaces generate review warnings.
- The journal-club scaffold generates successfully.

Sample audit: 7 slides, 1 native table, 1 native chart with 1 embedded workbook; 0 structural errors and 0 warnings. See [case-demo.audit.json](examples/case-demo.audit.json).

## Office Rendering

Application: Microsoft PowerPoint for Windows, executable version `16.0.19127.20302`.

The final generated PPTX opened without a repair prompt and all 7 slides exported successfully through PowerPoint's PNG export. Images are 1280 x 720. The chart's East Asian fonts were made explicit and its default gridlines and markers adjusted after the first visual inspection.

Preview contact sheet: [preview.png](examples/preview.png). Individual exported slides are in `examples/slides/`.

All 7 final slide previews were visually inspected. No obvious clipping or overlapping text was observed in this sample. This is a visual inspection, not a universal font-fit guarantee.

PowerPoint background COM export was unavailable in the task's Windows logon session; the actual export used the desktop application. No successful COM geometry measurement is claimed.

## Boundaries of This Record

- The editing round trip was verified with the library, not a full manual chart-data editing session in PowerPoint.
- WPS, LibreOffice, Keynote, macOS, and alternative fonts were not tested.
- The visual sample is one synthetic case, not a benchmark for all medical content or layouts.
- No real clinical claims or patient data were validated.
- On 2026-09-07, the bundled skill-creator validator could not run because PyYAML was absent. The 2026-09-21 release check above supersedes that earlier format-validation limitation.

## Reproduction

```sh
python scripts/build_deck.py examples/case-demo.json --output build/reproduction.pptx
python scripts/audit_pptx.py build/reproduction.pptx --manifest build/reproduction.manifest.json --output build/reproduction.audit.json
python -m unittest discover -s tests -v
```

Use a new output name for each build. Render that actual PPTX in the intended office application and inspect all slides before using real content.
