# Editable Follow-up Charts

Use a native PowerPoint chart with an embedded workbook when the user needs to change data later. Bind data labels to the numeric values; do not replace them with detached text boxes or flatten the chart to an image.

## Data and Units

- Preserve the supplied observations and time points. Do not invent extra visits, scores, treatments or patient characteristics from a reference image.
- Choose and document one percentage convention. If a mixed-unit chart stores percentage points, enter `12` for `12%` and use the literal-suffix format `0"%"`. If the workbook stores fractions, enter `0.12` and use `0%`. Applying `0%` to `12` produces `1200%`.
- Put each unit in the legend or axis labels and in the workbook headings. Sharing an axis does not make different clinical scales directly comparable.
- A category axis spaces visits equally. Use a date or numeric axis when elapsed-time spacing matters, or state the equal-visit-spacing convention in notes.
- When pages use different axis limits, make each scale visible. Do not compare their slopes directly. Use common limits when the requested purpose is comparison between pages.

## Labels, Markers and Layout

- Keep equal-valued observations at the same coordinates. A hollow outer marker and a smaller filled marker can reveal coincident points without changing the data.
- Place labels around points using above, below, left or right positions. Check the entire adjacent line segment: a label can avoid its own point while still intersecting another series.
- With `python-pptx`, explicitly set font family, size, weight and color on individual point-label overrides. In PowerPoint, newly created overrides may use default black text instead of the series styling.
- Confirm the percentage format also reaches point-specific labels. Check the first and last labels for clipping and label-to-axis collisions.
- A corner legend may shrink the plot if it reserves layout space. In `python-pptx 1.0.2`, `legend.include_in_layout` writes directly to the OOXML `c:overlay` value: `True` permits an overlay. Verify the actual result in PowerPoint and keep the legend clear of the lines.
- Match the reference's color hierarchy, whitespace and line weights without copying unrelated branding. Native labels, markers, gridlines and titles should remain editable.

## Appending and Preserving Slides

Work on a copy of the latest user-edited presentation. Each appended chart needs its own chart part and embedded workbook; sharing a workbook relationship can cause later data changes to affect another chart.

Reuse styles while keeping the new chart's workbook relationship. Do not attach a slide layout from another presentation package without properly importing its master and dependencies: duplicate ZIP part names can produce invalid files. Check that existing slides, chart data and notes remain intact after appending.

## Verify and Deliver

Run the structural audit with a manifest and compare the supplied numbers with both chart caches and embedded workbooks. Render the actual PPTX and inspect overlapping points, labels, units and fonts. A successful library reload does not prove that PowerPoint's interactive Edit Data command works.

If the user stops desktop control, stop immediately. Complete any authorized background checks and state which visual or editing checks remain unperformed. Tell the user how to use Edit Data and explain the percentage entry convention used in that file.
