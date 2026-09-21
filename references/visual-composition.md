# Visual Composition for Editable Medical Slides

Use this reference when authoring medical reports or improving a deck described as empty, repetitive, or text-heavy. These are composition decisions for the author, not new layouts supported automatically by `build_deck.py`.

## Choose the Visual Before Placing the Text

Identify the one relationship the audience should understand. Use a native diagram when geometry can explain it; use a sourced figure or replaceable photograph when the actual subject must be inspected.

| Content relationship | Useful editable composition | Evidence boundary |
| --- | --- | --- |
| Disease burden | Central disease node linked to skin, comorbidity, psychological and access dimensions | Equal node sizes do not mean equal prevalence or severity |
| Assessment | Central diagnostic question and surrounding assessment dimensions | A communication diagram is not a validated clinical algorithm |
| Local treatment | Simplified skin layers, application site and an action arrow | Label as a concept; do not imply anatomical scale or proven penetration depth |
| Phototherapy | Coverage or target-area schematic with treatment timeline | Device sketch is not a real device image; frequency and duration need sources |
| Targeted treatment | Short labeled mechanism chain with explicit inhibition marker | Do not invent molecular intermediates; identify simplification |
| Surgery or procedure | Candidate screening followed by donor-to-recipient workflow | Show selection requirements and complications as prominently as benefits |
| Treatment comparison | Native matrix or aligned benefit/risk statements around a diagram | No unsupported numeric scores, cross-trial rankings or fabricated rates |
| Collaboration or conclusion | Arrow spine, large step numbers, alternating callouts, destination circle | Use arrowheads for actual sequence or transfer; simple links for association |

Avoid a run of slides that only changes the heading above identical benefit/risk text columns. Retain common title, source and takeaway locations, but vary the central composition according to the content. A native chart remains a native chart with editable data; decorative bars must not replace it.

## Balance Density and Hierarchy

- Read a supplied reference as a visual system: alignment, rhythm, image weight, line width, accent use and size contrast. A faint full-slide background photograph does not by itself make a text page visually rich.
- For chapter pages, use a strong title, a relevant graphic on the opposite side and large numbered steps on a connected baseline. For a 16:9 deck, starting sizes around 38-44 pt for the title, 28-34 pt for numbers, and 24-28 pt for step labels are useful; verify the actual longest label. They are starting points, not fixed requirements.
- Make accidental asymmetry visible during review: mentally divide the usable slide into quadrants. A large empty upper-right region combined with small text elsewhere usually needs redistribution, a meaningful diagram, or a stronger image crop.
- Keep deliberate breathing room around the focal point and between reading groups. Do not try to fill every pixel or impose a fixed coverage percentage across all audiences.
- When the user requests a fuller composition, enlarge useful labels, shorten and regroup prose, and allocate substantial space to the main diagram. Do not add paragraphs solely to occupy space.
- For benefit/risk pages, keep both sides readable and comparable. Diagrams must not push contraindications or major warnings into tiny footnotes. Use a small number of concise visible points; put supporting explanation in notes.
- Use circles for entities or milestones, arrows for direction, and lines for association. Random circles, arrows or medical symbols create clutter without improving comprehension.

## Medical Images and Editability

Use actual authorized clinical images only when they add clinical value. Record source, permission and caption; preserve diagnostic content and aspect ratio. Do not use generated clinical photographs or synthetic before/after images as evidence. A simplified native anatomy or device sketch should be labeled as conceptual and should not imply observed patient findings.

Build simple diagrams from native PowerPoint shapes, connectors and editable text. Keep image objects separate and replaceable. Do not flatten the slide or add an invisible text overlay and describe the result as editable. Preserve user-template assets privately; exclude them from a public skill package.

Give every authored object a nonempty name. A point or unlabeled circle does not need a separate empty text box; do not overwrite the library's default object name with an empty label. An actual PowerPoint opening test caught a repair prompt that the bounds audit missed in a custom diagram draft.

## Visual Revision Acceptance

1. List the affected pages and the composition problem each revision addresses.
2. Render the final PPTX in the target application. Check the actual output, not only a source preview.
3. At full-slide size, inspect labels, arrow destinations, body wrapping, source lines and footer separation. Text boxes inside slide bounds can still overflow visually.
4. In a contact sheet, check reading rhythm and whether treatment slides have become a repeated template with interchangeable text.
5. Check that the main figure is legible, each visual has a purpose, risks remain visible, and no graphic implies invented quantities or clinical facts.
6. Keep the validated data chart, table and evidence notes intact. Record what was rendered and what editing or application compatibility checks were actually performed.

Keep connector segments out of label and paragraph areas, including chapter baselines that would otherwise run through step titles. Check the final line of each benefit group against the following risk heading; different paragraph spacing can erase a gap that seemed sufficient from coordinates alone. Re-render after fixing either problem.
