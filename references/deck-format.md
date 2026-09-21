# Deck JSON v0.1

UTF-8 JSON is parsed with the standard JSON parser. Paths resolve relative to the JSON file. Unknown fields fail validation so misspelled data fields do not disappear silently. See the examples for complete working inputs.

Root fields:

| Field | Meaning |
| --- | --- |
| `title` | Presentation title; required |
| `subtitle` | Optional short description |
| `demo` | Boolean; if true every slide is labeled synthetic |
| `font` | Installed font family, default `Microsoft YaHei` |
| `sources` | Mapping from source ID to full citation/provenance string; required |
| `slides` | Nonempty array, maximum 100 slides |

Each slide needs `id`, `layout`, and `title`. IDs must be unique, using ASCII letters, digits, hyphens or underscores. Optional common fields are `subtitle`, `notes`, and `source_ids`. Source IDs must exist at the root. The full source strings are placed in notes; short source IDs are shown at the bottom of the slide. Add an explicit reference slide using a table or bullets for an audience-readable bibliography.

| Layout | Content fields | Practical bounds |
| --- | --- | --- |
| `cover` | `body` string | A short introduction |
| `bullets` | `items`: list of strings | 1-6 concise points |
| `timeline` | `events`: list of `{label, text}` | 2-5 events |
| `table` | `columns`: strings; `rows`: equally sized rows of strings | 2-5 columns, 1-8 body rows |
| `chart` | `chart_type`: `line` or `bar`; `categories`: strings; `series`: `{name, values}`; optional `y_title`, `takeaway` | 2-12 categories, 1-4 series; finite numbers or `null` for missing values |
| `image` | `image`: local PNG/JPEG path; `caption`: string; `alt`: string | Image is proportionally contained, not cropped |

All fields except `chart_type`, `y_title`, and `takeaway` listed in the selected layout are required. Chart type defaults to `line`. Each series has the same number of values as categories and at least one finite value. `null` is displayed as a gap rather than zero. Percentages must use one convention consistently and include the unit in `y_title`.

Charts and timelines use equally spaced categories/events. They are not numeric time axes: the sample's week 0/2/4/8 labels indicate observation occasions, not proportional spacing. If numeric time spacing is material to the scientific argument, use an appropriate scatter/time-axis implementation instead of these bundled layouts and validate the resulting data.

Texts are plain strings; Markdown is not rendered. The builder raises an error for content that clearly exceeds its layout estimate rather than truncating it. Revise or split that slide. Stable slide IDs and named shapes link the manifest to PowerPoint's Selection Pane. The manifest stores exact expected text fragments and native chart/table data for structural verification.

The builder does not import arbitrary templates, extract PDFs, insert equations, or update individual slides in an existing PPTX. The agent can read source documents with available tools before creating the JSON. Do not pass a PPTX to the JSON builder or promise template fidelity from it.
