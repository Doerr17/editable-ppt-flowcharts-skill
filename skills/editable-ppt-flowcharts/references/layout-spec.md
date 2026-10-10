# Layout specification

The user supplies prose. The agent converts that prose into this small JSON format; do not require the user to write JSON. This format controls drawing, not scientific interpretation or automatic natural-language parsing.

Run `python scripts/build_ppt.py layout.json output.pptx` from the skill directory (or use absolute script paths).

```json
{
  "font": "Arial",
  "east_asian_font": "Microsoft YaHei",
  "slides": [{
    "title": "Feature Extraction",
    "subtitle": "From validated input to features",
    "notes": "Preserve full source conditions here.",
    "nodes": [
      {"id": "input", "rect": [0.5, 1.5, 2.6, 1.1], "title": "Input", "body": ["Paired files"], "accent": "gray"},
      {"id": "qc", "rect": [3.7, 1.5, 3.1, 1.1], "title": "Validate", "body": ["Check IDs and files"], "accent": "blue"},
      {"id": "caption", "kind": "text", "rect": [0.5, 3.0, 10.0, 0.4], "title": "Independent routes", "font_size": 19, "accent": "blue", "align": "left"}
    ],
    "edges": [
      {"from": "input", "to": "qc", "points": [[3.2, 2.05], [3.6, 2.05]], "style": "block"}
    ]
  }]
}
```

## Fields

- Canvas defaults to 13.333333 × 7.5 inches (16:9). Optional `canvas: [width, height]` preserves a supplied template's aspect ratio. All coordinates use inches.
- Root `font` defaults to Arial. `east_asian_font` applies to CJK runs and defaults to Microsoft YaHei; use the supplied template's actual fonts when available.
- Each slide requires `title` and `nodes`. Optional `subtitle`, `notes`, and `footer` are plain strings. Header and page number are drawn automatically.
- Each node requires a unique slide-local `id` and `rect: [x, y, width, height]`. Default `kind: "box"` draws an editable rounded rectangle; `kind: "text"` draws an editable text box. `title` is a string; `body` is a list of strings, one paragraph per entry. Use short lines and explicit paragraph breaks rather than font shrinking.
- `accent` is one of `blue`, `green`, `orange`, `purple`, `gray`. Optional `font_size` and `body_font_size` use points. Optional `align` is `left`, `center`, or `right` (boxes default to center; text defaults to left). Optional `bold` controls text-node emphasis.
- Each edge requires existing `from` and `to` node IDs, and at least two `points: [[x, y], ...]`. Points are the visible routed path; intermediate points make editable horizontal/vertical connector segments. Arrowheads appear only on the last segment. `arrow` defaults to true, `dashed` to false, and `accent` to blue.
- `style: "line"` is the default. `style: "block"` draws the short solid right-facing arrow used between adjacent stage boxes; it requires exactly two horizontal points increasing in x.
- A branch or merge is represented by separate source→target edges that share parts of the routed path. Repeat shared segments only when unavoidable. Do not draw an arrowhead on a shared bus.

## Layout decisions

- Start with 0.5-inch side margins. Header: title at y≈0.25, 32 pt; subtitle at y≈0.88, 16.5 pt. Diagram body starts near y=1.4 and ends by y=7.15; reserve the bottom-right corner for a 9.5-pt page number.
- A short sequence uses 3–5 stage boxes across the first row. With five boxes, widths of 1.8–2.4 inches are practical; leave visible space for arrows. Long labels need wider boxes, shorter wording, or another row.
- After preprocessing, route one bus to 2–4 class-specific or independent method boxes. Use 0.25–0.35 inch vertical gaps for connector turns. Auxiliary-only processing gets a gray box and a visibly separate output.
- For a long exclusion chain, show compact reference boxes on a separate row, with the pair-level or other acceptance rule directly below the chain. Route the retained-input node to independent branches below it.
- Prefer box headings 17–20 pt and body 15–17 pt. Short technical labels may use 14–14.5 pt. If important content needs smaller text, shorten or split the slide. These sizes are defaults, not a reason to alter the user's scientific conditions.
- Keep dotted/dashed connectors for auxiliary or informational contribution only. Solid arrows communicate data flow. A shared mapping/QC node must say "each route" when the two results are processed separately.
- A connector should terminate at a node edge, not cross text or run through a box. The generator draws the supplied route; it does not solve collisions or infer branch meaning.

The generator validates node IDs, connector references, and canvas bounds. With `--spec`, the PPT validator checks explicit run fonts, edge segment coordinates, and arrow direction. It uses the helper's shape names: `node:{id}` and `edge:{from}->{to}:{edge_number}:{segment_number}` (numbers start at 1). If adapting the helper, retain this naming and explicitly set both Latin and East Asian run fonts. The checker does not resolve inherited theme fonts or infer flow semantics. PowerPoint rendering remains necessary to detect actual text wrapping and text overflow.
