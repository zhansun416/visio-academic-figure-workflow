# Figure spec: dense scenes

The existing flat `figure_spec.json` remains supported. The renderer uses pixels in the reference image and inches on the Visio page. Use matching aspect ratios unless the user requests a different layout. The dense example contains nested groups, a matrix, native chart details, ports, cross-panel routes, labels, and a self-loop.

## Shapes and coordinates

`shapes` contains objects with a globally unique `id`, `kind`, and `bboxPx: [x,y,width,height]`. Coordinates start at the upper-left. Supported kinds are `rect`, `oval`, `circle`, `diamond`, `native`, `text`, `line`, `svg-asset`, and `group`.

A `group` is a native frame plus a real Visio group containing its `children`. Child bboxes use the immediate parent's upper-left as the origin; they are translated, not scaled. Nested groups use the same rule. A connector targeting a group id glues to its native frame. Internal connectors and labels are included in their endpoints' lowest common group; cross-group connectors remain in their common ancestor or on the page. This preserves individual editability and grouped movement. Groups cannot set `angleDeg`; rotate individual native/SVG children if needed.

```json
{
  "id": "module", "kind": "group", "bboxPx": [200,100,400,240],
  "fill": "#F5F7FA", "line": "#91A6B4",
  "children": [
    {"id":"module-title", "kind":"text", "bboxPx":[15,10,360,30], "text":"Controller", "align":"left"},
    {"id":"solver", "kind":"rect", "bboxPx":[100,90,180,70], "text":"Solve",
     "ports":[{"id":"in1","x":0,"y":0.3},{"id":"in2","x":0,"y":0.7}]}
  ]
}
```

The `solver` above occupies page pixels `[300,190,180,70]`. Put group titles in separate text children to control title position. Drawing order follows the declared objects, with a group's frame behind its children. Cross-container connectors are raised after grouping so opaque frames do not erase their visible segments. Choose routes that avoid covering unrelated text and nodes.

Text supports `text`, `font`, `fontSizePt`, `fontColor`, `bold`, `italic`, `align` (`left`, `center`, `right`), `verticalAlign` (`top`, `middle`, `bottom`), and `textMarginPt`. Defaults are Times New Roman, 10 pt, centered, middle, and 2 pt margins. Fill/line/style fields work for text boxes too, including explicit label backgrounds. Use real JSON `\n` for line breaks. Rich text runs, equations, curved paths, and arbitrary polygons need a task-local native builder; do not approximate them silently.

`line` allows zero width or zero height for decorative axes and separators, but not both. It draws from the lower-left to the upper-right of its bbox. Logical connections belong in `connectors`.

## Ports and connectors

`ports` add native connection rows. Port `x,y` are normalized to the unrotated shape's upper-left: `(0,0)` top-left, `(1,1)` bottom-right. They scale and rotate with the node. Port ids start with a letter and contain letters, digits, `_`, or `-`.

Endpoint values may be `auto`, `left`, `right`, `top`, `bottom`, `Xn`, `Connections.Xn`, or `port:<id>`. Use explicit endpoints for feedback loops and crowded boundaries. Native masters may have different cardinal row orders; never infer a direction from the row number.

```json
{
  "id":"feedback", "from":"decision", "to":"solver",
  "fromConnection":"left", "toConnection":"port:in2",
  "routing":"manual", "waypointsPx":[[260,480],[260,239]],
  "arrow":"end", "line":"#AE5A36", "linePattern":2,
  "label":{"text":"Refine", "bboxPx":[150,448,90,25], "fontSizePt":10}
}
```

`waypointsPx` and label bboxes always use **page/reference coordinates**, even when endpoints are nested. Waypoints are intermediate bends only; glued endpoints supply the first and last vertices. For an orthogonal path, align each successive point with its neighbor and the chosen ports. The renderer does not find obstacle-free lanes for you. Self-loops require explicit endpoints and waypoints. Repeated edges should have separate ports or lanes if the reference distinguishes them.

`routing` supports `auto` (legacy default), `orthogonal`, `straight`, or `manual`. Waypoints imply manual routing and cannot be combined with `orthogonal` or `straight`. The manual path remains a native connector with native geometry and glued ends. Its bends are authored geometry, not an automatic obstacle-avoidance route: after moving individual nodes, review the route. Moving its containing group carries internal bends with it.

`arrow`: `none` (default), `begin`, `end`, `both`. `lineJump`: `default`, `never` (default), `always`, `other`, `neither`. Jump behavior is Visio's routing behavior and must be visually checked; it is not evidence of a logical junction. Manual geometry has no guaranteed automatic jump rendering: draw the required crossing explicitly in a custom builder if a source-specific bridge cannot be represented faithfully.

## Validation

Run `preflight_figure_spec.ps1` before rendering. It validates unique ids across shapes and edges, native endpoint references, group expansion, numeric bounds, ports, and route data without Visio. Warnings identify children extending outside their group for review. It does not perform OCR or infer missing source content.

The renderer stores ids in `User.SpecId`; `::group` and `::label` suffixes are reserved. Run `validate_scene_spec.ps1 -SpecPath ... -VsdxPath ...` after saving: it reopens the VSDX and checks object presence, nesting, text, center/size, edge targets, arrows, label positions, and explicit bends with a default tolerance of 1 reference pixel. Follow this with output validation and region-by-region visual comparison.

## Visio API references

- [Selection and native grouping](https://learn.microsoft.com/en-us/office/vba/api/visio.page.createselection)
- [Page-to-shape coordinate transformation](https://learn.microsoft.com/en-us/office/vba/api/visio.shape.xyfrompage)
- [Connector routing modes](https://learn.microsoft.com/en-us/office/client-developer/visio/shaperoutestyle-cell-shape-layout-section)
- [Reroute control](https://learn.microsoft.com/en-us/office/client-developer/visio/confixedcode-cell-shape-layout-section)
