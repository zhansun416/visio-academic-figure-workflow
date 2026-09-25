# Dense reference reconstruction

Use this mode when detail, not just the high-level concept, must match a supplied figure. Many objects do not justify simplifying their content. The renderer only reproduces what the author puts into the spec; improving the renderer alone cannot repair an incomplete interpretation of the reference.

## Inspect and account for the source

Inspect the full-resolution source and readable crops of each panel. Record the page dimensions, panel bounding boxes, containment, alignment guides, color palette, and local spacing. Keep one coordinate system tied to the reference pixels.

Make a task-local ledger with these fields:

| Field | Purpose |
| --- | --- |
| source region and bbox | Which part of the actual reference is being accounted for |
| element inventory | Exact labels, repeated element counts, borders, icons, chart marks, legends, annotation lines |
| topology | Endpoint objects, branches, arrow directions, crossings versus junctions, feedback lanes |
| implementation ids | Native objects or grouped components that reproduce each source item |
| unresolved detail | Unreadable text or uncertain geometry, without invented content |
| visual status | Unbuilt, built, compared, or needs correction; close only after inspecting the rendered region |

Count repeated cells, layers, bars, ports, and small network nodes from the source. Do not silently replace twelve visible objects with three representative ones or use ellipses unless the reference does. Transcribe text before geometry so labels are not lost while focusing on layout. Include math symbols, subscripts, superscripts, legends, and panel identifiers in the inventory.

## Build one representative region, then scale

Choose the region that tests the hardest detail: for example a nested module with repeated small elements and a cross-boundary connector. Render it and compare it with the corresponding reference crop. Correct the primitive, spacing, and font choices before repeating them throughout the figure. Continue building autonomously; this is an internal quality check, not an approval checkpoint.

Use task-local loops to expand repeated motifs into explicit editable objects. A matrix should retain its cells and values; a mini-chart should retain axes, ticks, bars/curves, and labels; a layered network should retain its visible layers and links. The loops are a construction convenience, not permission to reduce the number of objects.

Choose native geometry appropriate to each object. Composite groups are valid for structures that cannot be represented by a single master. Attach edges to a native frame or native child with connection points. Decorative SVGs may be reused where their appearance matches; they must not replace a complete data panel or serve as logical connector endpoints. If a required primitive exceeds the JSON renderer's supported kinds, extend a task-local native COM builder using the compatibility notes rather than replacing it with a generic box. Preserve editable text separately from graphics.

## Resolve dense layout deliberately

- Work from panel boundaries to inner regions to details. Child bboxes are relative to their immediate group; do not mix local and page coordinates.
- Reserve annotation space and routing lanes before completing interiors. Multiple edges can use different normalized ports on the same node.
- Trace a long connection from source to target on the reference. Store its intermediate points explicitly when automatic routing changes the intended route. Distinguish true junctions from visual crossings using native junction nodes where needed.
- Keep visible text above the objects it labels. Use intentional label backgrounds only when the reference has them; do not hide mistakes with opaque white patches.
- Match text boxes, margins, line breaks, alignment, weight, and size. Do not shrink everything until overlaps disappear. Recheck measured geometry and box sizes first.
- Keep repeated motifs geometrically consistent while retaining exceptions in the source. A single master or helper can improve consistency without losing individual editability.

## Review fidelity in two passes

First inspect the full page for aspect ratio, panel proportions, hierarchy, occupied area, density, and major connection paths. Then compare every region at a readable scale for individual elements, line breaks, small symbols, thickness, colors, spacing, clipping, and connector attachment. Side-by-side registered crops or a diagnostic overlay can help; do not stretch either image to hide a mismatch.

Inspect the rendered VSDX after the final save and reopen. Grouping and rerouting can change geometry that looked correct during construction. The bundled `validate_scene_spec.ps1` catches source-spec-to-file drift, including missing objects, wrong endpoints, misplaced manual bends, and text changes. It does not claim image similarity, detect all text overflow, or prove that the source inventory was complete.

Use the source ledger as the completion gate: all readable source items mapped and visually compared; all logical edges accounted for; no unexplained omissions or substitutions. If a supplied image is too blurred to resolve an item, record and explain that specific limitation rather than guessing. Without a real reference, a synthetic regression figure proves capabilities only and must not be described as proof of fidelity to the user's failed example.
