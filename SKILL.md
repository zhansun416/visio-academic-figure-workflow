---
name: visio-academic-figure-workflow
description: Reconstruct academic diagrams as editable Visio VSDX with Python, including dense multi-panel figures, nested modules, rich SVG icons, repeated small elements, cross-links, and feedback loops. Use for reference-image reconstruction, native topology, SVG asset matching, and source-led visual verification.
---

# Visio academic figure workflow

Deliver an editable `.vsdx`. Keep specs, ledgers, previews and comparison reports in the working directory unless requested. Run Python commands from this skill's directory with absolute paths for task inputs and outputs.

## Fidelity and editability

- The source governs layout, wording, capitalization, notation, colors and density. An editable rough diagram is intermediate. Never collapse repeated objects, omit small labels or substitute a generic icon for a chart.
- Preserve source typography. Times New Roman is a fallback, not a rule to override a visible source font. Scale page and font sizes together; do not enlarge tiny labels to a fixed minimum and distort the source.
- Text, cells, charts, circuits, node groups and glued connectors should be native. Decorative SVGs are supported. Inspect the saved package to establish whether Visio imported each as native geometry or retained a foreign picture. A picture is not a native shape; a native group alone does not prove faithful geometry.
- Logical nodes use native connection-capable Masters; edges use native Dynamic connectors. Attach icon-bearing nodes to native frames, not imported SVG internals. Native lines/paths may represent circuit primitives and decorative strokes.
- Reopen, validate, render and inspect every critical region. COM success and object counts cannot prove visual fidelity.

## 1. Inventory the original source

Read [dense-reconstruction.md](references/dense-reconstruction.md) for dense sources or feedback about coarse results. Before building the spec, inventory each source region: labels, repeated counts, icons and visual features, symbols, topology and uncertain details. Use [reconstruction-manifest.md](references/reconstruction-manifest.md) for the source ledger and draft/delivery gates.

Use source pixels and aspect ratio. Inspect full resolution and readable crops. Build the most detail-sensitive region first and compare before repeating motifs. Unreadable details stay explicitly unresolved; do not invent exact measurements from blurred marks.

## 2. Check Python and Visio

```text
python -m pip install -r requirements.txt
python scripts/visio_workflow.py check-env --probe-visio
```

Python 3.10+ is required. Rendering/COM validation need Windows desktop Visio and pywin32. Preflight, SVG library and package/manifest inspection work without Visio; comparison needs Pillow and full-page PNG previews need pypdfium2. See [dependencies.md](references/dependencies.md). No command delegates to PowerShell.

## 3. Match a sufficiently rich SVG set

Inventory distinct icon families first. Record each source crop, candidate, match observations and unresolved differences. Judge silhouette, internal detail, stroke/fill, colors and optical size. Keyword matches are insufficient. Do not reuse a generic symbol for unrelated source objects.

1. Inspect supplied SVGs first. A complete usable package skips external searches.
2. Query the persistent library. Unknown-license candidates are not permission to redistribute.
3. Search missing families with `iconfont-svg-finder` when applicable or explicit candidate queries. Accurate hand-authored SVG details are acceptable; record source provenance and uncertainty.
4. Render representative multi-path, stroked, nested, multicolor and small-scale assets in Visio. Inspect icon interiors, not only outer bounds. Thin strokes and empty interiors must survive import/resizing.

```text
python scripts/visio_workflow.py svg init
python scripts/visio_workflow.py svg sync
python scripts/visio_workflow.py svg find "database shield solar wind"
python scripts/visio_workflow.py svg prepare <assets> --safe-dir <work>/visio-safe --complete
python scripts/visio_workflow.py svg candidates "solar panel" <work>/candidates --download 3
python scripts/visio_workflow.py svg register <assets> --source-label <source> --source-url <url> --license <license> --tags <tags>
```

Preserve explicit multicolor values. Resolve `currentColor` in a derived copy; keep originals intact. Unsupported effects and risky references fail preflight; do not silently rasterize. Review warnings before declaring completeness. See [asset-licensing.md](references/asset-licensing.md).

## 4. Build the native scene

Read [figure-spec.md](references/figure-spec.md). Templates demonstrate structure, not layouts to impose on the source. Use task-local Python loops to expand repeated motifs into individual editable objects. Native groups, paths, ports, route bends, self-loops and separate labels are supported.

```text
python scripts/visio_workflow.py --report <work>/preflight.json preflight <work>/figure.json
python scripts/visio_workflow.py --report <work>/render.json render <work>/figure.json <output>/figure.vsdx --preview <work>/preview.png
```

Use native frames with icons, labels and details as children. Reserve separate lanes/ports for cross-links and loops. Crossings are not junctions. Preserve source layering and avoid global auto-layout. See [visio-compatibility.md](references/visio-compatibility.md) for custom native builders and SVG stroke scaling.

## 5. Inspect actual objects and compare images

```text
python scripts/visio_workflow.py --report <work>/validation.json validate <output>/figure.vsdx --spec <work>/figure.json --require-glued --no-raster
python scripts/visio_workflow.py --report <work>/objects.json inspect <output>/figure.vsdx
python scripts/visio_workflow.py preview <output>/figure.vsdx <work>/preview.png
python scripts/compare_renders.py <source> <work>/preview.png <work>/comparison --manifest <work>/source-manifest.json
```

Map inventoried source components to **actual saved names/types**, including wrappers and imported children. Keep the source inventory independent; never generate it solely from output. Compare full-page and region side-by-side, overlay and difference images. Aspect mismatches fail instead of stretching. Metrics are diagnostic, not pass thresholds.

Inspect each critical region and then the whole source again. Correct omissions, broken SVG strokes, substituted icons, clipping, density, fonts and arrows. Recheck saved/reopened output and readable crops of every rich SVG family. Incomplete/unreviewed content blocks delivery. Record specific evidence and limitations rather than automatically marking reviews verified.

```text
python scripts/visio_workflow.py --report <work>/delivery-gate.json inspect <output>/figure.vsdx --manifest <work>/source-manifest.json --mode delivery
```

Delivery rejects incomplete/deferred content, undeclared editability changes, missing/type-mismatched mappings, unmapped objects and unverified source reviews. It cannot establish the truth of notes or detect content never inventoried.

Use `--expected-font`, `--min-font` and `--allowed-acronyms` only when grounded in the source/user requirements. `--axis-aligned` applies only when node centers should share axes. `finalize <vsdx> <archive>` moves extra top-level files outside a dedicated delivery directory; use after visual QA.

## Regression checks

```text
python -m unittest discover -s tests -v
python scripts/self_test.py --output <work>/regression --visio
```

Synthetic regressions prove capabilities, not source fidelity. User images and derived private assets remain local unless publication is authorized.
