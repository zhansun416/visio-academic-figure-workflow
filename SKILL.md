---
name: visio-academic-figure-workflow
description: Reconstruct academic diagrams as editable Visio VSDX, including dense multi-panel figures, nested modules, repeated small elements, cross-links, and feedback loops. Use for reference-image reconstruction, native topology, SVG asset matching, and detailed visual verification against the source.
---

# Visio academic figure workflow

Deliver one editable `.vsdx`. Use PNG/SVG/PDF only for internal preview or diagnostics unless the user explicitly requests them.

## Non-negotiable output rules

- Treat the reference image as the layout authority; rebuild the page when an existing VSDX fights the reference.
- Fidelity includes element completeness, geometry, typography, local detail, and topology. Do not summarize a dense panel into a few boxes, collapse repeated cells/nodes, replace a chart with a generic icon, or omit small labels without the user's instruction. An editable rough draft is an intermediate result.
- Preserve source wording, capitalization, acronyms, line breaks, and notation. Use the requested typography; default to Times New Roman for English when unspecified. Do not rename labels to satisfy a case validator.
- Keep topology nodes, connectors, decorative icons, and labels editable. Use SVG only as a decorative icon; do not use an imported SVG as a logical connector endpoint.
- Use a native connection-capable Master for each logical node and a native `Dynamic connector` for each logical edge. Glue `BeginX` and `EndX` to explicit `Connections.Xn` cells. Never assume that `X1`–`X4` mean the same directions in a customer or custom Master; resolve cardinal points from the actual `Connections.Xn/Yn` coordinates.
- Use coordinate-drawn lines only for decorative separators or a documented fallback.
- Inspect an internal rendered preview before final delivery. A successful COM save is not visual proof.

## 1. Inspect the reference and choose the construction scale

For dense figures, nested groups, small repeated structures, or feedback about coarse results, read [references/dense-reconstruction.md](references/dense-reconstruction.md) **before drawing**. Make a region-by-region element ledger, then build and inspect the most detail-sensitive representative region before scaling to the whole page. The ledger must come from the reference, not from the generated spec.

Use the reference's aspect ratio and pixel coordinates. Preserve occupied area, panel proportions, alignment, whitespace, and visual density. Increase the internal working scale when necessary; do not remove content to fit a generic flowchart template. Ask for a clearer reference only when unreadable detail prevents an accurate reconstruction; continue work on readable regions.

## 2. Verify the required dependencies

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check_dependencies.ps1 -RequireVisio -SkipSvgLibrary
```

Initialize and synchronize the SVG library when the figure needs SVG assets; unrelated library inconsistencies must not block a native-only figure. The library defaults to `%USERPROFILE%\Documents\Codex\svg-library`; override with `VISIO_FIGURE_SVG_LIBRARY`. Read `references/dependencies.md` when installing elsewhere. Run `self_test.ps1` when changing the skill or diagnosing the environment, not before every figure.

## 3. Apply the SVG gate when icons are needed

Inspect customer assets first:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\prepare_svg_assets.ps1 `
  -SvgDirectory <customer-svg-directory> `
  -ManifestPath <work>\svg-inventory.json `
  -VisioSafeDirectory <work>\visio-safe-svg `
  -CompleteSet
```

Use this order while this Skill is active:

1. Reuse a complete, valid customer SVG set and skip all searching.
2. Reuse valid customer assets, then query the persistent library for explicit gaps:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\find_svg_library.ps1 -Query <keywords>
   ```

3. If gaps remain, use `find_icon_candidates.ps1` for a small Iconify candidate set.
4. Invoke `iconfont-svg-finder` only for unresolved icons or when the user explicitly requests it.

Save only approved assets to the persistent library through `add_svg_library.ps1`. Keep source, license, and hash metadata. Run `sync_svg_library.ps1 -Repair` only after reviewing unindexed files.

### Preserve SVG color and compatibility

- Preserve explicit fills, strokes, gradients, and layers.
- Replace `currentColor` only in a derived Visio-safe copy when needed; keep the source unchanged. `render_figure.ps1` creates a reusable hash-and-color-keyed copy under `%LOCALAPPDATA%\Codex\visio-academic-figure-workflow\svg-cache` automatically.
- Keep explicit multicolor fills, strokes, gradients, and layers unchanged; validate them through the rendered preview.
- Reject scripts, remote references, unresolved `<use>`, and SVGs with unsupported compatibility-risk elements until corrected or individually preflighted in Visio.
- Read `references/asset-licensing.md` before redistributing any library.

## 4. Rebuild with native topology

- Edit a supplied VSDX when its native geometry remains useful; otherwise build a new page.
- Use `assets/templates/figure_spec.json` with `scripts/render_figure.ps1` for repeatable geometry. It accepts relative paths, absolute SVG paths, and `library:<file.svg>` references.
- For dense figures, use `group` with local-coordinate `children`, named `ports`, connector `waypointsPx`, and independently editable labels. Read [references/figure-spec.md](references/figure-spec.md); `assets/templates/dense_figure_spec.json` is an executable example, not a visual template to impose on the source.
- Run `preflight_figure_spec.ps1 -SpecPath <work>\figure.json` before opening Visio. It checks nested ids, dimensions, endpoints, ports, and routes. Keep every reference-ledger item mapped to one or more spec ids.
- For custom figures, use a task-local `.ps1` builder with short phase markers. Keep shape helpers output-silent and emit one marker before each unfamiliar SVG import.
- Use native cards or nodes as connector anchors and place decorative SVG icons separately. Do not waste time hunting for a domain-specific stencil when a native topology node plus a decorative icon preserves both editability and visual fidelity.
- `render_figure.ps1` defaults connector endpoints to `auto`. It enumerates only real `Connections.Xn/Yn` rows, transforms points to page coordinates, and resolves left/right/top/bottom with a small edge tolerance. This remains correct for reordered, inset, rotated, and grouped shapes. Custom COM builders should dot-source `scripts/visio_connection_common.ps1` and call `Get-VisioAutoConnectionPair` instead of hard-coding `X1`–`X4`.
- Use optional `angleDeg` on native or SVG spec shapes. Keep explicit `Xn`, `Connections.Xn`, and cardinal endpoint overrides when a figure requires a deliberate route; auto-resolution runs only for endpoints set to `auto`.
- Route and review sequential edges, branches, feedback loops, output edges, arrow direction, and connection-point attachment.
- Use explicit routes for crowded cross-links and loops. Allocate separate lanes and ports where the source does; a crossing is not a junction. Preserve the source's layering. Avoid a global auto-layout pass after matching the reference.

Read `references/visio-compatibility.md` before writing or debugging a custom COM builder.

## 5. Validate and finalize

After visual QA, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\validate_scene_spec.ps1 `
  -VsdxPath <output>\figure-final.vsdx -SpecPath <work>\figure.json `
  -ReportPath <work>\scene-validation.json

powershell -ExecutionPolicy Bypass -File .\scripts\validate_vsdx_output.ps1 `
  -VsdxPath <output>\figure-final.vsdx `
  -OutputDirectory <output> -RequireSingleOutput `
  -RequireGluedConnectors -DisallowRasterMedia -MinimumFontSizePt 8 `
  -ReportPath <work>\final-validation.json
```

The scene validator checks every declared shape, group, label, edge target, and authored waypoint after saving and reopening. It detects omission or drift during rendering, but cannot detect details omitted from the spec. Compare **every reference region** against the preview, including small repeated elements; close the ledger only after that comparison. Inspect both a full-page view and readable detail crops. Fix omissions, clipping, overlaps, unglued endpoints, wrong arrows, font errors, and unintended raster media before delivery.

For custom builders, retain equivalent object ids and verify against their source ledger. Match `-ExpectedFont`, `-AllowedAcronyms`, and `-MinimumFontSizePt` to the requested source typography; the example's 8 pt is a publication check, not permission to enlarge labels and distort a dense layout. Use `finalize_output.ps1` only after validation; deliver one final `.vsdx` unless the user requested additional formats.

For diagrams whose connected node centers are intended to share an exact horizontal or vertical axis, also pass `-RequireAxisAlignedConnectors`. This catches a glued connector that is still visibly slanted because the wrong connection-point row was selected.

## Bundled resources

- `scripts/`: library synchronization, SVG inspection/search, Visio rendering, validation, self-test, and finalization.
- `assets/svg-library/`: local package cache; the public distribution should replace it with an empty bootstrap library.
- `assets/templates/figure_spec.json`: native-node and glued-connector example.
- `references/visio-compatibility.md`: native Masters, COM execution, SVG preflight, and fallback rules.
- `references/asset-licensing.md`: provenance and release auditing.
- `references/dependencies.md`: installation and persistent-library configuration.
- `references/dense-reconstruction.md`: source-led inventory, local reconstruction, and fidelity review.
- `references/figure-spec.md`: nested groups, ports, manual routing, labels, and coordinate conventions.
