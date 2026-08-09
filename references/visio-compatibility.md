# Visio compatibility notes

These notes are based on the Windows Visio COM workflow used by this package.

## Tested constraints

- `DrawDiamond` and `DrawRoundedRectangle` are not available on some Visio COM surfaces used by the workflow. The standard `BASIC_M.VSSX` stencil can provide native `Diamond`, `Circle`, `Ellipse`, and other masters through `Page.Drop`; verify the stencil path and master before relying on it.
- `BASIC_M.VSSX` provides tested native geometry masters such as `Rectangle`, `Diamond`, `Circle`, and `Ellipse`. `SSFLOW_M.VSSX` provides a tested `Dynamic connector` master. Resolve the localized stencil path at runtime instead of hard-coding a machine-specific path.
- Writing unsupported `TxtMarginLeft` or `TxtMarginRight` cells can raise a misleading “unexpected end of file” error.
- `Page.Import` accepts local SVG files, but imported SVGs should be treated as graphic objects rather than native Visio connection-point shapes.
- A COM save succeeding does not prove that the page renders correctly. Always render and inspect a preview.

## Safe construction pattern

- Insert connector-bearing cards and nodes from a verified stencil master with `Page.Drop`; use `DrawRectangle` only for non-connectable decorative panels or when the stencil fallback is documented.
- Apply this to every semantic object that participates in the topology, including arrows, process symbols, network symbols, containers, and callouts. Prefer a single connection-capable native Master over a visual composite of lines, rectangles, or SVG fragments.
- Prefer native stencil masters through `Page.Drop` for diamonds, circles, ellipses, and other standard geometry. Use four line segments only as a documented fallback when the verified stencil/master is unavailable, and place decision text as a separate text object when it improves editability.
- Keep SVG, text, card, and connector objects separate so a missing import cannot hide text.
- Drop a `Dynamic connector`, then glue its `BeginX` and `EndX` cells to source/target `Connections.Xn` cells with `GlueTo`; set arrowheads and routing after glue. Use explicit line segments only for decorative separators or a documented fallback.
- In `scripts/render_figure.ps1`, `rect`, `oval`, `circle`, `diamond`, and `native` kinds use stencil Masters. Native and SVG shapes may set `angleDeg`. Connector specs may set `fromConnection` and `toConnection` to `auto`, a cardinal side, `X1`, or `Connections.X3`.
- For the tested bundled basic-shape masters, `X1` is bottom, `X2` is right, `X3` is top, and `X4` is left. This ordering is not portable: customer VSDX files and custom Masters may add a center point, use `X5` for right, or assign a different order. Use `CellExistsU` before reading a row because `CellsU` may return an empty cell for a missing row instead of throwing. Transform every local `Connections.Xn/Yn` point with `Shape.XYToPage(x, y, xPrime, yPrime)` before classifying page-left/right/top/bottom. `visio_connection_common.ps1` implements these rules and uses a 0.001-inch edge tolerance so a nearly equal corner does not outrank a centered side point.
- The renderer computes an automatic pair only when at least one endpoint is `auto`. Explicit `Xn`, `Connections.Xn`, and cardinal overrides remain available and are validated with `CellExistsU` before glue.
- Keep branch labels separate from connector geometry and inspect `Yes`/`No` directions after rendering.

## Output validation

- Use `validate_vsdx_output.ps1 -RequireGluedConnectors` to require both endpoints of every native dynamic connector to be glued.
- Add `-RequireAxisAlignedConnectors` for orthogonal node layouts. It recursively checks native connectors, compares target centers in page coordinates, and rejects any remaining endpoint-axis deviation above `-AxisAlignmentToleranceIn`.
- Use `-DisallowRasterMedia` for an SVG/native-vector-only deliverable.
- Use `-MinimumFontSizePt 8` for figures intended for papers, then visually inspect the preview at the expected reduced publication size.

## Large-builder execution

- Do not send a large builder as one long `powershell -Command` payload. Run a task-local `.ps1` file with explicit phase markers and a bounded session wait.
- Keep shape helpers output-silent; accidental COM shape objects in the PowerShell pipeline can block or flood the caller.
- Emit one short marker before each SVG import and preflight unfamiliar assets individually. If a run stalls, use the last phase/icon marker to isolate the blocking asset or layer and terminate only the task-owned automation process.

## SVG import guidance

- Prefer local SVG files with `xmlns` and `viewBox`.
- Preserve explicit multicolor fills, strokes, gradients, and layered paths.
- If an SVG uses `currentColor` and imports blank or incomplete, create a derived Visio-safe copy with a controlled default color. Keep the source unchanged.
- `render_figure.ps1` performs that `currentColor` conversion automatically and caches the derived copy by source hash and chosen color. Explicit multicolor SVG values are not recolored.
- Treat remote references, scripts, unresolved `<use>` elements, filters, masks, and animation as compatibility risks that require visual QA.
