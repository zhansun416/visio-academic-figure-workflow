# Source manifest and delivery gate

Author the inventory from the source **before** generating the scene. Schema `1.0` uses `pages`, source pixel `canvas`, `regions` and `elements`. Region/component boxes are absolute source pixels. A component can map to multiple native objects (e.g. all cells of a matrix); list actual saved names/types using `inspect`, not invented target names.

```json
{
  "schemaVersion": "1.0",
  "pages": [{
    "page": 1,
    "canvas": {"widthPx": 800, "heightPx": 600},
    "sourceReview": {"status":"review-needed"},
    "regions": [{
      "id":"matrix-region", "bbox":{"x":40,"y":80,"w":240,"h":180},
      "critical":true,
      "fidelity":{"status":"review-needed"},
      "elements":[{
        "id":"matrix", "bbox":{"x":50,"y":90,"w":210,"h":150},
        "sourceInventory":"Three rows, four columns, twelve distinct cells and their labels",
        "expectedEditability":"native",
        "representation":{"editability":"native","mode":"mixed"},
        "completion":{"status":"complete"},
        "objectMap":[{"name":"cell-1","type":"shape"}],
        "fidelity":{"status":"review-needed"}
      }]
    }]
  }]
}
```

This shortened example is a draft: map all real objects, not only `cell-1`, before delivery. Keep descriptions of icons' visible features, repeat counts, topology and unreadable marks in `sourceInventory` or separate asset-match records.

## Fields

- `expectedEditability`: `native` or `asset`.
- `representation.editability`: `native`, `asset`, `partial`.
- `representation.mode`: `native-text`, `native-shape`, `native-line`, `native-connector`, `native-group`, `svg-picture`, `raster-picture`, `mixed`, `deferred`.
- `completion.status`: `complete`, `partial`, `deferred`.
- `objectMap`: actual `name` and `type`. Types: `text`, `shape`, `line`, `connector`, `group`, `picture-svg`, `picture-raster`, `picture-unknown`. Each saved object belongs to one source component. Include group wrappers and imported native children.
- `fidelity.status`: `verified`, `review-needed`, `failed`, `not-applicable`. Verified records require `methods` (e.g. full-source, regional-crops, overlay) and concrete `notes`.

Assets require `assetReason` and `editabilityBoundary`. Native-to-asset/partial representations and incomplete components require `degradation.reason` and `degradation.editabilityBoundary`. Record approximations independently of editability: native paths can still be visually inaccurate.

The package reader distinguishes actual ForeignData/media from groups and geometry. Imported SVGs sometimes become native editable paths in Visio; verify their saved children. Do not infer a native shape from filename, producer intent, or an SVG parent group containing a foreign picture.

## Commands and guarantees

```text
python scripts/visio_workflow.py check-manifest source-manifest.json --mode draft
python scripts/visio_workflow.py inspect figure.vsdx --manifest source-manifest.json --mode delivery
python scripts/compare_renders.py reference.png preview.png comparison --manifest source-manifest.json --page 1
```

Draft checks declared structures and mapping consistency while allowing explicitly incomplete work. Delivery requires all elements complete, all mappings present and correctly typed, every output object mapped, whole-source review, all element reviews and critical-region reviews verified. A standalone manifest check cannot check the VSDX; delivery must use `inspect ... --manifest`.

Never mark review fields verified by looping over the generated scene. Inspect source/render comparisons and record real observations. Failed region reviews remain failed. An unresolved source detail must remain visible in the ledger and the user-facing result. A structurally valid ledger cannot detect a source item that was never inventoried.

## Comparison artifacts

The Pillow comparison writes HTML, full-page images and region crops: source, rendered, side-by-side, 50% overlay and enhanced difference. The original images remain unchanged. Aspect mismatch (>0.5% by default) prevents resizing/comparison. Proportional fitting allows minor pixel rounding only. Full page bounds must be retained during export; content-tight cropping cannot be repaired by stretching. Diagnostic MAE/RMS scores are not acceptance thresholds.
