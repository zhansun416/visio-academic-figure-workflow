# Visio Academic Figure Workflow

Reconstruct dense reference diagrams as editable Visio VSDX with Python. Source-led inventories and regional visual review preserve detailed panels, repeated marks, rich icons, labels and topology.

## Capabilities

- Native nested groups, named ports, glued connectors, feedback routes and independent labels.
- Native paths, circuits, editable text, matrices and chart marks.
- Rich SVG preparation, import stroke scaling and a persistent library with provenance/hashes.
- Saved-file verification of groups, text, geometry, endpoint targets, arrows and bends.
- Source manifests with actual object mappings and draft/delivery gates.
- Full-page/regional side-by-side, overlay and difference images without stretching.

The source-manifest and comparison approach draws from `image-to-editable-ppt`; see [third-party notices](THIRD_PARTY_NOTICES.md). Structural checks alone do not establish visual fidelity.

## Quick start

Python 3.10+, Pillow, pypdfium2 for PNG previews, and Windows desktop Visio with pywin32 for native rendering/validation. Portable inspection and library operations do not require Visio.

```text
python -m pip install -r requirements.txt
python scripts/visio_workflow.py check-env --probe-visio
python -m unittest discover -s tests -v
python scripts/self_test.py --output work/regression --visio
python scripts/visio_workflow.py preflight assets/templates/dense_figure_spec.json
python scripts/visio_workflow.py render assets/templates/dense_figure_spec.json work/dense.vsdx --preview work/dense.png
python scripts/visio_workflow.py validate work/dense.vsdx --spec assets/templates/dense_figure_spec.json --require-glued --no-raster
```

Run from the repository root. `--report <path>` precedes the subcommand. See [SKILL.md](SKILL.md), [figure spec](references/figure-spec.md), [source manifest](references/reconstruction-manifest.md), [dependencies](references/dependencies.md) and [SVG licensing](references/asset-licensing.md).

The persistent library defaults to Documents/Codex/svg-library. Override with `VISIO_FIGURE_SVG_LIBRARY`. Private SVG collections and user images are not included publicly.

## Migration

Former `.ps1` entry points are replaced by `visio_workflow.py`, `compare_renders.py` and `self_test.py`. Existing flat/dense specs and `library.json` records remain supported. No command launches PowerShell. Native rendering still requires Windows desktop Visio.

Manual bends need review when individual nodes move. SVGs can become native groups or foreign media depending on Visio/source; saved-package inspection is authoritative. Review icon interiors and the original source before delivery.

## License

MIT; see [LICENSE](LICENSE) and [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
