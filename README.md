# Visio Academic Figure Workflow

Rebuild academic flowcharts and framework figures as editable Visio VSDX files from reference images, SVG packages, persistent SVG libraries, or existing Visio sources.

> Final figures may still require minor manual fine-tuning.

## What it provides

- Paper-ready, editable VSDX reconstruction for academic flowcharts and framework figures.
- Times New Roman typography enforcement for English figure text.
- Native Visio nodes and glued dynamic connectors, with explicit connection-point handling to reduce misplaced arrows.
- A customer-SVG-first asset gate: when a complete usable SVG package is supplied, iconfont search is skipped.
- Support for monochrome, multicolor, and `currentColor` SVG assets, with preparation and compatibility checks for Visio.
- A persistent user SVG library that can be reused across projects and conversations.
- Validation for asset completeness, SVG usability, connection integrity, and final VSDX output.

## Requirements

- Windows 10/11.
- Microsoft Visio desktop (required for VSDX generation and COM-based validation).
- Windows PowerShell 5.1 or PowerShell 7.
- Optional: Python 3.10+ and PyYAML for authoring-time validation of the skill package.

This workflow is designed for local Windows execution. GitHub Actions cannot perform the Visio COM steps on a standard hosted runner.

## Quick start

Run these commands from the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\initialize_svg_library.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\check_dependencies.ps1 -RequireVisio
powershell -ExecutionPolicy Bypass -File .\scripts\self_test.ps1
```

The default persistent SVG library is created under:

```text
%USERPROFILE%\Documents\Codex\svg-library
```

Set `VISIO_FIGURE_SVG_LIBRARY` to use another shared library location. The public repository contains only an empty bootstrap library; customer-owned or project-specific SVGs should be added locally after checking their licensing.

## Typical asset workflow

1. Inspect the reference image and supplied assets.
2. Use the supplied SVG package first. If it covers the required icons, skip iconfont search.
3. If assets are missing, search and collect additional SVG candidates, then prepare them for Visio.
4. Build the figure with explicit node roles and connection points.
5. Validate typography, layout, arrows, asset coverage, and VSDX output.
6. Open the final VSDX in Visio and make any remaining fine adjustments required by the target paper layout.

Useful commands:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\sync_svg_library.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\find_svg_library.ps1 -Query "calendar parking charging"
powershell -ExecutionPolicy Bypass -File .\scripts\check_dependencies.ps1 -RequireVisio
powershell -ExecutionPolicy Bypass -File .\scripts\self_test.ps1
```

See [`SKILL.md`](SKILL.md) for the complete workflow and script routing. See [`references/dependencies.md`](references/dependencies.md) and [`references/asset-licensing.md`](references/asset-licensing.md) for dependency and asset-use guidance.

## Repository layout

```text
SKILL.md                         Main workflow instructions
agents/openai.yaml               Skill metadata
scripts/                         PowerShell workflow and validation scripts
assets/templates/                Figure specification template
assets/svg-library/              Empty public bootstrap library
references/                      Dependency, licensing, and Visio notes
```

## License

MIT License. See [`LICENSE`](LICENSE).

