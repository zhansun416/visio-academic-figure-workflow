# Dependencies and portability

## Required at runtime

- Windows PowerShell 5.1 or PowerShell 7.
- Microsoft Visio with COM automation enabled. The workflow is designed for Windows desktop Visio and does not require Visio Online. The installed Visio content must expose `BASIC_M.VSSX` (`Rectangle`, `Circle`, `Ellipse`, and `Diamond`) and `SSFLOW_M.VSSX` (`Dynamic connector`); set `VISIO_STENCIL_ROOT` when they are outside the normal Office tree.
- A local working directory with write permission for temporary previews, manifests, and the final VSDX.

## Bundled or optional

- The public package ships an empty `assets/svg-library` bootstrap directory. The actual working library is initialized outside the Skill package so it persists across projects, conversations, and Skill updates.
- LibreOffice is optional and only used for internal conversion or document diagnostics; it is not required to deliver a VSDX.
- `iconfont-svg-finder` is an optional external skill. Use it only for missing icons or when the user explicitly requests icon searching.

## Persistent library location

By default, scripts use `%USERPROFILE%\Documents\Codex\svg-library`. Run `initialize_svg_library.ps1` once after installation. Set `VISIO_FIGURE_SVG_LIBRARY` to use another shared location. The path is user-level rather than project-level, so a new project or conversation can query the same approved collection.

Run `sync_svg_library.ps1` before drawing. The dependency check fails on unindexed files, missing files, or hash mismatches unless `-AllowInconsistentLibrary` is explicitly supplied for diagnosis. Manifest writes use a cross-process lock and atomic replacement so concurrent conversations do not silently overwrite one another.

After installation, run `scripts\self_test.ps1`. It builds a disposable SVG library, tests complete-package short-circuiting and search, then creates and validates a real VSDX unless `-SkipVisio` is supplied. Its Visio tests cover horizontal and vertical auto-routing, explicit `Xn` compatibility, real-row enumeration, rotated and grouped endpoints, near-equal edge tolerance, glued connectors, and axis alignment. Use `-KeepArtifacts` only when the generated preview or VSDX needs inspection.

## External library override

Set `VISIO_FIGURE_SVG_LIBRARY` to point to a user-maintained persistent SVG library when desired. If it is unset, the scripts use the default user-level path above. This keeps the Skill shareable while allowing each user or organization to maintain a larger private collection.

The local installed copy may contain a curated personal cache. That cache is intentionally not part of the public GitHub package.

## Runtime versus authoring dependencies

The runtime scripts are PowerShell and do not require Python. Python plus PyYAML is only needed for the optional Skill-authoring `quick_validate.py` check, not for reconstructing or validating a Visio figure.
