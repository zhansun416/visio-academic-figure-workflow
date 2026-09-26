# Dependencies and portability

Python 3.10+ is the scripting runtime. Install `requirements.txt`: pywin32 on Windows, Pillow for comparison, pypdfium2 for full-page PNG previews. Native rendering/COM validation require Windows desktop Visio. SVG library, preflight and ZIP/XML/manifest inspection use the standard library. No command delegates to PowerShell.

Visio must expose `BASIC_M.VSSX` (`Rectangle`, `Circle`, `Ellipse`, `Diamond`) and `SSFLOW_M.VSSX` (`Dynamic connector`). Discovery scans Office installations; set `VISIO_STENCIL_ROOT` for custom locations. Use `check-env --probe-visio` to verify actual masters.

The renderer owns a separate Visio application and closes only its documents/application. Makepy wrappers use a process-owned temporary cache; never delete another Office workflow's global cache. Phases go to stderr and JSON results to stdout and optional `--report`.

## Persistent SVG library

Default: the Windows user's actual Documents directory plus `Codex/svg-library` (or `~/Documents/Codex/svg-library` elsewhere). Override with `VISIO_FIGURE_SVG_LIBRARY` or `svg --library <directory> ...`. The public package has an empty bootstrap library; private installed caches must survive updates.

`svg init` creates a missing library. `svg sync` diagnoses unindexed/missing/duplicate/hash-mismatched entries. `svg sync --repair` indexes valid unindexed files as `unknown-review`; it never silently rewrites hashes or deletes missing assets. Writes use atomic replacement and a Python-process lock. Unrelated JavaScript/legacy writers must coordinate externally; they do not share this lock.

## Verification

`python -m unittest discover -s tests -v` tests portable contracts and negative cases. `python scripts/self_test.py --output <work>/regression --visio` renders/reopens Visio fixtures in separate processes. Standard hosted CI runners can run portable tests only.

Skill-authoring validation optionally needs PyYAML. LibreOffice is not needed for VSDX. `iconfont-svg-finder` is optional for missing assets or explicit icon searches.
