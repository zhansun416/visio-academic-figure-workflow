# SVG asset licensing

The public package starts with an empty `assets/svg-library/library.json`. Customer-approved SVGs are added to the persistent user-level library, whose manifest records each source label, source URL where known, license field, and SHA-256 hash.

## Local use versus public release

- Assets marked `MIT`, `ISC`, or `Apache-2.0` have an identified upstream license in the current manifest. Preserve the source URL and attribution requirements when redistributing them.
- Assets marked `unknown-review` came from a local cache or a source whose redistribution license was not verified. They are included for local workflow testing only.
- Before publishing this package to GitHub, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\audit_library_licenses.ps1 -PackageLibrary -PublicRelease
```

`-PackageLibrary` audits the empty or curated library shipped inside the Skill. Omit it, or pass `-LibraryRoot`, to audit the persistent user library. A public-release audit requires a consistent manifest and zero review-required entries. Replace or remove every unknown-license asset before public release; do not infer permission from visual similarity or from the fact that an SVG was downloadable.

## Provenance

When adding an icon, retain its upstream URL, package name, license, and SHA-256 hash in `library.json`. Use `add_svg_library.ps1` for additions to the persistent library, then manually correct the license and source metadata when the upstream information is available. Local source paths are omitted by default so customer or workstation paths are not leaked into a shared manifest; use `-IncludeSourcePath` only when that provenance is intentionally required. The public package's bootstrap directory and the user-level library are intentionally separate.
