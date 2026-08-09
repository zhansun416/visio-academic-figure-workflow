param(
    [string]$LibraryRoot
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'svg_library_common.ps1')

$root = Resolve-SvgLibraryRoot $LibraryRoot
$paths = Initialize-SvgLibraryInternal $root
$state = Get-SvgLibraryState $root
[ordered]@{
    libraryRoot = $paths.root
    manifest = $paths.manifest
    iconsDirectory = $paths.icons
    svgFileCount = $state.svgFileCount
    manifestEntryCount = $state.manifestEntryCount
    consistent = $state.consistent
    initialized = $true
} | ConvertTo-Json -Depth 6
