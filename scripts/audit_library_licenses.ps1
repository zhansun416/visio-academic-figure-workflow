param(
    [string]$LibraryRoot,
    [switch]$PackageLibrary,
    [switch]$PublicRelease
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'svg_library_common.ps1')

if ($PackageLibrary -and $LibraryRoot) { throw 'Use either -PackageLibrary or -LibraryRoot, not both.' }
$root = if ($PackageLibrary) {
    [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\assets\svg-library'))
} else {
    Resolve-SvgLibraryRoot $LibraryRoot
}
$state = Get-SvgLibraryState $root
$review = @($state.entries | Where-Object {
    $license = if ($_.PSObject.Properties['license']) { [string]$_.license } else { '' }
    [string]::IsNullOrWhiteSpace($license) -or $license -match '(?i)unknown|review|unverified|proprietary'
})
$knownCount = $state.manifestEntryCount - $review.Count
$report = [ordered]@{
    libraryRoot = $state.root
    svgFileCount = $state.svgFileCount
    entryCount = $state.manifestEntryCount
    knownLicenseCount = $knownCount
    reviewRequiredCount = $review.Count
    unindexedFiles = @($state.unindexedFiles)
    missingFiles = @($state.missingFiles)
    hashMismatchFiles = @($state.hashMismatchFiles)
    reviewRequired = @($review | ForEach-Object {
        [ordered]@{
            file = $_.file
            license = if ($_.PSObject.Properties['license']) { $_.license } else { $null }
            sourceUrl = if ($_.PSObject.Properties['sourceUrl']) { $_.sourceUrl } else { $null }
            sourceLabel = if ($_.PSObject.Properties['sourceLabel']) { $_.sourceLabel } else { $null }
        }
    })
    libraryConsistent = $state.consistent
    publicReleaseReady = ($state.consistent -and $review.Count -eq 0)
}
$report | ConvertTo-Json -Depth 8
if ($PublicRelease -and -not $report.publicReleaseReady) { exit 1 }
exit 0
