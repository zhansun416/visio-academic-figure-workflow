param(
    [string]$LibraryRoot,
    [switch]$Repair
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'svg_library_common.ps1')

$root = Resolve-SvgLibraryRoot $LibraryRoot
$repaired = New-Object System.Collections.Generic.List[string]
if ($Repair) {
    Invoke-WithSvgLibraryLock $root {
        $state = Get-SvgLibraryState $root
        $entries = @($state.entries)
        foreach ($fileName in @($state.unindexedFiles)) {
            $path = Join-Path $state.iconsDirectory $fileName
            $assessment = Get-SvgFileAssessment $path
            if (-not $assessment.readyForVisioImport) {
                throw "Cannot index invalid or risky SVG '$fileName': $(@($assessment.warnings) -join '; ')"
            }
            $entries += [pscustomobject]@{
                file = $fileName
                path = (Join-Path 'icons' $fileName)
                sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
                tags = @(Get-SvgTagsFromName $fileName)
                category = 'recovered'
                sourceLabel = 'recovered-unindexed'
                sourceUrl = $null
                license = 'unknown-review'
                addedAt = (Get-Date).ToString('s')
            }
            $repaired.Add($fileName)
        }
        if ($repaired.Count -gt 0) { Write-SvgLibraryManifestAtomic $root $entries | Out-Null }
    } | Out-Null
}

$final = Get-SvgLibraryState $root
$report = [ordered]@{
    libraryRoot = $final.root
    svgFileCount = $final.svgFileCount
    manifestEntryCount = $final.manifestEntryCount
    unindexedFiles = @($final.unindexedFiles)
    missingFiles = @($final.missingFiles)
    hashMismatchFiles = @($final.hashMismatchFiles)
    repairedFiles = @($repaired)
    consistent = $final.consistent
}
$report | ConvertTo-Json -Depth 8
if (-not $report.consistent) { exit 1 }
exit 0
