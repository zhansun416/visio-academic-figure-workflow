param(
    [Parameter(Mandatory=$true)]
    [string]$SourceDirectory,

    [string]$LibraryRoot,
    [string]$Category = 'general',
    [string]$SourceLabel = 'local-cache',
    [string]$SourceUrl,
    [string]$License = 'unknown-review',
    [string[]]$Tags,
    [switch]$IncludeSourcePath
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'svg_library_common.ps1')

$root = Resolve-SvgLibraryRoot $LibraryRoot
$source = (Resolve-Path -LiteralPath $SourceDirectory).Path
$files = @(Get-ChildItem -LiteralPath $source -Filter '*.svg' -File | Sort-Object Name)
if ($files.Count -eq 0) { throw "No SVG files found: $source" }

$prepared = foreach ($file in $files) {
    $assessment = Get-SvgFileAssessment $file.FullName
    if (-not $assessment.readyForVisioImport) {
        throw "SVG is not safe for the persistent library '$($file.Name)': $(@($assessment.warnings) -join '; ')"
    }
    [pscustomobject]@{
        file = $file
        hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        tags = if ($Tags) { @($Tags) } else { @(Get-SvgTagsFromName $file.Name) }
    }
}

$added = New-Object System.Collections.Generic.List[string]
$updated = New-Object System.Collections.Generic.List[string]
$skipped = New-Object System.Collections.Generic.List[string]
Invoke-WithSvgLibraryLock $root {
    $manifest = Read-SvgLibraryManifest $root
    $entries = @($manifest.entries)

    foreach ($item in $prepared) {
        $existingHash = @($entries | Where-Object { ([string]$_.sha256).ToLowerInvariant() -eq $item.hash } | Select-Object -First 1)
        if ($existingHash.Count -gt 0) {
            $entry = $existingHash[0]
            $changed = $false
            $currentSourceUrl = if ($entry.PSObject.Properties['sourceUrl']) { [string]$entry.sourceUrl } else { '' }
            $currentLicense = if ($entry.PSObject.Properties['license']) { [string]$entry.license } else { '' }
            if ($SourceUrl -and $currentSourceUrl -ne $SourceUrl) { $entry | Add-Member -NotePropertyName sourceUrl -NotePropertyValue $SourceUrl -Force; $changed = $true }
            if ($License -and $License -ne 'unknown-review' -and $currentLicense -ne $License) { $entry | Add-Member -NotePropertyName license -NotePropertyValue $License -Force; $changed = $true }
            if ($changed) { $updated.Add([string]$entry.file) } else { $skipped.Add([string]$entry.file) }
            continue
        }

        $targetName = $item.file.Name
        $target = Join-Path $manifest.paths.icons $targetName
        if (Test-Path -LiteralPath $target) {
            $targetHash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash.ToLowerInvariant()
            if ($targetHash -ne $item.hash) {
                $stem = [IO.Path]::GetFileNameWithoutExtension($targetName)
                $targetName = $stem + '-' + $item.hash.Substring(0, 8) + '.svg'
                $target = Join-Path $manifest.paths.icons $targetName
            }
        }
        if (-not (Test-Path -LiteralPath $target)) { Copy-Item -LiteralPath $item.file.FullName -Destination $target }

        $entryData = [ordered]@{
            file = $targetName
            path = (Join-Path 'icons' $targetName)
            sha256 = $item.hash
            tags = @($item.tags)
            category = $Category
            sourceLabel = $SourceLabel
            sourceUrl = if ($SourceUrl) { $SourceUrl } else { $null }
            license = $License
            addedAt = (Get-Date).ToString('s')
        }
        if ($IncludeSourcePath) { $entryData.sourcePath = $item.file.FullName }
        $entries += [pscustomobject]$entryData
        $added.Add($targetName)
    }

    if ($added.Count -gt 0 -or $updated.Count -gt 0) { Write-SvgLibraryManifestAtomic $root $entries | Out-Null }
} | Out-Null

$total = (Get-SvgLibraryState $root).manifestEntryCount

[ordered]@{
    libraryRoot = $root
    sourceDirectory = $source
    addedCount = $added.Count
    updatedCount = $updated.Count
    skippedDuplicateCount = $skipped.Count
    totalEntryCount = $total
    addedFiles = @($added)
    manifest = Join-Path $root 'library.json'
} | ConvertTo-Json -Depth 8
