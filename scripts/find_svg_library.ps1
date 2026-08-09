param(
    [Parameter(Mandatory=$true)]
    [string]$Query,
    [string]$LibraryRoot,
    [int]$Limit = 20
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'svg_library_common.ps1')

if ([string]::IsNullOrWhiteSpace($Query)) { throw 'Query cannot be empty.' }
if ($Limit -lt 1) { $Limit = 1 }
if ($Limit -gt 100) { $Limit = 100 }
$root = Resolve-SvgLibraryRoot $LibraryRoot
$manifest = Read-SvgLibraryManifest $root
$terms = @($Query.ToLowerInvariant() -split '[^\p{L}\p{Nd}]+' | Where-Object { $_ })

$matches = foreach ($entry in @($manifest.entries)) {
    $file = if ($entry.PSObject.Properties['file']) { [string]$entry.file } else { '' }
    $category = if ($entry.PSObject.Properties['category']) { [string]$entry.category } else { '' }
    $role = if ($entry.PSObject.Properties['role']) { [string]$entry.role } else { '' }
    $matchLabel = if ($entry.PSObject.Properties['matchLabel']) { [string]$entry.matchLabel } else { '' }
    $tags = if ($entry.PSObject.Properties['tags']) { @($entry.tags | ForEach-Object { ([string]$_).ToLowerInvariant() }) } else { @() }
    $haystack = ($file + ' ' + $category + ' ' + $role + ' ' + $matchLabel + ' ' + ($tags -join ' ')).ToLowerInvariant()
    $score = 0
    foreach ($term in $terms) {
        if ($tags -contains $term) { $score += 4 }
        elseif ($file.ToLowerInvariant().Contains($term)) { $score += 2 }
        elseif ($haystack.Contains($term)) { $score += 1 }
    }
    if ($score -le 0) { continue }

    $fullPath = Resolve-SvgEntryPath $root $entry
    if (-not (Test-Path -LiteralPath $fullPath)) { continue }
    [pscustomobject]@{
        score = $score
        file = $file
        path = if ($entry.PSObject.Properties['path']) { [string]$entry.path } else { Join-Path 'icons' $file }
        fullPath = $fullPath
        tags = @($tags)
        category = $category
        role = $role
        matchLabel = $matchLabel
        sourceLabel = if ($entry.PSObject.Properties['sourceLabel']) { $entry.sourceLabel } else { $null }
        license = if ($entry.PSObject.Properties['license']) { $entry.license } else { $null }
        sourceUrl = if ($entry.PSObject.Properties['sourceUrl']) { $entry.sourceUrl } else { $null }
    }
}

$sorted = @($matches | Sort-Object @{Expression='score';Descending=$true}, file | Select-Object -First $Limit)
ConvertTo-Json -InputObject $sorted -Depth 8
