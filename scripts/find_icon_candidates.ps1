param(
    [Parameter(Mandatory=$true)]
    [string]$Query,

    [Parameter(Mandatory=$true)]
    [string]$OutputDir,

    [string]$Prefix,

    [int]$Limit = 12,

    [int]$DownloadTop = 3,

    [int]$TimeoutSec = 20
)

$ErrorActionPreference = 'Stop'
if ($Limit -lt 1) { $Limit = 1 }
if ($Limit -gt 64) { $Limit = 64 }
if ($DownloadTop -lt 0) { $DownloadTop = 0 }
if ($DownloadTop -gt 10) { $DownloadTop = 10 }
if ($TimeoutSec -lt 5) { $TimeoutSec = 5 }
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$queryEncoded = [uri]::EscapeDataString($Query)
$uri = "https://api.iconify.design/search?query=$queryEncoded&limit=$Limit"
if ($Prefix) {
    $uri += "&prefix=$([uri]::EscapeDataString($Prefix))"
}

$response = Invoke-RestMethod -Uri $uri -Method Get -TimeoutSec $TimeoutSec
$icons = @($response.icons)
$records = New-Object System.Collections.Generic.List[object]
$downloaded = 0

foreach ($icon in $icons) {
    $parts = $icon -split ':', 2
    if ($parts.Count -ne 2) { continue }
    $iconPrefix = $parts[0]
    $iconName = $parts[1]
    $collectionInfo = $null
    if ($response.collections) {
        $prop = $response.collections.PSObject.Properties | Where-Object { $_.Name -eq $iconPrefix } | Select-Object -First 1
        if ($prop) { $collectionInfo = $prop.Value }
    }

    $safeName = [regex]::Replace("$iconPrefix-$iconName", '[^A-Za-z0-9._-]', '_')
    $svgPath = Join-Path $OutputDir ($safeName + '.svg')
    $downloadStatus = 'not-attempted'
    $sourceUrl = "https://api.iconify.design/$iconPrefix/$iconName.svg?download=1"

    if ($downloaded -lt $DownloadTop) {
        try {
            Invoke-WebRequest -Uri $sourceUrl -OutFile $svgPath -UseBasicParsing -TimeoutSec $TimeoutSec
            $svgText = Get-Content -LiteralPath $svgPath -Raw -Encoding UTF8
            if ($svgText -notmatch '<svg\b') {
                throw 'Downloaded response is not an SVG document.'
            }
            $downloadStatus = 'downloaded'
            $downloaded++
        } catch {
            $downloadStatus = 'download-failed: ' + $_.Exception.Message
            if (Test-Path -LiteralPath $svgPath) { Remove-Item -LiteralPath $svgPath -Force }
        }
    }

    $records.Add([ordered]@{
        icon = $icon
        localPath = if ($downloadStatus -eq 'downloaded') { $svgPath } else { $null }
        sourceUrl = $sourceUrl
        collection = if ($collectionInfo) { $collectionInfo.name } else { $null }
        license = if ($collectionInfo -and $collectionInfo.license) { $collectionInfo.license.spdx } else { $null }
        matchLabel = 'candidate'
        downloadStatus = $downloadStatus
        query = $Query
    })
}

$manifestPath = Join-Path $OutputDir 'iconify-candidates.json'
$json = ConvertTo-Json -InputObject @($records) -Depth 10
[IO.File]::WriteAllText([IO.Path]::GetFullPath($manifestPath), $json, (New-Object Text.UTF8Encoding($false)))
Write-Output ("Candidates: {0}" -f $records.Count)
Write-Output ("Manifest: {0}" -f $manifestPath)
