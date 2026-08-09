param(
    [Parameter(Mandatory=$true)]
    [string]$SvgDirectory,
    [string]$ManifestPath,
    [switch]$CompleteSet,
    [string]$VisioSafeDirectory,
    [string]$VisioDefaultColor = '#1E2A44'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'svg_library_common.ps1')

if ($VisioDefaultColor -notmatch '^#[0-9A-Fa-f]{6}$') { throw 'VisioDefaultColor must be a six-digit hex color.' }
$resolvedDirectory = (Resolve-Path -LiteralPath $SvgDirectory).Path
$resolvedVisioDirectory = $null
if ($VisioSafeDirectory) {
    $resolvedVisioDirectory = [IO.Path]::GetFullPath($VisioSafeDirectory)
    New-Item -ItemType Directory -Force -Path $resolvedVisioDirectory | Out-Null
}
$files = @(Get-ChildItem -LiteralPath $resolvedDirectory -Filter '*.svg' -File | Sort-Object Name)
$records = @(foreach ($file in $files) {
    $assessment = Get-SvgFileAssessment $file.FullName
    $raw = $assessment.raw
    $xml = $assessment.xml
    $usesCurrentColor = ($raw -match '(?i)\bcurrentColor\b')
    $colorValues = @([regex]::Matches($raw, '(?i)(?:fill|stroke)\s*=\s*["'']([^"'']+)["'']') | ForEach-Object { $_.Groups[1].Value.Trim() } | Where-Object {
        $_ -and $_ -notmatch '^(?i:none|currentColor|inherit|transparent)$' -and $_ -notmatch '^url\('
    } | Sort-Object -Unique)
    $visioSafePath = $null
    if ($resolvedVisioDirectory -and $assessment.readyForVisioImport) {
        $visioSafePath = Join-Path $resolvedVisioDirectory $file.Name
        if ($usesCurrentColor) {
            $safeSvg = [regex]::Replace($raw, '(?i)\bcurrentColor\b', $VisioDefaultColor)
            [IO.File]::WriteAllText($visioSafePath, $safeSvg, (New-Object Text.UTF8Encoding($false)))
        } else {
            Copy-Item -LiteralPath $file.FullName -Destination $visioSafePath -Force
        }
    }

    [ordered]@{
        file = $file.Name
        path = $file.FullName
        validXml = $assessment.validXml
        root = $assessment.root
        xmlns = $assessment.xmlns
        viewBox = $assessment.viewBox
        pathCount = if ($xml) { @($xml.SelectNodes('//*[local-name()="path"]')).Count } else { 0 }
        shapeCount = if ($xml) { @($xml.SelectNodes('//*[local-name()="rect" or local-name()="circle" or local-name()="ellipse" or local-name()="line" or local-name()="polyline" or local-name()="polygon"]')).Count } else { 0 }
        externalReferences = @($assessment.externalReferences)
        compatibilityRisks = @($assessment.compatibilityRisks)
        warnings = @($assessment.warnings)
        usesCurrentColor = $usesCurrentColor
        explicitColors = @($colorValues)
        explicitColorCount = $colorValues.Count
        colorMode = if ($colorValues.Count -gt 1) { 'multicolor-or-layered' } elseif ($usesCurrentColor) { 'theme-color' } else { 'single-or-unknown' }
        visioSafePath = $visioSafePath
        readyForVisioImport = $assessment.readyForVisioImport
    }
})

$valid = @($records | Where-Object { $_.readyForVisioImport })
$invalid = @($records | Where-Object { -not $_.readyForVisioImport })
$skipFinder = ($CompleteSet -and $records.Count -gt 0 -and $invalid.Count -eq 0)
$report = [ordered]@{
    directory = $resolvedDirectory
    visioSafeDirectory = $resolvedVisioDirectory
    assetCount = $records.Count
    readyAssetCount = $valid.Count
    invalidAssetCount = $invalid.Count
    completeSetDeclared = [bool]$CompleteSet
    skipIconfontSvgFinder = $skipFinder
    nextStep = if ($skipFinder) { 'Use customer SVG assets directly; do not run iconfont-svg-finder.' } else { 'Use ready customer SVG assets first; search only for explicitly missing assets.' }
    assets = @($records)
}

$json = $report | ConvertTo-Json -Depth 10
if ($ManifestPath) {
    $manifestFull = [IO.Path]::GetFullPath($ManifestPath)
    $manifestDir = Split-Path -Parent $manifestFull
    if ($manifestDir -and -not (Test-Path -LiteralPath $manifestDir)) { New-Item -ItemType Directory -Force -Path $manifestDir | Out-Null }
    [IO.File]::WriteAllText($manifestFull, $json, (New-Object Text.UTF8Encoding($false)))
}
Write-Output $json
