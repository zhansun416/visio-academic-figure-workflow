param(
    [Parameter(Mandatory=$true)]
    [string]$SvgPath,
    [string]$ManifestPath
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'svg_library_common.ps1')

$assessment = Get-SvgFileAssessment $SvgPath
$xml = $assessment.xml
$raw = $assessment.raw
$record = [ordered]@{
    path = $assessment.path
    validXml = $assessment.validXml
    root = $assessment.root
    xmlns = $assessment.xmlns
    viewBox = $assessment.viewBox
    pathCount = if ($xml) { @($xml.SelectNodes('//*[local-name()="path"]')).Count } else { 0 }
    shapeCount = if ($xml) { @($xml.SelectNodes('//*[local-name()="rect" or local-name()="circle" or local-name()="ellipse" or local-name()="line" or local-name()="polyline" or local-name()="polygon"]')).Count } else { 0 }
    textCount = if ($xml) { @($xml.SelectNodes('//*[local-name()="text"]')).Count } else { 0 }
    usesCurrentColor = ($raw -match '(?i)\bcurrentColor\b')
    externalReferences = @($assessment.externalReferences)
    compatibilityRisks = @($assessment.compatibilityRisks)
    warnings = @($assessment.warnings)
    completeForDirectReuse = $assessment.readyForVisioImport
}

$json = $record | ConvertTo-Json -Depth 8
if ($ManifestPath) {
    $manifestFull = [IO.Path]::GetFullPath($ManifestPath)
    $manifestDir = Split-Path -Parent $manifestFull
    if ($manifestDir -and -not (Test-Path -LiteralPath $manifestDir)) { New-Item -ItemType Directory -Force -Path $manifestDir | Out-Null }
    [IO.File]::WriteAllText($manifestFull, $json, (New-Object Text.UTF8Encoding($false)))
}
Write-Output $json
