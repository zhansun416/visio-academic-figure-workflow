param([Parameter(Mandatory=$true)][string]$SpecPath, [string]$ReportPath)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'figure_spec_common.ps1')
$spec = Get-Content -LiteralPath $SpecPath -Raw -Encoding UTF8 | ConvertFrom-Json
$scene = Expand-FigureSpec $spec
$report = [ordered]@{
    passed = $true
    shapeCount = $scene.Shapes.Count
    groupCount = $scene.Groups.Count
    connectorCount = $scene.Connectors.Count
    textCount = @($scene.Shapes | Where-Object { $_.text }).Count + @($scene.Connectors | Where-Object { $_.label.text }).Count
    manualRouteCount = @($scene.Connectors | Where-Object { $_.waypointsPx }).Count
    warnings = $scene.Warnings
}
$json = $report | ConvertTo-Json -Depth 6
if ($ReportPath) { [IO.File]::WriteAllText([IO.Path]::GetFullPath($ReportPath), $json, (New-Object Text.UTF8Encoding($false))) }
$json
