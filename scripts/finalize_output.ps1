param(
    [Parameter(Mandatory=$true)]
    [string]$OutputDirectory,

    [Parameter(Mandatory=$true)]
    [string]$FinalVsdxName,

    [Parameter(Mandatory=$true)]
    [string]$ArchiveDirectory
)

$ErrorActionPreference = 'Stop'
$resolvedOutput = (Resolve-Path -LiteralPath $OutputDirectory).Path
if ([IO.Path]::GetFileName($FinalVsdxName) -ne $FinalVsdxName -or [IO.Path]::GetExtension($FinalVsdxName) -ne '.vsdx') { throw 'FinalVsdxName must be a .vsdx file name without a directory.' }
$finalPath = Join-Path $resolvedOutput $FinalVsdxName
if (-not (Test-Path -LiteralPath $finalPath)) { throw "Final VSDX not found: $finalPath" }
New-Item -ItemType Directory -Force -Path $ArchiveDirectory | Out-Null

$moved = New-Object System.Collections.Generic.List[string]
foreach ($file in @(Get-ChildItem -LiteralPath $resolvedOutput -File)) {
    if ($file.FullName -eq ([IO.Path]::GetFullPath($finalPath))) { continue }
    $destination = Join-Path $ArchiveDirectory $file.Name
    if (Test-Path -LiteralPath $destination) {
        $destination = Join-Path $ArchiveDirectory (([IO.Path]::GetFileNameWithoutExtension($file.Name)) + '-' + (Get-Date -Format 'yyyyMMdd-HHmmssfff') + $file.Extension)
    }
    Move-Item -LiteralPath $file.FullName -Destination $destination
    $moved.Add($file.Name)
}

$remaining = @(Get-ChildItem -LiteralPath $resolvedOutput -File)
$report = [ordered]@{
    outputDirectory = $resolvedOutput
    finalFile = $finalPath
    remainingFiles = @($remaining | ForEach-Object { $_.Name })
    archivedFiles = @($moved)
    passed = ($remaining.Count -eq 1 -and $remaining[0].Name -eq $FinalVsdxName)
}
$report | ConvertTo-Json -Depth 6
if (-not $report.passed) { exit 1 }
exit 0
