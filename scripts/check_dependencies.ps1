param(
    [string]$LibraryRoot,
    [string]$AssetDirectory,
    [switch]$ProbeVisio,
    [switch]$RequireVisio,
    [switch]$RequireLibreOffice,
    [switch]$AllowInconsistentLibrary
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'svg_library_common.ps1')

$root = Resolve-SvgLibraryRoot $LibraryRoot
$paths = Initialize-SvgLibraryInternal $root
if (-not $AssetDirectory) { $AssetDirectory = $paths.icons }
$state = Get-SvgLibraryState $root
$checks = [ordered]@{}
$checks.powerShell = [ordered]@{ passed = $true; version = $PSVersionTable.PSVersion.ToString() }
$checks.compression = [ordered]@{ passed = $false; detail = $null }
$checks.visio = [ordered]@{ passed = $null; detail = 'not requested' }
$checks.visioStencils = [ordered]@{ passed = $null; detail = 'not requested'; files = @() }
$checks.svgLibrary = [ordered]@{
    passed = ($state.consistent -or $AllowInconsistentLibrary)
    consistent = $state.consistent
    path = $root
    svgFileCount = $state.svgFileCount
    manifestEntryCount = $state.manifestEntryCount
    unindexedCount = $state.unindexedFiles.Count
    missingCount = $state.missingFiles.Count
    hashMismatchCount = $state.hashMismatchFiles.Count
}
$checks.assetDirectory = [ordered]@{ passed = (Test-Path -LiteralPath $AssetDirectory); count = 0; path = [IO.Path]::GetFullPath($AssetDirectory) }
$checks.libreOffice = [ordered]@{ passed = $false; path = $null }

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $checks.compression = [ordered]@{ passed = $true; detail = 'System.IO.Compression.FileSystem loaded' }
} catch {
    $checks.compression = [ordered]@{ passed = $false; detail = $_.Exception.Message }
}

if ($checks.assetDirectory.passed) {
    $checks.assetDirectory.count = @(Get-ChildItem -LiteralPath $AssetDirectory -Filter '*.svg' -File).Count
}

if ($ProbeVisio -or $RequireVisio) {
    $visio = $null
    $openedStencils = New-Object System.Collections.Generic.List[object]
    try {
        $visio = New-Object -ComObject Visio.Application
        $checks.visio = [ordered]@{ passed = $true; detail = 'Visio.Application COM created' }
        $stencilRoots = New-Object System.Collections.Generic.List[string]
        if ($env:VISIO_STENCIL_ROOT) { $stencilRoots.Add($env:VISIO_STENCIL_ROOT) }
        if ($env:ProgramFiles) { $stencilRoots.Add((Join-Path $env:ProgramFiles 'Microsoft Office')) }
        if (${env:ProgramFiles(x86)}) { $stencilRoots.Add((Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office')) }
        $requiredStencils = [ordered]@{ 'BASIC_M.VSSX' = @('Rectangle','Circle','Ellipse','Diamond'); 'SSFLOW_M.VSSX' = @('Dynamic connector') }
        $stencilFiles = New-Object System.Collections.Generic.List[string]
        foreach ($stencilName in $requiredStencils.Keys) {
            $match = $null
            foreach ($candidateRoot in @($stencilRoots | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -Unique)) {
                $match = Get-ChildItem -LiteralPath $candidateRoot -Recurse -File -Filter $stencilName -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($match) { break }
            }
            if (-not $match) { throw "Required Visio stencil was not found: $stencilName" }
            $stencilDoc = $visio.Documents.Open($match.FullName)
            $openedStencils.Add($stencilDoc)
            foreach ($masterName in $requiredStencils[$stencilName]) {
                try { [void]$stencilDoc.Masters.ItemU($masterName) }
                catch { throw "Required Visio master '$masterName' was not found in $stencilName." }
            }
            $stencilFiles.Add($match.FullName)
        }
        $checks.visioStencils = [ordered]@{ passed = $true; detail = 'Required stencils and masters opened successfully'; files = $stencilFiles.ToArray() }
    } catch {
        if (-not $checks.visio.passed) { $checks.visio = [ordered]@{ passed = $false; detail = $_.Exception.Message } }
        $checks.visioStencils = [ordered]@{ passed = $false; detail = $_.Exception.Message; files = @() }
    } finally {
        foreach ($stencilDoc in $openedStencils) { try { $stencilDoc.Close() } catch {} }
        if ($visio) { try { $visio.Quit() } catch {} }
    }
}

$libreCandidates = New-Object System.Collections.Generic.List[string]
if ($env:ProgramFiles) { $libreCandidates.Add((Join-Path $env:ProgramFiles 'LibreOffice\program\soffice.exe')) }
if (${env:ProgramFiles(x86)}) { $libreCandidates.Add((Join-Path ${env:ProgramFiles(x86)} 'LibreOffice\program\soffice.exe')) }
$sofficeCommand = Get-Command soffice.exe -ErrorAction SilentlyContinue
if ($sofficeCommand) { $libreCandidates.Add($sofficeCommand.Source) }
$librePath = @($libreCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1)
if ($librePath.Count -gt 0) { $checks.libreOffice = [ordered]@{ passed = $true; path = $librePath[0] } }

$required = @($checks.compression.passed, $checks.svgLibrary.passed, $checks.assetDirectory.passed)
if ($RequireVisio) { $required += [bool]$checks.visio.passed; $required += [bool]$checks.visioStencils.passed }
if ($RequireLibreOffice) { $required += $checks.libreOffice.passed }
$report = [ordered]@{ passed = (@($required | Where-Object { -not $_ }).Count -eq 0); checks = $checks }
$report | ConvertTo-Json -Depth 8
if (-not $report.passed) { exit 1 }
exit 0
