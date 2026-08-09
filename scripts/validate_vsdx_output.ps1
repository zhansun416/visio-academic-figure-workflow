param(
    [Parameter(Mandatory=$true)]
    [string]$VsdxPath,
    [string]$ExpectedFont = 'Times New Roman',
    [string]$OutputDirectory,
    [switch]$RequireSingleOutput,
    [switch]$RequireGluedConnectors,
    [switch]$RequireAxisAlignedConnectors,
    [double]$AxisAlignmentToleranceIn = 0.0001,
    [switch]$DisallowRasterMedia,
    [double]$MinimumFontSizePt = 0,
    [string[]]$AllowedAcronyms = @('EV','SOC','SOCP','OPF','SOCP-OPF','V2G','SROBUST'),
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'visio_connection_common.ps1')
if ($AxisAlignmentToleranceIn -le 0) { throw 'AxisAlignmentToleranceIn must be positive.' }
$resolved = (Resolve-Path -LiteralPath $VsdxPath).Path
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead($resolved)
try {
    $pageEntries = @($zip.Entries | Where-Object { $_.FullName -match '^visio/pages/page\d+\.xml$' })
    $media = @($zip.Entries | Where-Object { $_.FullName -like 'visio/media/*' })
} finally { $zip.Dispose() }
$rasterMedia = @($media | Where-Object { $_.FullName -match '(?i)\.(png|jpe?g|bmp|gif|tiff?)$' })

$visio = New-Object -ComObject Visio.Application
$visio.Visible = $false
$doc = $null
$pageShapeCount = 0
$totalShapeCount = 0
$textCount = 0
$dynamicConnectorCount = 0
$decorativeOneDCount = 0
$fonts = New-Object 'System.Collections.Generic.HashSet[string]'
$texts = New-Object System.Collections.Generic.List[string]
$lowFontShapes = New-Object System.Collections.Generic.List[object]
$ungluedEndpoints = New-Object System.Collections.Generic.List[object]
$misalignedAxisConnectors = New-Object System.Collections.Generic.List[object]
$masterlessTwoDCount = 0

function Walk-Shape($Shape, [int]$PageIndex, $Collector) {
    [void]$Collector.Add($Shape)
    $text = ''
    try { $text = [string]$Shape.Text } catch {}
    if ($text.Trim()) {
        [void]$script:texts.Add($text)
        try { [void]$script:fonts.Add([string]$Shape.CellsU('Char.Font').FormulaU) } catch {}
        if ($MinimumFontSizePt -gt 0) {
            try {
                $sizePt = [double]$Shape.CellsU('Char.Size').Result('pt')
                if ($sizePt -gt 0 -and $sizePt -lt $MinimumFontSizePt) {
                    $script:lowFontShapes.Add([ordered]@{ page = $PageIndex; shapeId = $Shape.ID; sizePt = [math]::Round($sizePt, 2); text = $text.Trim() })
                }
            } catch {}
        }
    }
    try {
        if ($Shape.OneD -eq 0 -and $null -eq $Shape.Master) { $script:masterlessTwoDCount++ }
    } catch {}
    for ($i = 1; $i -le $Shape.Shapes.Count; $i++) { Walk-Shape $Shape.Shapes.Item($i) $PageIndex $Collector }
}

function Get-ConnectorEndpointPagePoint($Shape, [ValidateSet('Begin','End')] [string]$Endpoint) {
    $x = [double]$Shape.CellsU($Endpoint + 'X').ResultIU
    $y = [double]$Shape.CellsU($Endpoint + 'Y').ResultIU
    $parent = $null
    try { $parent = $Shape.ContainingShape } catch {}
    if ($parent -and [int]$parent.ID -ne 0) {
        return Convert-VisioLocalPointToPage -Shape $parent -X $x -Y $y
    }
    return [pscustomobject]@{ X = $x; Y = $y }
}

try {
    $doc = $visio.Documents.Open($resolved)
    for ($pageIndex = 1; $pageIndex -le $doc.Pages.Count; $pageIndex++) {
        $page = $doc.Pages.Item($pageIndex)
        $pageShapeCount += $page.Shapes.Count
        $pageShapes = New-Object System.Collections.Generic.List[object]
        for ($i = 1; $i -le $page.Shapes.Count; $i++) { Walk-Shape $page.Shapes.Item($i) $pageIndex $pageShapes }
        $totalShapeCount += $pageShapes.Count

        $glue = @{}
        $glueTargets = @{}
        for ($i = 1; $i -le $page.Connects.Count; $i++) {
            $connect = $page.Connects.Item($i)
            try {
                $shapeId = [int]$connect.FromSheet.ID
                $cell = [string]$connect.FromCell.NameU
                if ([string]::IsNullOrWhiteSpace($cell)) { $cell = [string]$connect.FromCell.Name }
                if (-not $glue.ContainsKey($shapeId)) { $glue[$shapeId] = @{} }
                $glue[$shapeId][$cell] = $true
                if (-not $glueTargets.ContainsKey($shapeId)) { $glueTargets[$shapeId] = @{} }
                $glueTargets[$shapeId][$cell] = $connect.ToSheet
            } catch {}
        }

        for ($i = 0; $i -lt $pageShapes.Count; $i++) {
            $shape = $pageShapes[$i]
            $isOneD = $false
            try { $isOneD = ($shape.OneD -ne 0) } catch {}
            if (-not $isOneD) { continue }
            $masterName = ''
            try { if ($shape.Master) { $masterName = [string]$shape.Master.NameU } } catch {}
            if ($masterName -notmatch '(?i)connector') { $decorativeOneDCount++; continue }
            $dynamicConnectorCount++
            foreach ($endpoint in @('BeginX','EndX')) {
                $isGlued = $glue.ContainsKey([int]$shape.ID) -and $glue[[int]$shape.ID].ContainsKey($endpoint)
                if (-not $isGlued) {
                    $ungluedEndpoints.Add([ordered]@{ page = $pageIndex; shapeId = $shape.ID; master = $masterName; endpoint = $endpoint })
                }
            }
            if ($RequireAxisAlignedConnectors -and $glueTargets.ContainsKey([int]$shape.ID) -and $glueTargets[[int]$shape.ID].ContainsKey('BeginX') -and $glueTargets[[int]$shape.ID].ContainsKey('EndX')) {
                $beginTarget = $glueTargets[[int]$shape.ID]['BeginX']
                $endTarget = $glueTargets[[int]$shape.ID]['EndX']
                $beginCenter = Get-VisioShapePageCenter -Shape $beginTarget
                $endCenter = Get-VisioShapePageCenter -Shape $endTarget
                $targetDeltaX = [math]::Abs($endCenter.X - $beginCenter.X)
                $targetDeltaY = [math]::Abs($endCenter.Y - $beginCenter.Y)
                $axis = $null
                $deviation = 0.0
                if ($targetDeltaY -le $AxisAlignmentToleranceIn -and $targetDeltaX -gt $AxisAlignmentToleranceIn) {
                    $axis = 'horizontal'
                    $beginPoint = Get-ConnectorEndpointPagePoint -Shape $shape -Endpoint Begin
                    $endPoint = Get-ConnectorEndpointPagePoint -Shape $shape -Endpoint End
                    $deviation = [math]::Abs($endPoint.Y - $beginPoint.Y)
                } elseif ($targetDeltaX -le $AxisAlignmentToleranceIn -and $targetDeltaY -gt $AxisAlignmentToleranceIn) {
                    $axis = 'vertical'
                    $beginPoint = Get-ConnectorEndpointPagePoint -Shape $shape -Endpoint Begin
                    $endPoint = Get-ConnectorEndpointPagePoint -Shape $shape -Endpoint End
                    $deviation = [math]::Abs($endPoint.X - $beginPoint.X)
                }
                if ($axis -and $deviation -gt $AxisAlignmentToleranceIn) {
                    $misalignedAxisConnectors.Add([ordered]@{ page = $pageIndex; shapeId = $shape.ID; master = $masterName; axis = $axis; deviationInches = [math]::Round($deviation, 6) })
                }
            }
        }
    }
    $textCount = $texts.Count
} finally {
    if ($doc) { try { $doc.Saved = $true } catch {}; try { $doc.Close() } catch {} }
    if ($visio) { try { $visio.Quit() } catch {} }
}

$outputFiles = @()
if ($OutputDirectory -and (Test-Path -LiteralPath $OutputDirectory)) { $outputFiles = @(Get-ChildItem -LiteralPath $OutputDirectory -File) }
$singleOutputPassed = (-not $RequireSingleOutput) -or ($outputFiles.Count -eq 1 -and $outputFiles[0].FullName -eq $resolved)
$wrongFont = @($fonts | Where-Object { $_ -notmatch [regex]::Escape($ExpectedFont) })
$literalEscapes = @($texts | Where-Object { $_ -match '`n|`r|`t' })
$allowed = @{}; foreach ($item in $AllowedAcronyms) { $allowed[$item.ToUpperInvariant()] = $true }
$unapprovedAllCaps = New-Object System.Collections.Generic.List[string]
foreach ($text in $texts) {
    foreach ($match in [regex]::Matches($text, '\b[A-Z][A-Z0-9-]{2,}\b')) {
        $word = $match.Value.ToUpperInvariant()
        if (-not $allowed.ContainsKey($word)) { $unapprovedAllCaps.Add($match.Value) }
    }
}
$topologyPassed = (-not $RequireGluedConnectors) -or ($ungluedEndpoints.Count -eq 0)
$axisAlignmentPassed = (-not $RequireAxisAlignedConnectors) -or ($misalignedAxisConnectors.Count -eq 0)
$rasterPassed = (-not $DisallowRasterMedia) -or ($rasterMedia.Count -eq 0)
$fontSizePassed = ($MinimumFontSizePt -le 0) -or ($lowFontShapes.Count -eq 0)
$casePassed = ($unapprovedAllCaps.Count -eq 0)
$fontValues = [string[]]::new($fonts.Count)
$fonts.CopyTo($fontValues)
[Array]::Sort($fontValues)
$lowFontValues = $lowFontShapes.ToArray()
$ungluedValues = $ungluedEndpoints.ToArray()
$misalignedAxisValues = $misalignedAxisConnectors.ToArray()
$allCapsSet = New-Object 'System.Collections.Generic.SortedSet[string]' ([StringComparer]::OrdinalIgnoreCase)
foreach ($value in $unapprovedAllCaps) { [void]$allCapsSet.Add($value) }
$allCapsValues = [string[]]::new($allCapsSet.Count)
$allCapsSet.CopyTo($allCapsValues)
$rasterMediaValues = [string[]]@($rasterMedia | ForEach-Object { $_.FullName })
$outputFileValues = [string[]]@($outputFiles | ForEach-Object { $_.Name })
$report = [ordered]@{
    file = $resolved
    fileBytes = (Get-Item -LiteralPath $resolved).Length
    pageCount = $pageEntries.Count
    pageShapeCount = $pageShapeCount
    totalShapeCount = $totalShapeCount
    textCount = $textCount
    fonts = $fontValues
    wrongFontFormulas = $wrongFont
    lowFontShapes = $lowFontValues
    literalEscapes = $literalEscapes.Count
    unapprovedAllCaps = $allCapsValues
    mediaCount = $media.Count
    rasterMedia = $rasterMediaValues
    dynamicConnectorCount = $dynamicConnectorCount
    decorativeOneDCount = $decorativeOneDCount
    ungluedConnectorEndpoints = $ungluedValues
    misalignedAxisConnectors = $misalignedAxisValues
    axisAlignmentToleranceIn = $AxisAlignmentToleranceIn
    masterlessTwoDCount = $masterlessTwoDCount
    singleOutputPassed = $singleOutputPassed
    outputFiles = $outputFileValues
    passed = ($pageEntries.Count -gt 0 -and $pageShapeCount -gt 0 -and $wrongFont.Count -eq 0 -and $literalEscapes.Count -eq 0 -and $singleOutputPassed -and $topologyPassed -and $axisAlignmentPassed -and $rasterPassed -and $fontSizePassed -and $casePassed)
}
$json = $report | ConvertTo-Json -Depth 10
if ($ReportPath) {
    $reportDir = Split-Path -Parent ([IO.Path]::GetFullPath($ReportPath))
    if ($reportDir -and -not (Test-Path -LiteralPath $reportDir)) { New-Item -ItemType Directory -Force -Path $reportDir | Out-Null }
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($ReportPath), $json, (New-Object Text.UTF8Encoding($false)))
}
Write-Output $json
if (-not $report.passed) { exit 1 }
exit 0
