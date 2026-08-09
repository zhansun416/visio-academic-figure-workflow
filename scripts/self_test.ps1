param(
    [switch]$SkipVisio,
    [switch]$KeepArtifacts
)

$ErrorActionPreference = 'Stop'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('visio-academic-figure-selftest-' + [Guid]::NewGuid().ToString('N'))
$failure = $null
$report = $null

function Invoke-ConnectionPointRegressionTests {
    . (Join-Path $PSScriptRoot 'visio_connection_common.ps1')
    $visio = New-Object -ComObject Visio.Application
    $visio.Visible = $false
    $doc = $null
    $stencil = $null
    try {
        $roots = New-Object System.Collections.Generic.List[string]
        if ($env:VISIO_STENCIL_ROOT) { $roots.Add($env:VISIO_STENCIL_ROOT) }
        if ($env:ProgramFiles) { $roots.Add((Join-Path $env:ProgramFiles 'Microsoft Office')) }
        if (${env:ProgramFiles(x86)}) { $roots.Add((Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office')) }
        $stencilPath = $null
        foreach ($root in @($roots | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -Unique)) {
            $candidate = Get-ChildItem -LiteralPath $root -Recurse -File -Filter 'BASIC_M.VSSX' -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($candidate) { $stencilPath = $candidate.FullName; break }
        }
        if (-not $stencilPath) { throw 'BASIC_M.VSSX was not found for connection-point regression tests.' }

        $doc = $visio.Documents.Add('')
        $page = $doc.Pages.Item(1)
        $drawingWindow = $visio.ActiveWindow
        $stencil = $visio.Documents.Open($stencilPath)
        $master = $stencil.Masters.ItemU('Rectangle')

        $plain = $page.Drop($master, 2, 2)
        $actualPoints = @(Get-VisioConnectionPointCandidates -Shape $plain)
        if ($actualPoints.Count -ne 5) {
            throw "Connection-point enumeration returned $($actualPoints.Count) rows; expected the 5 real Rectangle rows."
        }

        $rotated = $page.Drop($master, 5, 2)
        $rotated.CellsU('Width').FormulaU = '2 in'
        $rotated.CellsU('Height').FormulaU = '1 in'
        $rotated.CellsU('Angle').FormulaU = '90 deg'
        $rotatedRight = Get-VisioCardinalConnectionName -Shape $rotated -Direction right -ShapeId 'rotated'
        if ($rotatedRight -ne 'X1') {
            throw "Rotated page-right connection resolved to $rotatedRight; expected X1 after page-coordinate transformation."
        }

        $tolerance = $page.Drop($master, 8, 2)
        $tolerance.CellsU('Connections.X1').FormulaU = 'Width'
        $tolerance.CellsU('Connections.Y1').FormulaU = '0'
        $tolerance.CellsU('Connections.X2').FormulaU = 'Width-0.0000005 in'
        $tolerance.CellsU('Connections.Y2').FormulaU = 'Height/2'
        $tolerantRight = Get-VisioCardinalConnectionName -Shape $tolerance -Direction right -ShapeId 'tolerance'
        if ($tolerantRight -ne 'X2') {
            throw "Near-equal page-right candidates resolved to $tolerantRight; expected centered X2 within tolerance."
        }

        $groupA = $page.Drop($master, 11, 2)
        $groupB = $page.Drop($master, 13, 2)
        $drawingWindow.Activate()
        try { $drawingWindow.DeselectAll() } catch {}
        $drawingWindow.Select($groupA, 2)
        $drawingWindow.Select($groupB, 2)
        $group = $drawingWindow.Selection.Group()
        $group.CellsU('Angle').FormulaU = '90 deg'
        $groupChild = $group.Shapes.Item(1)
        $groupedRight = Get-VisioCardinalConnectionName -Shape $groupChild -Direction right -ShapeId 'grouped-rotated'
        if ($groupedRight -ne 'X1') {
            throw "Grouped and rotated page-right connection resolved to $groupedRight; expected X1."
        }

        return [ordered]@{
            realConnectionPointCount = $actualPoints.Count
            rotatedRightConnection = $rotatedRight
            tolerantRightConnection = $tolerantRight
            groupedRotatedRightConnection = $groupedRight
        }
    } finally {
        if ($doc) { try { $doc.Saved = $true } catch {}; try { $doc.Close() } catch {} }
        if ($stencil) { try { $stencil.Close() } catch {} }
        if ($visio) { try { $visio.Quit() } catch {} }
    }
}

try {
    $source = Join-Path $testRoot 'source'
    $library = Join-Path $testRoot 'library'
    $safe = Join-Path $testRoot 'visio-safe'
    $delivery = Join-Path $testRoot 'delivery'
    $preview = Join-Path $testRoot 'preview'
    New-Item -ItemType Directory -Force -Path $source,$delivery,$preview | Out-Null
    $fixtureSvg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><path fill="currentColor" d="M6 30 32 8l26 22v26H39V40H25v16H6Z"/></svg>'
    [IO.File]::WriteAllText((Join-Path $source 'home.svg'), $fixtureSvg, (New-Object Text.UTF8Encoding($false)))

    & (Join-Path $PSScriptRoot 'initialize_svg_library.ps1') -LibraryRoot $library | Out-Null
    $add = (& (Join-Path $PSScriptRoot 'add_svg_library.ps1') -SourceDirectory $source -LibraryRoot $library -Category 'self-test' -SourceLabel 'generated-self-test' | ConvertFrom-Json)
    $sync = (& (Join-Path $PSScriptRoot 'sync_svg_library.ps1') -LibraryRoot $library | ConvertFrom-Json)
    $matches = @(& (Join-Path $PSScriptRoot 'find_svg_library.ps1') -LibraryRoot $library -Query 'home' -Limit 3 | ConvertFrom-Json)
    $assets = (& (Join-Path $PSScriptRoot 'prepare_svg_assets.ps1') -SvgDirectory $source -VisioSafeDirectory $safe -CompleteSet | ConvertFrom-Json)

    $visioCheck = $null
    $connectionCheck = $null
    if (-not $SkipVisio) {
        $connectionCheck = Invoke-ConnectionPointRegressionTests
        $specPath = Join-Path $testRoot 'self-test-spec.json'
        $spec = @'
{
  "reference": { "widthPx": 800, "heightPx": 450 },
  "page": { "widthIn": 8, "heightIn": 4.5 },
  "styles": {
    "module": { "fill": "#F3F8FF", "line": "#1F5FB8", "lineWeightPt": 1, "font": "Times New Roman", "fontSizePt": 12, "align": "center" },
    "connector": { "line": "#111111", "lineWeightPt": 1 }
  },
  "shapes": [
    { "id": "icon", "kind": "svg-asset", "bboxPx": [40, 125, 160, 200], "assetRef": "library:home.svg" },
    { "id": "input", "kind": "rect", "bboxPx": [280, 170, 150, 80], "text": "Input", "style": "module" },
    { "id": "output", "kind": "rect", "bboxPx": [610, 170, 150, 80], "text": "Output", "style": "module" },
    { "id": "upper", "kind": "rect", "bboxPx": [470, 20, 100, 55], "text": "Upper", "style": "module", "angleDeg": 90 },
    { "id": "lower", "kind": "rect", "bboxPx": [470, 370, 100, 55], "text": "Lower", "style": "module" }
  ],
  "connectors": [
    { "id": "flow", "from": "input", "to": "output", "fromConnection": "auto", "toConnection": "auto", "arrow": "end", "style": "connector" },
    { "id": "explicit-return", "from": "output", "to": "input", "fromConnection": "X4", "toConnection": "Connections.X2", "arrow": "none", "style": "connector" },
    { "id": "vertical-flow", "from": "upper", "to": "lower", "fromConnection": "auto", "toConnection": "auto", "arrow": "end", "style": "connector" }
  ]
}
'@
        [IO.File]::WriteAllText($specPath, $spec, (New-Object Text.UTF8Encoding($false)))
        $vsdx = Join-Path $delivery 'self-test.vsdx'
        $png = Join-Path $preview 'self-test.png'
        & (Join-Path $PSScriptRoot 'render_figure.ps1') -SpecPath $specPath -OutputPath $vsdx -PreviewPath $png -LibraryRoot $library | Out-Null
        $visioCheck = (& (Join-Path $PSScriptRoot 'validate_vsdx_output.ps1') -VsdxPath $vsdx -ExpectedFont 'Times New Roman' -OutputDirectory $delivery -RequireSingleOutput -RequireGluedConnectors -RequireAxisAlignedConnectors -DisallowRasterMedia -MinimumFontSizePt 8 | ConvertFrom-Json)
    }

    $passed = ($add.addedCount -eq 1 -and $sync.consistent -and $matches.Count -ge 1 -and $assets.skipIconfontSvgFinder -and ($SkipVisio -or ($visioCheck.passed -and $visioCheck.dynamicConnectorCount -eq 3 -and $connectionCheck.realConnectionPointCount -eq 5 -and $connectionCheck.rotatedRightConnection -eq 'X1' -and $connectionCheck.tolerantRightConnection -eq 'X2' -and $connectionCheck.groupedRotatedRightConnection -eq 'X1')))
    $report = [ordered]@{
        passed = $passed
        testRoot = if ($KeepArtifacts) { $testRoot } else { $null }
        libraryConsistent = [bool]$sync.consistent
        libraryMatchCount = $matches.Count
        completeSvgSetSkipsFinder = [bool]$assets.skipIconfontSvgFinder
        visioTested = (-not $SkipVisio)
        visioValidationPassed = if ($SkipVisio) { $null } else { [bool]$visioCheck.passed }
        gluedConnectorCount = if ($SkipVisio) { $null } else { [int]$visioCheck.dynamicConnectorCount }
        ungluedEndpointCount = if ($SkipVisio) { $null } else { @($visioCheck.ungluedConnectorEndpoints).Count }
        misalignedAxisConnectorCount = if ($SkipVisio) { $null } else { @($visioCheck.misalignedAxisConnectors).Count }
        connectionPointRegression = $connectionCheck
    }
    if (-not $passed) { throw 'One or more self-test assertions failed.' }
} catch {
    $failure = $_
} finally {
    if (-not $KeepArtifacts -and (Test-Path -LiteralPath $testRoot)) {
        $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
        $resolvedTestRoot = [IO.Path]::GetFullPath($testRoot)
        if (-not $resolvedTestRoot.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or (Split-Path -Leaf $resolvedTestRoot) -notlike 'visio-academic-figure-selftest-*') {
            throw "Refusing to remove an unexpected self-test path: $resolvedTestRoot"
        }
        Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force
    }
}

if ($failure) {
    [ordered]@{ passed = $false; error = $failure.Exception.Message; testRoot = if ($KeepArtifacts) { $testRoot } else { $null } } | ConvertTo-Json -Depth 8
    exit 1
}
$report | ConvertTo-Json -Depth 8
exit 0
