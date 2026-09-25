param(
    [Parameter(Mandatory=$true)]
    [string]$SpecPath,

    [Parameter(Mandatory=$true)]
    [string]$OutputPath,

    [string]$PreviewPath,

    [string[]]$ExportFormats,

    [string]$LibraryRoot,

    [string]$SvgDefaultColor = '#1E2A44',

    [switch]$Visible,

    [switch]$Overwrite
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'svg_library_common.ps1')
. (Join-Path $PSScriptRoot 'visio_connection_common.ps1')
. (Join-Path $PSScriptRoot 'figure_spec_common.ps1')
. (Join-Path $PSScriptRoot 'visio_scene_common.ps1')
if ($SvgDefaultColor -notmatch '^#[0-9A-Fa-f]{6}$') { throw 'SvgDefaultColor must be a six-digit hex color.' }
$resolvedSpec = (Resolve-Path -LiteralPath $SpecPath).Path
$specDir = Split-Path -Parent $resolvedSpec
$spec = Get-Content -LiteralPath $resolvedSpec -Raw -Encoding UTF8 | ConvertFrom-Json

$scene = Expand-FigureSpec $spec
foreach ($warning in $scene.Warnings) { Write-Warning $warning }

$outputFull = [IO.Path]::GetFullPath($OutputPath)
$outputDir = Split-Path -Parent $outputFull
if (-not (Test-Path -LiteralPath $outputDir)) { New-Item -ItemType Directory -Force -Path $outputDir | Out-Null }

if (Test-Path -LiteralPath $outputFull) {
    if (-not $Overwrite) { throw "Output exists. Re-run with -Overwrite: $outputFull" }
    $backup = Join-Path $outputDir (([IO.Path]::GetFileNameWithoutExtension($outputFull)) + '.backup-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.vsdx')
    Copy-Item -LiteralPath $outputFull -Destination $backup
    Write-Output "Backup: $backup"
}

$script:PageW = [double]$spec.page.widthIn
$script:PageH = [double]$spec.page.heightIn
$script:RefW = [double]$spec.reference.widthPx
$script:RefH = [double]$spec.reference.heightPx
$script:Page = $null
$script:Visio = $null
$script:StencilDocs = @{}
$script:StencilPaths = @{}
$shapeById = @{}
$portCells = @{}
$connectorById = @{}
$labelById = @{}

function Convert-ColorFormula([string]$Color) {
    if (-not $Color -or $Color -eq 'none') { return $null }
    if ($Color -match '^#([0-9A-Fa-f]{6})$') {
        $r = [Convert]::ToInt32($matches[1].Substring(0,2),16)
        $g = [Convert]::ToInt32($matches[1].Substring(2,2),16)
        $b = [Convert]::ToInt32($matches[1].Substring(4,2),16)
        return "RGB($r,$g,$b)"
    }
    return $Color
}

function Set-Cell($Shape, [string]$Cell, [string]$Formula) {
    try {
        $Shape.CellsU($Cell).FormulaU = $Formula
    } catch {
        throw "Failed to set Visio cell '$Cell': $($_.Exception.Message)"
    }
}

function Set-OptionalCell($Shape, [string]$Cell, [string]$Formula) {
    try {
        $Shape.CellsU($Cell).FormulaU = $Formula
        return $true
    } catch {
        Write-Verbose ("Optional Visio cell '{0}' is unavailable: {1}" -f $Cell, $_.Exception.Message)
        return $false
    }
}

function Get-Style($Item) {
    $style = $null
    $styleName = [string](Get-Value $Item 'style' '')
    $stylesProperty = $spec.PSObject.Properties | Where-Object { $_.Name -eq 'styles' } | Select-Object -First 1
    if ($styleName -and $stylesProperty -and $stylesProperty.Value) {
        $prop = $stylesProperty.Value.PSObject.Properties | Where-Object { $_.Name -eq $styleName } | Select-Object -First 1
        if ($prop) { $style = $prop.Value }
    }
    return $style
}

function Get-Value($Item, [string]$Name, $Default) {
    if ($null -eq $Item) { return $Default }
    $prop = $Item.PSObject.Properties | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
    if ($prop -and $null -ne $prop.Value) { return $prop.Value }
    return $Default
}

function Get-Bounds([object[]]$Bbox) {
    if ($Bbox.Count -ne 4) { throw 'bboxPx must contain [x, y, width, height].' }
    $x = [double]$Bbox[0]
    $y = [double]$Bbox[1]
    $w = [double]$Bbox[2]
    $h = [double]$Bbox[3]
    if ($w -lt 0 -or $h -lt 0) { throw 'bboxPx dimensions must be nonnegative.' }
    return [ordered]@{
        left = $script:PageW * $x / $script:RefW
        right = $script:PageW * ($x + $w) / $script:RefW
        top = $script:PageH - ($script:PageH * $y / $script:RefH)
        bottom = $script:PageH - ($script:PageH * ($y + $h) / $script:RefH)
    }
}

function Apply-Style($Shape, $Item, $Style, [bool]$NoFill = $false) {
    $fill = Get-Value $Item 'fill' (Get-Value $Style 'fill' 'none')
    $line = Get-Value $Item 'line' (Get-Value $Style 'line' 'none')
    if ($NoFill) { $fill = 'none' }
    if ($fill -eq 'none') {
        Set-Cell $Shape 'FillPattern' '0'
    } else {
        Set-Cell $Shape 'FillPattern' '1'
        Set-Cell $Shape 'FillForegnd' (Convert-ColorFormula $fill)
    }
    if ($line -eq 'none') {
        Set-Cell $Shape 'LinePattern' '0'
    } else {
        $pattern = Get-Value $Item 'linePattern' (Get-Value $Style 'linePattern' 1)
        Set-Cell $Shape 'LinePattern' ([string]$pattern)
        Set-Cell $Shape 'LineColor' (Convert-ColorFormula $line)
        $weight = [double](Get-Value $Item 'lineWeightPt' (Get-Value $Style 'lineWeightPt' 0.8))
        Set-Cell $Shape 'LineWeight' "$weight pt"
    }
    $rounding = Get-Value $Item 'roundingPx' (Get-Value $Style 'roundingPx' 0)
    if ([double]$rounding -gt 0) {
        [void](Set-OptionalCell $Shape 'Rounding' (([double]$rounding * $script:PageW / $script:RefW).ToString([Globalization.CultureInfo]::InvariantCulture) + ' in'))
    }
}

function Apply-Arrow($Shape, $Item) {
    Set-Cell $Shape 'BeginArrow' '0'
    Set-Cell $Shape 'EndArrow' '0'
    $arrow = [string](Get-Value $Item 'arrow' 'none')
    if ($arrow -eq 'end' -or $arrow -eq 'both') { Set-Cell $Shape 'EndArrow' '4' }
    if ($arrow -eq 'begin' -or $arrow -eq 'both') { Set-Cell $Shape 'BeginArrow' '4' }
}

function Apply-Text($Shape, $Item, $Style) {
    $text = Get-Value $Item 'text' ''
    if (-not $text) { return }
    $Shape.Text = [string]$text
    $font = [string](Get-Value $Item 'font' (Get-Value $Style 'font' 'Times New Roman'))
    $fontFormula = 'FONT("' + $font.Replace('"','""') + '")'
    Set-Cell $Shape 'Char.Font' $fontFormula
    Set-Cell $Shape 'Char.Size' (([double](Get-Value $Item 'fontSizePt' (Get-Value $Style 'fontSizePt' 10))).ToString([Globalization.CultureInfo]::InvariantCulture) + ' pt')
    $color = Convert-ColorFormula (Get-Value $Item 'fontColor' (Get-Value $Style 'fontColor' '#111111'))
    Set-Cell $Shape 'Char.Color' $color
    $bold = [bool](Get-Value $Item 'bold' (Get-Value $Style 'bold' $false))
    $italic = [bool](Get-Value $Item 'italic' (Get-Value $Style 'italic' $false))
    $textStyle = 0
    if ($bold) { $textStyle += 1 }
    if ($italic) { $textStyle += 2 }
    Set-Cell $Shape 'Char.Style' ([string]$textStyle)
    $alignName = [string](Get-Value $Item 'align' (Get-Value $Style 'align' 'center'))
    $align = switch ($alignName.ToLowerInvariant()) { 'left' { 0 } 'right' { 2 } default { 1 } }
    Set-Cell $Shape 'Para.HorzAlign' ([string]$align)
    $vertical = [string](Get-Value $Item 'verticalAlign' (Get-Value $Style 'verticalAlign' 'middle'))
    Set-Cell $Shape 'VerticalAlign' ([string](@{ top = 0; middle = 1; bottom = 2 }[$vertical]))
    $margin = [double](Get-Value $Item 'textMarginPt' (Get-Value $Style 'textMarginPt' 2))
    foreach ($cell in @('LeftMargin','RightMargin','TopMargin','BottomMargin')) { Set-Cell $Shape $cell ($margin.ToString([Globalization.CultureInfo]::InvariantCulture) + ' pt') }
}

function Resolve-AssetPath([string]$AssetRef) {
    if ($AssetRef -match '^library:(.+)$') {
        $fileName = $matches[1].Trim()
        $root = Resolve-SvgLibraryRoot $LibraryRoot
        $manifest = Read-SvgLibraryManifest $root
        $entry = @($manifest.entries | Where-Object { [string]$_.file -eq $fileName } | Select-Object -First 1)
        if ($entry.Count -eq 0) { throw "SVG library asset not found: $fileName" }
        $candidate = Resolve-SvgEntryPath $root $entry[0]
    } elseif ([IO.Path]::IsPathRooted($AssetRef)) {
        $candidate = [IO.Path]::GetFullPath($AssetRef)
    } else {
        $candidate = [IO.Path]::GetFullPath((Join-Path $specDir $AssetRef))
    }
    if (-not (Test-Path -LiteralPath $candidate)) { throw "SVG asset not found: $candidate" }
    $assessment = Get-SvgFileAssessment $candidate
    if (-not $assessment.readyForVisioImport) { throw "SVG asset is not safe for Visio import '$candidate': $(@($assessment.warnings) -join '; ')" }
    if ($assessment.raw -match '(?i)\bcurrentColor\b') {
        $cacheRoot = Join-Path $env:LOCALAPPDATA 'Codex\visio-academic-figure-workflow\svg-cache'
        $hash = (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash.ToLowerInvariant()
        $colorKey = $SvgDefaultColor.TrimStart('#').ToLowerInvariant()
        $cacheFile = Join-Path $cacheRoot (([IO.Path]::GetFileNameWithoutExtension($candidate)) + '-' + $hash.Substring(0, 12) + '-' + $colorKey + '.svg')
        $candidate = Invoke-WithSvgLibraryLock $cacheRoot {
            if (-not (Test-Path -LiteralPath $cacheRoot)) { New-Item -ItemType Directory -Force -Path $cacheRoot | Out-Null }
            if (-not (Test-Path -LiteralPath $cacheFile)) {
                $safeSvg = [regex]::Replace($assessment.raw, '(?i)\bcurrentColor\b', $SvgDefaultColor)
                [IO.File]::WriteAllText($cacheFile, $safeSvg, (New-Object Text.UTF8Encoding($false)))
            }
            return $cacheFile
        }
    }
    return $candidate
}

function Resolve-VisioStencil([string]$FileName) {
    if ($script:StencilPaths.ContainsKey($FileName)) { return $script:StencilPaths[$FileName] }
    $roots = @()
    if ($env:VISIO_STENCIL_ROOT) { $roots += $env:VISIO_STENCIL_ROOT }
    if ($env:ProgramFiles) { $roots += (Join-Path $env:ProgramFiles 'Microsoft Office') }
    if (${env:ProgramFiles(x86)}) { $roots += (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office') }
    foreach ($root in ($roots | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique)) {
        $match = Get-ChildItem -LiteralPath $root -Recurse -File -Filter $FileName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($match) { $script:StencilPaths[$FileName] = $match.FullName; return $match.FullName }
    }
    throw "Visio stencil not found: $FileName. Set VISIO_STENCIL_ROOT to the Visio Content directory."
}

function Get-StencilDocument([string]$FileName) {
    if (-not $script:StencilDocs.ContainsKey($FileName)) {
        $path = Resolve-VisioStencil $FileName
        $script:StencilDocs[$FileName] = $script:Visio.Documents.Open($path)
        Write-Verbose ("Stencil: {0} -> {1}" -f $FileName, $path)
    }
    return $script:StencilDocs[$FileName]
}

function Get-NativeMaster($Item, [string]$Kind) {
    $masterName = [string](Get-Value $Item 'master' '')
    if (-not $masterName) {
        $masterName = switch ($Kind.ToLowerInvariant()) {
            'rect' { 'Rectangle' }
            'oval' { 'Ellipse' }
            'circle' { 'Circle' }
            'diamond' { 'Diamond' }
            default { throw "Native shape kind '$Kind' requires a master field." }
        }
    }
    $stencilFile = [string](Get-Value $Item 'stencil' 'BASIC_M.VSSX')
    $stencil = Get-StencilDocument $stencilFile
    try { return $stencil.Masters.ItemU($masterName) }
    catch { throw "Visio master '$masterName' was not found in $stencilFile." }
}

function Set-ShapeBounds($Shape, $Bounds) {
    Set-Cell $Shape 'PinX' ((($Bounds.left + $Bounds.right) / 2).ToString([Globalization.CultureInfo]::InvariantCulture) + ' in')
    Set-Cell $Shape 'PinY' ((($Bounds.bottom + $Bounds.top) / 2).ToString([Globalization.CultureInfo]::InvariantCulture) + ' in')
    Set-Cell $Shape 'Width' (($Bounds.right - $Bounds.left).ToString([Globalization.CultureInfo]::InvariantCulture) + ' in')
    Set-Cell $Shape 'Height' (($Bounds.top - $Bounds.bottom).ToString([Globalization.CultureInfo]::InvariantCulture) + ' in')
}

function Apply-Transform($Shape, $Item) {
    $angleDeg = [double](Get-Value $Item 'angleDeg' 0)
    if ([math]::Abs($angleDeg) -gt 0.0000001) {
        Set-Cell $Shape 'Angle' ($angleDeg.ToString([Globalization.CultureInfo]::InvariantCulture) + ' deg')
    }
}

function Resolve-ConnectionCell($Shape, [string]$Connection, [string]$ShapeId) {
    if ($Connection -like 'port:*') { return $Shape.CellsU($portCells[$ShapeId][$Connection.Substring(5)]) }
    if ($Connection -match '^(?i:left|right|top|bottom)$') {
        $Connection = Get-VisioCardinalConnectionName -Shape $Shape -Direction ($Connection.ToLowerInvariant()) -ShapeId $ShapeId
    }
    $cellName = if ($Connection -match '^Connections\.') { $Connection } else { "Connections.$Connection" }
    if ($Shape.CellExistsU($cellName, 0) -eq 0) {
        throw "Shape '$ShapeId' has no readable connection cell '$cellName'. Use a connection-capable native Master or set the endpoint to auto."
    }
    return $Shape.CellsU($cellName)
}

function Draw-Shape($Item) {
    $id = [string]$Item.id
    if (-not $id) { throw 'Every shape requires an id.' }
    if ($shapeById.ContainsKey($id)) { throw "Duplicate shape id: $id" }
    $kind = [string]$Item.kind
    $bounds = Get-Bounds $Item.bboxPx
    $style = Get-Style $Item
    $shape = $null
    switch ($kind.ToLowerInvariant()) {
        'group' { $master = Get-NativeMaster $Item 'rect'; $shape = $script:Page.Drop($master, (($bounds.left + $bounds.right) / 2), (($bounds.bottom + $bounds.top) / 2)); Set-ShapeBounds $shape $bounds; Apply-Style $shape $Item $style; Apply-Text $shape $Item $style; Set-Cell $shape 'ShapePermeableX' '1'; Set-Cell $shape 'ShapePermeableY' '1' }
        'rect' { $master = Get-NativeMaster $Item $kind; $shape = $script:Page.Drop($master, (($bounds.left + $bounds.right) / 2), (($bounds.bottom + $bounds.top) / 2)); Set-ShapeBounds $shape $bounds; Apply-Transform $shape $Item; Apply-Style $shape $Item $style; Apply-Text $shape $Item $style }
        'oval' { $master = Get-NativeMaster $Item $kind; $shape = $script:Page.Drop($master, (($bounds.left + $bounds.right) / 2), (($bounds.bottom + $bounds.top) / 2)); Set-ShapeBounds $shape $bounds; Apply-Transform $shape $Item; Apply-Style $shape $Item $style; Apply-Text $shape $Item $style }
        'circle' { $master = Get-NativeMaster $Item $kind; $shape = $script:Page.Drop($master, (($bounds.left + $bounds.right) / 2), (($bounds.bottom + $bounds.top) / 2)); Set-ShapeBounds $shape $bounds; Apply-Transform $shape $Item; Apply-Style $shape $Item $style; Apply-Text $shape $Item $style }
        'diamond' { $master = Get-NativeMaster $Item $kind; $shape = $script:Page.Drop($master, (($bounds.left + $bounds.right) / 2), (($bounds.bottom + $bounds.top) / 2)); Set-ShapeBounds $shape $bounds; Apply-Transform $shape $Item; Apply-Style $shape $Item $style; Apply-Text $shape $Item $style }
        'native' { $master = Get-NativeMaster $Item $kind; $shape = $script:Page.Drop($master, (($bounds.left + $bounds.right) / 2), (($bounds.bottom + $bounds.top) / 2)); Set-ShapeBounds $shape $bounds; Apply-Transform $shape $Item; Apply-Style $shape $Item $style; Apply-Text $shape $Item $style }
        'text' { $shape = $script:Page.DrawRectangle($bounds.left,$bounds.bottom,$bounds.right,$bounds.top); Apply-Style $shape $Item $style; Apply-Transform $shape $Item; Apply-Text $shape $Item $style; Set-Cell $shape 'ShapePermeableX' '1'; Set-Cell $shape 'ShapePermeableY' '1' }
        'line' { $shape = $script:Page.DrawLine($bounds.left,$bounds.bottom,$bounds.right,$bounds.top); Apply-Style $shape $Item $style; Apply-Arrow $shape $Item }
        'svg-asset' {
            $assetPath = Resolve-AssetPath ([string]$Item.assetRef)
            Write-Output ("Import: {0}" -f $id)
            $shape = $script:Page.Import($assetPath)
            Set-Cell $shape 'PinX' ((($bounds.left + $bounds.right) / 2).ToString([Globalization.CultureInfo]::InvariantCulture) + ' in')
            Set-Cell $shape 'PinY' ((($bounds.bottom + $bounds.top) / 2).ToString([Globalization.CultureInfo]::InvariantCulture) + ' in')
            Set-Cell $shape 'Width' (($bounds.right - $bounds.left).ToString([Globalization.CultureInfo]::InvariantCulture) + ' in')
            Set-Cell $shape 'Height' (($bounds.top - $bounds.bottom).ToString([Globalization.CultureInfo]::InvariantCulture) + ' in')
            Apply-Transform $shape $Item
        }
        default { throw "Unsupported shape kind '$kind'." }
    }
    Set-VisioSpecIdentity $shape $id
    $portCells[$id] = Add-VisioSpecPorts $shape (Get-Value $Item 'ports' @())
    $shapeById[$id] = $shape
}

function Draw-Connector($Item) {
    $fromId = [string]$Item.from
    $toId = [string]$Item.to
    if (-not $shapeById.ContainsKey($fromId) -or -not $shapeById.ContainsKey($toId)) {
        throw "Connector references unknown shapes: $fromId -> $toId"
    }
    $fromShape = $shapeById[$fromId]
    $toShape = $shapeById[$toId]
    $connectorMaster = (Get-StencilDocument 'SSFLOW_M.VSSX').Masters.ItemU('Dynamic connector')
    $connector = $script:Page.Drop($connectorMaster, 0, 0)
    $fromConnection = [string](Get-Value $Item 'fromConnection' 'auto')
    $toConnection = [string](Get-Value $Item 'toConnection' 'auto')
    $fromAuto = [string]::IsNullOrWhiteSpace($fromConnection) -or $fromConnection -eq 'auto'
    $toAuto = [string]::IsNullOrWhiteSpace($toConnection) -or $toConnection -eq 'auto'
    if ($fromAuto -or $toAuto) {
        $connectionTolerance = [double](Get-Value $Item 'connectionToleranceIn' 0.001)
        $autoPair = Get-VisioAutoConnectionPair -FromShape $fromShape -ToShape $toShape -FromShapeId $fromId -ToShapeId $toId -CoordinateToleranceIn $connectionTolerance
        if ($fromAuto) { $fromConnection = $autoPair.FromConnection }
        if ($toAuto) { $toConnection = $autoPair.ToConnection }
    }
    $connector.CellsU('BeginX').GlueTo((Resolve-ConnectionCell $fromShape $fromConnection $fromId))
    $connector.CellsU('EndX').GlueTo((Resolve-ConnectionCell $toShape $toConnection $toId))
    $style = Get-Style $Item
    $lineValue = [string](Get-Value $Item 'line' (Get-Value $Style 'line' '#111111'))
    if ($lineValue -eq 'none') {
        Set-Cell $connector 'LinePattern' '0'
    } else {
        Set-Cell $connector 'LineColor' (Convert-ColorFormula $lineValue)
        Set-Cell $connector 'LinePattern' ([string](Get-Value $Item 'linePattern' (Get-Value $Style 'linePattern' 1)))
        Set-Cell $connector 'LineWeight' (([double](Get-Value $Item 'lineWeightPt' (Get-Value $Style 'lineWeightPt' 0.8))).ToString([Globalization.CultureInfo]::InvariantCulture) + ' pt')
    }
    Apply-Arrow $connector $Item
    Set-VisioSpecIdentity $connector ([string]$Item.id)
    $routing = [string](Get-Value $Item 'routing' 'auto')
    if ($routing -eq 'orthogonal') { Set-Cell $connector 'ShapeRouteStyle' '1' }
    if ($routing -eq 'straight') { Set-Cell $connector 'ShapeRouteStyle' '2' }
    $jump = [string](Get-Value $Item 'lineJump' 'never')
    $jumpCode = @{ default = 0; never = 1; always = 2; other = 3; neither = 4 }
    if (-not $jumpCode.ContainsKey($jump)) { throw "Unsupported lineJump '$jump'." }
    Set-Cell $connector 'ConLineJumpCode' ([string]$jumpCode[$jump])
    $connectorById[[string]$Item.id] = $connector
    if (Get-Value $Item 'label' $null) {
        $labelSpec = [ordered]@{ id = ([string]$Item.id + '::label'); kind = 'text' }
        foreach ($prop in $Item.label.PSObject.Properties) { $labelSpec[$prop.Name] = $prop.Value }
        $labelSpec.id = [string]$Item.id + '::label'
        $labelSpec.kind = 'text'
        Draw-Shape ([pscustomobject]$labelSpec)
        $labelById[[string]$Item.id] = $shapeById[$labelSpec.id]
    }
}

$script:Visio = New-Object -ComObject Visio.Application
$script:Visio.Visible = [bool]$Visible
$doc = $null
try {
    $doc = $script:Visio.Documents.Add('')
    Write-Output 'Phase: document-created'
    $script:Page = $doc.Pages.Item(1)
    Write-Output 'Phase: page-selected'
    Set-Cell $script:Page.PageSheet 'PageWidth' "$script:PageW in"
    Set-Cell $script:Page.PageSheet 'PageHeight' "$script:PageH in"

    foreach ($item in @($scene.Shapes)) {
        Draw-Shape $item
        Write-Output ("Shape: {0}" -f $item.id)
    }
    foreach ($item in @($scene.Connectors)) {
        Draw-Connector $item
        Write-Output ("Connector: {0}" -f $item.id)
    }

    Write-Output 'Phase: grouping'
    $groupById = @{}
    foreach ($groupSpec in @($scene.Groups | Sort-Object depth -Descending)) {
        $selection = $script:Page.CreateSelection(0)
        $selection.Select($shapeById[$groupSpec.id], 2)
        foreach ($child in @($scene.Shapes | Where-Object { $_.parentId -eq $groupSpec.id })) {
            $member = if ($child.kind -eq 'group') { $groupById[$child.id] } else { $shapeById[$child.id] }
            $selection.Select($member, 2)
        }
        foreach ($edge in @($scene.Connectors | Where-Object { $_.parentId -eq $groupSpec.id })) {
            $selection.Select($connectorById[$edge.id], 2)
            if ($labelById.ContainsKey($edge.id)) { $selection.Select($labelById[$edge.id], 2) }
        }
        $group = $selection.Group()
        Set-VisioSpecIdentity $group ($groupSpec.id + '::group')
        Set-Cell $group 'SelectMode' '1'
        $groupById[$groupSpec.id] = $group
        Write-Output ("Group: {0}" -f $groupSpec.id)
    }
    Write-Output 'Phase: routes'
    foreach ($item in @($scene.Connectors)) {
        if (Get-Value $item 'waypointsPx' $null) { Set-VisioManualRoute $connectorById[[string]$item.id] $item.waypointsPx $script:PageW $script:PageH $script:RefW $script:RefH }
        # At each containment level, edges must remain above child frames.
        $connectorById[$item.id].BringToFront()
        if ($labelById.ContainsKey($item.id)) { $labelById[$item.id].BringToFront() }
    }

    Write-Output 'Phase: saving'
    $doc.SaveAs($outputFull) | Out-Null
    Write-Output "Saved: $outputFull"

    $formats = New-Object System.Collections.Generic.List[string]
    if ($PreviewPath) { $formats.Add('png') | Out-Null }
    foreach ($group in @($ExportFormats)) {
        foreach ($format in ($group -split ',')) {
            $clean = $format.Trim().TrimStart('.').ToLowerInvariant()
            if ($clean -and -not $formats.Contains($clean)) { $formats.Add($clean) | Out-Null }
        }
    }
    foreach ($format in $formats) {
        $path = if ($format -eq 'png' -and $PreviewPath) { [IO.Path]::GetFullPath($PreviewPath) } else { Join-Path $outputDir (([IO.Path]::GetFileNameWithoutExtension($outputFull)) + '.' + $format) }
        $parent = Split-Path -Parent $path
        if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }
        switch ($format) {
            'png' { $script:Page.Export($path) }
            'svg' { $script:Page.Export($path) }
            'pdf' { $doc.ExportAsFixedFormat(1, $path, 1, 0) }
            default { throw "Unsupported export format '$format'. Use png, svg, or pdf." }
        }
        Write-Output ("Exported {0}: {1}" -f $format.ToUpperInvariant(), $path)
    }
} finally {
    if ($doc -ne $null) {
        try { $doc.Saved = $true } catch {}
        try { $doc.Close() } catch {}
    }
    if ($script:StencilDocs) { foreach ($stencil in @($script:StencilDocs.Values)) { try { $stencil.Close() } catch {} } }
    if ($script:Visio -ne $null) { try { $script:Visio.Quit() } catch {} }
}
