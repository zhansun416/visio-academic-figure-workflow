param(
    [Parameter(Mandatory=$true)][string]$VsdxPath,
    [Parameter(Mandatory=$true)][string]$SpecPath,
    [double]$TolerancePx = 1,
    [string]$ReportPath
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'figure_spec_common.ps1')
. (Join-Path $PSScriptRoot 'visio_connection_common.ps1')
if ($TolerancePx -le 0) { throw 'TolerancePx must be positive.' }
$spec = Get-Content -LiteralPath $SpecPath -Raw -Encoding UTF8 | ConvertFrom-Json
$scene = Expand-FigureSpec $spec
$actual = @{}
$issues = New-Object 'System.Collections.Generic.List[string]'
$glue = @{}
$glueCells = @{}
function Collect-SceneShape($Shape) {
    if ($Shape.CellExistsU('User.SpecId', 0) -ne 0) {
        $id = [string]$Shape.CellsU('User.SpecId').ResultStr('')
        if ($actual.ContainsKey($id)) { $issues.Add("Duplicate output id '$id'.") }
        $actual[$id] = $Shape
    }
    # FromConnects/Connects on the actual shape also covers nested connectors.
    for ($i = 1; $i -le $Shape.Connects.Count; $i++) {
        $connection = $Shape.Connects.Item($i)
        $key = [string]$connection.FromSheet.ID + ':' + [string]$connection.FromCell.Name
        $glue[$key] = $connection.ToSheet
        $glueCells[$key] = $connection.ToCell
    }
    for ($i = 1; $i -le $Shape.Shapes.Count; $i++) { Collect-SceneShape $Shape.Shapes.Item($i) }
}
function Normalize-Text([string]$Value) { return (($Value -replace '\r\n?', "`n").Trim()) }
function Check-Text($Shape, [string]$Expected, [string]$Id) {
    if ((Normalize-Text ([string]$Shape.Text)) -cne (Normalize-Text $Expected)) { $issues.Add("Text mismatch for '$Id'.") }
}
function Check-Point($Point, [double]$X, [double]$Y, [string]$Context) {
    $pxX = $Point.X * $spec.reference.widthPx / $spec.page.widthIn
    $pxY = ($spec.page.heightIn - $Point.Y) * $spec.reference.heightPx / $spec.page.heightIn
    if ([math]::Abs($pxX - $X) -gt $TolerancePx -or [math]::Abs($pxY - $Y) -gt $TolerancePx) { $issues.Add("Position mismatch for $Context (expected $X,$Y px; actual $([math]::Round($pxX,2)),$([math]::Round($pxY,2)) px).") }
}
function Check-Parent($Shape, [string]$Expected, [string]$Id) {
    $parent = $Shape.ContainingShape
    $parentId = ''
    if ($parent -and $parent.ID -ne 0 -and $parent.CellExistsU('User.SpecId', 0) -ne 0) { $parentId = [string]$parent.CellsU('User.SpecId').ResultStr('') }
    $expectedId = if ($Expected) { $Expected + '::group' } else { '' }
    if ($parentId -ne $expectedId) { $issues.Add("Group mismatch for '$Id': expected '$expectedId', actual '$parentId'.") }
}
$visio = New-Object -ComObject Visio.Application
$visio.Visible = $false
$doc = $null
try {
    $doc = $visio.Documents.Open((Resolve-Path -LiteralPath $VsdxPath).Path)
    if ($doc.Pages.Count -ne 1) { $issues.Add('Spec renderer expects exactly one page.') }
    $page = $doc.Pages.Item(1)
    for ($i = 1; $i -le $page.Shapes.Count; $i++) { Collect-SceneShape $page.Shapes.Item($i) }
    foreach ($item in $scene.Shapes) {
        $id = [string]$item.id
        if (-not $actual.ContainsKey($id)) { $issues.Add("Missing shape '$id'."); continue }
        $shape = $actual[$id]
        Check-Text $shape ([string]$item.text) $id
        $bbox = $item.bboxPx
        Check-Point (Get-VisioShapePageCenter $shape) ($bbox[0] + $bbox[2]/2) ($bbox[1] + $bbox[3]/2) "shape '$id'"
        $widthPx = $shape.CellsU('Width').ResultIU * $spec.reference.widthPx / $spec.page.widthIn
        $heightPx = $shape.CellsU('Height').ResultIU * $spec.reference.heightPx / $spec.page.heightIn
        if ($item.kind -ne 'line' -and ([math]::Abs($widthPx - $bbox[2]) -gt $TolerancePx -or [math]::Abs($heightPx - $bbox[3]) -gt $TolerancePx)) { $issues.Add("Size mismatch for '$id'.") }
        $expectedParent = if ($item.kind -eq 'group') { $id } else { [string]$item.parentId }
        Check-Parent $shape $expectedParent $id
    }
    foreach ($group in $scene.Groups) {
        $id = $group.id + '::group'
        if (-not $actual.ContainsKey($id)) { $issues.Add("Missing native group '$id'."); continue }
        Check-Parent $actual[$id] ([string]$group.parentId) $id
    }
    foreach ($edge in $scene.Connectors) {
        $id = [string]$edge.id
        if (-not $actual.ContainsKey($id)) { $issues.Add("Missing connector '$id'."); continue }
        $shape = $actual[$id]
        Check-Parent $shape ([string]$edge.parentId) $id
        foreach ($end in @('Begin','End')) {
            $targetId = if ($end -eq 'Begin') { [string]$edge.from } else { [string]$edge.to }
            $key = [string]$shape.ID + ':' + $end + 'X'
            if (-not $glue.ContainsKey($key) -or -not $actual.ContainsKey($targetId) -or $glue[$key].ID -ne $actual[$targetId].ID) { $issues.Add("Connector '$id' $end is not glued to '$targetId'.") }
            else {
                $property = if ($end -eq 'Begin') { 'fromConnection' } else { 'toConnection' }
                $connectionSpec = [string](Get-SpecValue $edge $property 'auto')
                if ($connectionSpec -like 'port:*') {
                    $port = @($scene.ById[$targetId].ports | Where-Object { $_.id -eq $connectionSpec.Substring(5) })[0]
                    $target = $actual[$targetId]
                    $expectedX = $target.CellsU('Width').ResultIU * [double]$port.x
                    $expectedY = $target.CellsU('Height').ResultIU * (1.0 - [double]$port.y)
                    $row = $glueCells[$key].Row
                    if ([math]::Abs($glueCells[$key].ResultIU - $expectedX) -gt 0.000001 -or [math]::Abs($target.CellsSRC(7,$row,1).ResultIU - $expectedY) -gt 0.000001) { $issues.Add("Wrong named port for '$id' $end.") }
                }
            }
        }
        $arrow = [string](Get-SpecValue $edge 'arrow' 'none')
        foreach ($end in @('Begin','End')) {
            $expectedArrow = $arrow -eq 'both' -or $arrow -eq $end.ToLowerInvariant()
            if (($shape.CellsU($end + 'Arrow').ResultIU -gt 0) -ne $expectedArrow) { $issues.Add("Arrow mismatch for '$id' $end.") }
        }
        $points = @(Get-SpecValue $edge 'waypointsPx' @())
        if ($points.Count -gt 0) {
            if ($shape.RowCount(11) -ne ($points.Count + 3)) { $issues.Add("Manual route vertex count changed for '$id'.") }
            else {
                for ($j = 0; $j -lt $points.Count; $j++) {
                    $p = Convert-VisioLocalPointToPage $shape $shape.CellsSRC(11, ($j+2), 0).ResultIU $shape.CellsSRC(11, ($j+2), 1).ResultIU
                    Check-Point $p $points[$j][0] $points[$j][1] "connector '$id' waypoint $j"
                }
                foreach ($end in @('Begin','End')) {
                    $r = if ($end -eq 'Begin') { 1 } else { $points.Count + 2 }
                    $vertex = Convert-VisioLocalPointToPage $shape $shape.CellsSRC(11,$r,0).ResultIU $shape.CellsSRC(11,$r,1).ResultIU
                    $parent = $shape.ContainingShape
                    $ex = $shape.CellsU($end+'X').ResultIU
                    $ey = $shape.CellsU($end+'Y').ResultIU
                    $endpoint = if ($parent -and $parent.ID -ne 0) { Convert-VisioLocalPointToPage $parent $ex $ey } else { [pscustomobject]@{ X=$ex; Y=$ey } }
                    Check-Point $vertex ($endpoint.X * $spec.reference.widthPx / $spec.page.widthIn) (($spec.page.heightIn-$endpoint.Y) * $spec.reference.heightPx / $spec.page.heightIn) "connector '$id' visible $end"
                }
            }
        }
        if ($edge.label) {
            $labelId = $id + '::label'
            if (-not $actual.ContainsKey($labelId)) { $issues.Add("Missing label '$labelId'.") }
            else {
                Check-Text $actual[$labelId] ([string]$edge.label.text) $labelId
                $b = $edge.label.bboxPx
                Check-Point (Get-VisioShapePageCenter $actual[$labelId]) ($b[0]+$b[2]/2) ($b[1]+$b[3]/2) "label '$labelId'"
            }
        }
    }
} finally {
    if ($doc) { try { $doc.Saved = $true; $doc.Close() } catch {} }
    if ($visio) { try { $visio.Quit() } catch {} }
}
$report = [ordered]@{ passed = ($issues.Count -eq 0); expectedShapes = $scene.Shapes.Count; expectedGroups = $scene.Groups.Count; expectedConnectors = $scene.Connectors.Count; actualTaggedObjects = $actual.Count; tolerancePx = $TolerancePx; issues = $issues.ToArray(); note = 'Checks spec-to-VSDX fidelity. Reference-image completeness and visual quality still require region-by-region comparison.' }
$json = $report | ConvertTo-Json -Depth 8
if ($ReportPath) { [IO.File]::WriteAllText([IO.Path]::GetFullPath($ReportPath), $json, (New-Object Text.UTF8Encoding($false))) }
$json
if (-not $report.passed) { exit 1 }
