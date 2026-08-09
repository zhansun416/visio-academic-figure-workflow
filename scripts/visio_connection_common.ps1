$ErrorActionPreference = 'Stop'

function Convert-VisioLocalPointToPage {
    param(
        [Parameter(Mandatory = $true)] $Shape,
        [Parameter(Mandatory = $true)] [double]$X,
        [Parameter(Mandatory = $true)] [double]$Y
    )

    [double]$pageX = 0
    [double]$pageY = 0
    try {
        $Shape.XYToPage($X, $Y, [ref]$pageX, [ref]$pageY)
    } catch {
        throw "Failed to transform local Visio coordinates for shape '$($Shape.NameU)': $($_.Exception.Message)"
    }
    return [pscustomobject]@{ X = $pageX; Y = $pageY }
}

function Get-VisioShapePageCenter {
    param([Parameter(Mandatory = $true)] $Shape)

    $localX = [double]$Shape.CellsU('LocPinX').ResultIU
    $localY = [double]$Shape.CellsU('LocPinY').ResultIU
    return Convert-VisioLocalPointToPage -Shape $Shape -X $localX -Y $localY
}

function Get-VisioConnectionPointCandidates {
    param(
        [Parameter(Mandatory = $true)] $Shape,
        [int]$MaximumConnectionPoints = 128
    )

    if ($MaximumConnectionPoints -le 0) { throw 'MaximumConnectionPoints must be positive.' }
    $points = New-Object System.Collections.Generic.List[object]
    $consecutiveMisses = 0
    for ($index = 1; $index -le $MaximumConnectionPoints; $index++) {
        $xName = "Connections.X$index"
        $yName = "Connections.Y$index"
        $hasX = ($Shape.CellExistsU($xName, 0) -ne 0)
        $hasY = ($Shape.CellExistsU($yName, 0) -ne 0)
        if (-not ($hasX -and $hasY)) {
            $consecutiveMisses++
            if ($index -gt 8 -and $consecutiveMisses -ge 8) { break }
            continue
        }

        $localX = [double]$Shape.CellsU($xName).ResultIU
        $localY = [double]$Shape.CellsU($yName).ResultIU
        $pagePoint = Convert-VisioLocalPointToPage -Shape $Shape -X $localX -Y $localY
        $points.Add([pscustomobject]@{
            Index = $index
            Name = "X$index"
            CellName = $xName
            LocalX = $localX
            LocalY = $localY
            PageX = [double]$pagePoint.X
            PageY = [double]$pagePoint.Y
            X = [double]$pagePoint.X
            Y = [double]$pagePoint.Y
        })
        $consecutiveMisses = 0
    }
    return $points.ToArray()
}

function Get-VisioCardinalConnectionName {
    param(
        [Parameter(Mandatory = $true)] $Shape,
        [Parameter(Mandatory = $true)]
        [ValidateSet('left','right','top','bottom')]
        [string]$Direction,
        [string]$ShapeId = '',
        [double]$CoordinateToleranceIn = 0.001
    )

    if ($CoordinateToleranceIn -le 0) { throw 'CoordinateToleranceIn must be positive.' }
    $points = @(Get-VisioConnectionPointCandidates -Shape $Shape)
    if ($points.Count -eq 0) {
        $label = if ($ShapeId) { $ShapeId } else { [string]$Shape.NameU }
        throw "Shape '$label' has no readable Connections.Xn/Yn cells. Use a connection-capable native Master."
    }

    $center = Get-VisioShapePageCenter -Shape $Shape
    $axisProperty = if ($Direction -in @('left','right')) { 'PageX' } else { 'PageY' }
    $extreme = if ($Direction -in @('right','top')) {
        [double](($points | Measure-Object -Property $axisProperty -Maximum).Maximum)
    } else {
        [double](($points | Measure-Object -Property $axisProperty -Minimum).Minimum)
    }
    $edgePoints = @($points | Where-Object { [math]::Abs(([double]$_.$axisProperty) - $extreme) -le $CoordinateToleranceIn })
    if ($edgePoints.Count -eq 0) { $edgePoints = $points }

    if ($Direction -in @('left','right')) {
        $selected = $edgePoints | Sort-Object `
            @{ Expression = { [math]::Abs($_.PageY - $center.Y) }; Descending = $false }, `
            @{ Expression = { [math]::Abs($_.PageX - $extreme) }; Descending = $false }, `
            @{ Expression = { $_.Index }; Descending = $false } | Select-Object -First 1
    } else {
        $selected = $edgePoints | Sort-Object `
            @{ Expression = { [math]::Abs($_.PageX - $center.X) }; Descending = $false }, `
            @{ Expression = { [math]::Abs($_.PageY - $extreme) }; Descending = $false }, `
            @{ Expression = { $_.Index }; Descending = $false } | Select-Object -First 1
    }
    return [string]$selected.Name
}

function Get-VisioAutoConnectionPair {
    param(
        [Parameter(Mandatory = $true)] $FromShape,
        [Parameter(Mandatory = $true)] $ToShape,
        [string]$FromShapeId = '',
        [string]$ToShapeId = '',
        [double]$CoordinateToleranceIn = 0.001
    )

    $fromCenter = Get-VisioShapePageCenter -Shape $FromShape
    $toCenter = Get-VisioShapePageCenter -Shape $ToShape
    $deltaX = $toCenter.X - $fromCenter.X
    $deltaY = $toCenter.Y - $fromCenter.Y

    if ([math]::Abs($deltaX) -ge [math]::Abs($deltaY)) {
        if ($deltaX -ge 0) { $fromDirection = 'right'; $toDirection = 'left' }
        else { $fromDirection = 'left'; $toDirection = 'right' }
    } else {
        if ($deltaY -ge 0) { $fromDirection = 'top'; $toDirection = 'bottom' }
        else { $fromDirection = 'bottom'; $toDirection = 'top' }
    }

    return [pscustomobject]@{
        FromConnection = Get-VisioCardinalConnectionName -Shape $FromShape -Direction $fromDirection -ShapeId $FromShapeId -CoordinateToleranceIn $CoordinateToleranceIn
        ToConnection = Get-VisioCardinalConnectionName -Shape $ToShape -Direction $toDirection -ShapeId $ToShapeId -CoordinateToleranceIn $CoordinateToleranceIn
        FromDirection = $fromDirection
        ToDirection = $toDirection
    }
}
