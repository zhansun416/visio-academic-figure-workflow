# Pure spec normalization: no Visio or SVG library required.
function Get-SpecValue($Item, [string]$Name, $Default) {
    if ($null -eq $Item) { return $Default }
    $property = $Item.PSObject.Properties[$Name]
    if ($property -and $null -ne $property.Value) { return $property.Value }
    return $Default
}

function Assert-FiniteNumbers($Values, [int]$Count, [string]$Context) {
    if (@($Values).Count -ne $Count) { throw "$Context requires $Count numbers." }
    foreach ($value in $Values) {
        if ($null -eq $value -or $value -is [string] -or $value -is [bool]) { throw "$Context requires numeric values." }
        try { $number = [double]$value } catch { throw "$Context requires numeric values." }
        if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) { throw "$Context requires finite numbers." }
    }
}

function Expand-FigureSpec($Spec) {
    Assert-FiniteNumbers @($Spec.reference.widthPx, $Spec.reference.heightPx, $Spec.page.widthIn, $Spec.page.heightIn) 4 'Reference and page dimensions'
    if ($Spec.reference.widthPx -le 0 -or $Spec.reference.heightPx -le 0 -or $Spec.page.widthIn -le 0 -or $Spec.page.heightIn -le 0) { throw 'Reference and page dimensions must be positive.' }
    $flat = New-Object 'System.Collections.Generic.List[object]'
    $groups = New-Object 'System.Collections.Generic.List[object]'
    $byId = @{}
    $parents = @{}
    $warnings = New-Object 'System.Collections.Generic.List[string]'
    $kinds = @('rect','oval','circle','diamond','native','text','line','svg-asset','group')
    function Expand-Items($Items, [double]$OffsetX, [double]$OffsetY, [string]$ParentId, [int]$Depth) {
        foreach ($item in @($Items)) {
            if ($null -eq $item) { continue }
            $id = [string]$item.id
            if ([string]::IsNullOrWhiteSpace($id) -or $id.Contains('::')) { throw 'Shape ids must be non-empty and must not contain reserved ::.' }
            if ($byId.ContainsKey($id)) { throw "Duplicate figure id: $id" }
            $kind = ([string]$item.kind).ToLowerInvariant()
            if ($kinds -notcontains $kind) { throw "Unsupported shape kind '$kind' in '$id'." }
            Assert-FiniteNumbers $item.bboxPx 4 "Shape '$id' bboxPx"
            if ($item.bboxPx[2] -lt 0 -or $item.bboxPx[3] -lt 0 -or ($kind -ne 'line' -and ($item.bboxPx[2] -eq 0 -or $item.bboxPx[3] -eq 0))) { throw "Shape '$id' has invalid bbox dimensions." }
            if ($kind -eq 'line' -and $item.bboxPx[2] -eq 0 -and $item.bboxPx[3] -eq 0) { throw "Line '$id' must have nonzero length." }
            if ($kind -eq 'svg-asset' -and -not $item.assetRef) { throw "SVG shape '$id' requires assetRef." }
            if ($kind -eq 'group' -and [double](Get-SpecValue $item 'angleDeg' 0) -ne 0) { throw "Group '$id': rotate individual children; group rotation would invalidate reference coordinates." }
            $copy = [ordered]@{}
            foreach ($prop in $item.PSObject.Properties) { $copy[$prop.Name] = $prop.Value }
            $copy.kind = $kind
            $copy.bboxPx = @(([double]$item.bboxPx[0] + $OffsetX), ([double]$item.bboxPx[1] + $OffsetY), [double]$item.bboxPx[2], [double]$item.bboxPx[3])
            $copy.parentId = $ParentId
            $normalized = [pscustomobject]$copy
            $byId[$id] = $normalized
            $parents[$id] = $ParentId
            $flat.Add($normalized)
            $portIds = @{}
            foreach ($port in @(Get-SpecValue $item 'ports' @())) {
                $portId = [string]$port.id
                if ($portId -notmatch '^[A-Za-z][A-Za-z0-9_-]*$' -or $portIds.ContainsKey($portId)) { throw "Shape '$id' has an invalid or duplicate port id '$portId'." }
                Assert-FiniteNumbers @($port.x, $port.y) 2 "Port '$id/$portId'"
                if ($port.x -lt 0 -or $port.x -gt 1 -or $port.y -lt 0 -or $port.y -gt 1) { throw "Port '$id/$portId' must lie in normalized [0,1] bounds." }
                $portIds[$portId] = $true
            }
            if ($ParentId) {
                $parent = $byId[$ParentId]
                if ($item.bboxPx[0] -lt 0 -or $item.bboxPx[1] -lt 0 -or ($item.bboxPx[0] + $item.bboxPx[2]) -gt $parent.bboxPx[2] -or ($item.bboxPx[1] + $item.bboxPx[3]) -gt $parent.bboxPx[3]) { $warnings.Add("Shape '$id' extends outside group '$ParentId'; inspect intentional protrusions.") }
            }
            if ($kind -eq 'group') {
                $groups.Add([pscustomobject]@{ id = $id; parentId = $ParentId; depth = $Depth })
                Expand-Items (Get-SpecValue $item 'children' @()) $copy.bboxPx[0] $copy.bboxPx[1] $id ($Depth + 1)
            } elseif (@(Get-SpecValue $item 'children' @()).Count -gt 0) { throw "Only group shapes may have children: '$id'." }
        }
    }
    Expand-Items $Spec.shapes 0 0 '' 0
    if ($flat.Count -eq 0) { throw 'Spec must contain at least one shape.' }
    $edges = New-Object 'System.Collections.Generic.List[object]'
    foreach ($edge in @(Get-SpecValue $Spec 'connectors' @())) {
        $id = [string]$edge.id
        if ([string]::IsNullOrWhiteSpace($id) -or $id.Contains('::') -or $byId.ContainsKey($id)) { throw "Invalid or duplicate connector id: '$id'." }
        foreach ($end in @('from','to')) {
            $node = [string]$edge.$end
            if (-not $byId.ContainsKey($node) -or $byId[$node].kind -notin @('rect','oval','circle','diamond','native','group')) { throw "Connector '$id' $end must reference a native node or group: '$node'." }
            $connection = [string](Get-SpecValue $edge ($end + 'Connection') 'auto')
            if ($connection -like 'port:*') {
                $portName = $connection.Substring(5)
                if (-not (@(Get-SpecValue $byId[$node] 'ports' @()) | Where-Object { $_.id -eq $portName })) { throw "Connector '$id' references missing port '$node/$portName'." }
            } elseif ($connection -notmatch '^(auto|left|right|top|bottom|(Connections\.)?X[1-9][0-9]*)$') { throw "Connector '$id' has invalid endpoint '$connection'." }
        }
        $route = [string](Get-SpecValue $edge 'routing' 'auto')
        if ($route -notin @('auto','orthogonal','straight','manual')) { throw "Connector '$id' has unsupported routing '$route'." }
        $waypoints = @(Get-SpecValue $edge 'waypointsPx' @())
        foreach ($point in $waypoints) { Assert-FiniteNumbers $point 2 "Connector '$id' waypoint" }
        if ($route -eq 'manual' -and $waypoints.Count -eq 0) { throw "Manual connector '$id' needs waypointsPx." }
        if ($waypoints.Count -gt 0 -and $route -notin @('auto','manual')) { throw "Connector '$id': waypointsPx cannot be combined with '$route'." }
        if ($edge.from -eq $edge.to -and ($waypoints.Count -eq 0 -or (Get-SpecValue $edge 'fromConnection' 'auto') -eq 'auto' -or (Get-SpecValue $edge 'toConnection' 'auto') -eq 'auto')) { throw "Self-loop '$id' requires explicit endpoints and waypointsPx." }
        if (Get-SpecValue $edge 'label' $null) {
            Assert-FiniteNumbers $edge.label.bboxPx 4 "Connector '$id' label bboxPx"
            if ($edge.label.bboxPx[2] -le 0 -or $edge.label.bboxPx[3] -le 0) { throw "Connector '$id' label requires positive dimensions." }
        }
        # Lowest common container owns an edge. A group's frame belongs to that group.
        $ancestors = @{}
        $cursor = [string]$edge.from
        if ($byId[$cursor].kind -ne 'group') { $cursor = $parents[$cursor] }
        while ($cursor) { $ancestors[$cursor] = $true; $cursor = $parents[$cursor] }
        $cursor = [string]$edge.to
        if ($byId[$cursor].kind -ne 'group') { $cursor = $parents[$cursor] }
        while ($cursor -and -not $ancestors.ContainsKey($cursor)) { $cursor = $parents[$cursor] }
        $copy = [ordered]@{}
        foreach ($prop in $edge.PSObject.Properties) { $copy[$prop.Name] = $prop.Value }
        $copy.parentId = $cursor
        $normalized = [pscustomobject]$copy
        $edges.Add($normalized)
        $byId[$id] = $normalized
    }
    return [pscustomobject]@{ Shapes = $flat.ToArray(); Groups = $groups.ToArray(); Connectors = $edges.ToArray(); ById = $byId; Warnings = $warnings.ToArray() }
}
