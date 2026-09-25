function Set-VisioSpecIdentity($Shape, [string]$Id) {
    if ($Shape.SectionExists(242, 0) -eq 0) { [void]$Shape.AddSection(242) }
    if ($Shape.CellExistsU('User.SpecId', 0) -eq 0) { [void]$Shape.AddNamedRow(242, 'SpecId', 0) }
    $Shape.CellsU('User.SpecId').FormulaU = '"' + $Id.Replace('"','""') + '"'
}

function Add-VisioSpecPorts($Shape, $Ports) {
    $result = @{}
    foreach ($port in @($Ports)) {
        if ($null -eq $port) { continue }
        if ($Shape.SectionExists(7, 0) -eq 0) { [void]$Shape.AddSection(7) }
        # Numeric rows preserve compatibility with existing cardinal enumeration.
        $row = $Shape.AddRow(7, -2, 153)
        $x = ([double]$port.x).ToString([Globalization.CultureInfo]::InvariantCulture)
        $y = (1.0 - [double]$port.y).ToString([Globalization.CultureInfo]::InvariantCulture)
        $Shape.CellsSRC(7, $row, 0).FormulaU = "Width*$x"
        $Shape.CellsSRC(7, $row, 1).FormulaU = "Height*$y"
        $Shape.CellsSRC(7, $row, 2).FormulaU = '0'
        $Shape.CellsSRC(7, $row, 3).FormulaU = '0'
        $result[[string]$port.id] = 'Connections.X' + ($row + 1)
    }
    return $result
}

function Set-VisioManualRoute($Connector, $Waypoints, [double]$PageW, [double]$PageH, [double]$RefW, [double]$RefH) {
    $Connector.CellsU('ConFixedCode').FormulaU = '2'
    # Visio regenerates its inherited Geometry1 during grouping/glue updates.
    # Keep it hidden and use a separate native Geometry2 for authored vertices.
    $Connector.CellsU('Geometry1.NoShow').FormulaU = '1'
    [void]$Connector.AddSection(11)
    [void]$Connector.AddRow(11, 0, 137)
    [void]$Connector.AddRow(11, -2, 138)
    $Connector.CellsU('Geometry2.X1').FormulaU = '0'
    $Connector.CellsU('Geometry2.Y1').FormulaU = '0'
    foreach ($point in $Waypoints) {
        [double]$localX = 0
        [double]$localY = 0
        $Connector.XYFromPage(($PageW * [double]$point[0] / $RefW), ($PageH - $PageH * [double]$point[1] / $RefH), [ref]$localX, [ref]$localY)
        $row = $Connector.AddRow(11, -2, 139)
        $Connector.CellsSRC(11, $row, 0).FormulaU = $localX.ToString([Globalization.CultureInfo]::InvariantCulture) + ' in'
        $Connector.CellsSRC(11, $row, 1).FormulaU = $localY.ToString([Globalization.CultureInfo]::InvariantCulture) + ' in'
    }
    $row = $Connector.AddRow(11, -2, 139)
    $Connector.CellsSRC(11, $row, 0).FormulaU = 'Width'
    $Connector.CellsSRC(11, $row, 1).FormulaU = 'Height'
    $Connector.CellsU('Geometry2.NoFill').FormulaU = '1'
    $Connector.CellsU('Geometry2.NoShow').FormulaU = '0'
}
