# Isolated COM edit phase used by test_dense_figures.ps1.
param([Parameter(Mandatory=$true)][string]$VsdxPath, [Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop'
$app=New-Object -ComObject Visio.Application
$app.Visible=$false
$doc=$null
try {
    $doc=$app.Documents.Open((Resolve-Path -LiteralPath $VsdxPath).Path)
    $page=$doc.Pages.Item(1)
    $group=$null
    for($i=1;$i -le $page.Shapes.Count;$i++) {
        $shape=$page.Shapes.Item($i)
        if($shape.CellExistsU('User.SpecId',0) -ne 0 -and $shape.CellsU('User.SpecId').ResultStr('') -eq 'results::group') { $group=$shape; break }
    }
    if(-not $group) { throw 'Missing results group in movement test.' }
    $x=[double]$group.CellsU('PinX').ResultIU+0.5
    $group.CellsU('PinX').FormulaU=$x.ToString([Globalization.CultureInfo]::InvariantCulture)+' in'
    [void]$doc.SaveAs([IO.Path]::GetFullPath($OutputPath))
} finally {
    if($doc) { try { $doc.Saved=$true; $doc.Close() } catch {} }
    if($app) { try { $app.Quit() } catch {} }
}
