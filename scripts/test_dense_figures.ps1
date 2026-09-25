param([Parameter(Mandatory=$true)][string]$OutputDirectory, [switch]$SkipVisio)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'figure_spec_common.ps1')
$root=[IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $root | Out-Null
$fixture=Join-Path (Split-Path $PSScriptRoot -Parent) 'assets/templates/dense_figure_spec.json'
$spec=Get-Content -LiteralPath $fixture -Raw -Encoding UTF8 | ConvertFrom-Json
$scene=Expand-FigureSpec $spec
if ($scene.Shapes.Count -ne 59 -or $scene.Groups.Count -ne 5 -or $scene.Connectors.Count -ne 15) { throw 'Dense fixture object counts changed unexpectedly.' }
if ($scene.ById['core-1-0'].bboxPx[0] -ne 615 -or $scene.ById['core-1-0'].bboxPx[1] -ne 550) { throw 'Nested coordinates were not translated correctly.' }
if ($scene.ById['stage-transfer'].parentId -ne 'model' -or $scene.ById['accept'].parentId -ne '' -or $scene.ById['refresh'].parentId -ne 'results') { throw 'Incorrect connector group ownership.' }
$rejected=0
foreach($test in @('duplicate','port','self-loop','endpoint','coordinate')) {
    $bad=Get-Content -LiteralPath $fixture -Raw -Encoding UTF8 | ConvertFrom-Json
    switch($test){
        'duplicate' { $bad.shapes[1].children[0].id='title' }
        'port' { $bad.connectors[0] | Add-Member fromConnection 'port:missing' }
        'self-loop' { $bad.connectors[0].to=$bad.connectors[0].from }
        'endpoint' { $bad.connectors[0].to='title' }
        'coordinate' { $bad.shapes[0].bboxPx[0]='not-a-number' }
    }
    $failed=$false
    try { $null=Expand-FigureSpec $bad } catch { $failed=$true }
    if(-not $failed){throw "Preflight accepted invalid $test fixture."}
    $rejected++
}
$validation=$null
$negative=$null
$movement=$null
if(-not $SkipVisio){
    $vsdx=Join-Path $root 'dense.vsdx'
    $preview=Join-Path $root 'dense.png'
    # Keep renderer COM proxies in a separate process so they cannot retain a
    # closed Visio instance when this test later opens the file for editing.
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'render_figure.ps1') -SpecPath $fixture -OutputPath $vsdx -PreviewPath $preview | Out-Null
    if($LASTEXITCODE -ne 0){throw 'Dense fixture rendering failed.'}
    $result=& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'validate_scene_spec.ps1') -VsdxPath $vsdx -SpecPath $fixture -ReportPath (Join-Path $root 'scene-validation.json')
    if($LASTEXITCODE -ne 0){throw ($result -join "`n")}
    $validation=$result | ConvertFrom-Json
    $outputCheck=& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'validate_vsdx_output.ps1') -VsdxPath $vsdx -RequireGluedConnectors -DisallowRasterMedia -MinimumFontSizePt 8
    if($LASTEXITCODE -ne 0){throw ($outputCheck -join "`n")}

    # Move a native group, save/reopen, and check that its details and internal
    # loop travel together while the cross-group edge remains attached.
    $movedSpec=Get-Content -LiteralPath $fixture -Raw -Encoding UTF8 | ConvertFrom-Json
    ($movedSpec.shapes | Where-Object { $_.id -eq 'results' }).bboxPx[0] += 50
    foreach($point in ($movedSpec.connectors | Where-Object { $_.id -eq 'refresh' }).waypointsPx){$point[0] += 50}
    $movedSpecPath=Join-Path $root 'moved-spec.json'
    $movedSpec | ConvertTo-Json -Depth 40 | Set-Content -LiteralPath $movedSpecPath -Encoding UTF8
    $movedPath=Join-Path $root 'moved.vsdx'
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'test_move_fixture_group.ps1') -VsdxPath $vsdx -OutputPath $movedPath | Out-Null
    if($LASTEXITCODE -ne 0){throw 'Group movement edit phase failed.'}
    $result=& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'validate_scene_spec.ps1') -VsdxPath $movedPath -SpecPath $movedSpecPath
    if($LASTEXITCODE -ne 0){throw ($result -join "`n")}
    $movement=$result | ConvertFrom-Json

    # A save cannot pass just because some text/geometry exists: intentionally
    # changed source expectations must be caught in the reopened document.
    $spec.shapes[1].children[0].text='Different source title'
    $spec.connectors[7].waypointsPx[0][0] += 25
    $badPath=Join-Path $root 'changed-expectations.json'
    $spec | ConvertTo-Json -Depth 40 | Set-Content -LiteralPath $badPath -Encoding UTF8
    $result=& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'validate_scene_spec.ps1') -VsdxPath $vsdx -SpecPath $badPath
    if($LASTEXITCODE -eq 0){throw 'Scene validation failed to reject changed text and route expectations.'}
    $negative=$result | ConvertFrom-Json
    if(-not ($negative.issues -match 'Text mismatch') -or -not ($negative.issues -match 'waypoint')){throw 'Negative validation did not detect both changed text and changed route.'}
}
[ordered]@{passed=$true;preflightRejectedCases=$rejected;visioTested=(-not $SkipVisio);sceneValidation=$validation;groupMovementPassed=if($movement){$movement.passed}else{$null};negativeChecks=if($negative){$negative.issues}else{@()};artifacts=$root} | ConvertTo-Json -Depth 10
