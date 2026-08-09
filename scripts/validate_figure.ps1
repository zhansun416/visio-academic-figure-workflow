param(
    [Parameter(Mandatory=$true)]
    [string]$VsdxPath,

    [string]$SpecPath,

    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
$resolved = (Resolve-Path -LiteralPath $VsdxPath).Path
$spec = $null
if ($SpecPath) { $spec = Get-Content -LiteralPath (Resolve-Path -LiteralPath $SpecPath).Path -Raw -Encoding UTF8 | ConvertFrom-Json }

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead($resolved)
$media = @()
$shapeCount = 0
try {
    $media = @($zip.Entries | Where-Object { $_.FullName -like 'visio/media/*' } | Sort-Object FullName)
    $pageEntries = @($zip.Entries | Where-Object { $_.FullName -match '^visio/pages/page\d+\.xml$' })
    foreach ($pageEntry in $pageEntries) {
        $reader = [IO.StreamReader]::new($pageEntry.Open())
        $pageXml = $reader.ReadToEnd()
        $reader.Close()
        [xml]$xml = $pageXml
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace('v', 'http://schemas.microsoft.com/office/visio/2012/main')
        $shapeCount += $xml.SelectNodes('//v:Shape', $ns).Count
    }
} finally {
    $zip.Dispose()
}

$visio = New-Object -ComObject Visio.Application
$visio.Visible = $false
$doc = $null
$pageShapeCount = 0
$texts = New-Object System.Collections.Generic.List[string]
function Collect-ShapeText($Shape) {
    if ($Shape.Text) { $texts.Add(([string]$Shape.Text).Trim()) }
    for ($i = 1; $i -le $Shape.Shapes.Count; $i++) { Collect-ShapeText $Shape.Shapes.Item($i) }
}
try {
    $doc = $visio.Documents.Open($resolved)
    for ($pageIndex = 1; $pageIndex -le $doc.Pages.Count; $pageIndex++) {
        $page = $doc.Pages.Item($pageIndex)
        $pageShapeCount += $page.Shapes.Count
        for ($i = 1; $i -le $page.Shapes.Count; $i++) { Collect-ShapeText $page.Shapes.Item($i) }
    }
} finally {
    if ($doc -ne $null) {
        try { $doc.Saved = $true } catch {}
        try { $doc.Close() } catch {}
    }
    if ($visio -ne $null) { try { $visio.Quit() } catch {} }
}

$expectedTexts = @()
if ($spec) {
    $expectedTexts = @($spec.shapes | Where-Object { $_.text } | ForEach-Object { [string]$_.text })
}
$missingTexts = @($expectedTexts | Where-Object { $texts -notcontains $_ })
$rasterMedia = @($media | Where-Object { $_.FullName -match '\.(png|jpg|jpeg)$' })
$largeMedia = @($media | Where-Object { $_.Length -gt 1000000 })
$checks = [ordered]@{
    fileExists = (Test-Path -LiteralPath $resolved)
    fileBytes = (Get-Item -LiteralPath $resolved).Length
    packageShapeCount = $shapeCount
    pageShapeCount = $pageShapeCount
    mediaCount = $media.Count
    rasterMedia = @($rasterMedia | ForEach-Object { $_.FullName })
    largeMedia = @($largeMedia | ForEach-Object { $_.FullName })
    expectedTextCount = $expectedTexts.Count
    missingTexts = $missingTexts
    passed = ($shapeCount -gt 0 -and $pageShapeCount -gt 0 -and $missingTexts.Count -eq 0)
    notes = @('Raster media is reported for review; use validate_vsdx_output.ps1 -DisallowRasterMedia to enforce a raster-free deliverable.')
}
$json = $checks | ConvertTo-Json -Depth 10
if ($ReportPath) {
    $reportDir = Split-Path -Parent ([IO.Path]::GetFullPath($ReportPath))
    if ($reportDir -and -not (Test-Path -LiteralPath $reportDir)) { New-Item -ItemType Directory -Force -Path $reportDir | Out-Null }
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($ReportPath), $json, (New-Object Text.UTF8Encoding($false)))
}
Write-Output $json
if (-not $checks.passed) { exit 1 }
exit 0
