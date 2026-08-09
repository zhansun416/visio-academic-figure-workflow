Set-StrictMode -Version 2.0

function Resolve-SvgLibraryRoot {
    param([string]$LibraryRoot)

    if ($LibraryRoot) { return [IO.Path]::GetFullPath($LibraryRoot) }
    if ($env:VISIO_FIGURE_SVG_LIBRARY) { return [IO.Path]::GetFullPath($env:VISIO_FIGURE_SVG_LIBRARY) }

    $documents = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
    if ([string]::IsNullOrWhiteSpace($documents)) {
        if ([string]::IsNullOrWhiteSpace($env:USERPROFILE)) { throw 'Cannot resolve the user Documents directory.' }
        $documents = Join-Path $env:USERPROFILE 'Documents'
    }
    return [IO.Path]::GetFullPath((Join-Path $documents 'Codex\svg-library'))
}

function Get-SvgLibraryPaths {
    param([Parameter(Mandatory=$true)][string]$LibraryRoot)

    $root = Resolve-SvgLibraryRoot $LibraryRoot
    return [pscustomobject]@{
        root = $root
        icons = Join-Path $root 'icons'
        manifest = Join-Path $root 'library.json'
        backup = Join-Path $root 'library.json.bak'
    }
}

function Initialize-SvgLibraryInternal {
    param([string]$LibraryRoot)

    $paths = Get-SvgLibraryPaths (Resolve-SvgLibraryRoot $LibraryRoot)
    New-Item -ItemType Directory -Force -Path $paths.icons | Out-Null
    if (-not (Test-Path -LiteralPath $paths.manifest)) {
        $json = ([ordered]@{ libraryVersion = 1; root = '.'; entries = @() } | ConvertTo-Json -Depth 6)
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        try {
            $stream = New-Object IO.FileStream($paths.manifest, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
            try {
                $bytes = $utf8.GetBytes($json)
                $stream.Write($bytes, 0, $bytes.Length)
            } finally {
                $stream.Dispose()
            }
        } catch [IO.IOException] {
            if (-not (Test-Path -LiteralPath $paths.manifest)) { throw }
        }
    }
    return $paths
}

function Get-SvgLibraryMutexName {
    param([Parameter(Mandatory=$true)][string]$LibraryRoot)

    $bytes = [Text.Encoding]::UTF8.GetBytes(([IO.Path]::GetFullPath($LibraryRoot)).ToLowerInvariant())
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '') }
    finally { $sha.Dispose() }
    return 'Local\VisioAcademicFigureSvgLibrary_' + $hash.Substring(0, 24)
}

function Invoke-WithSvgLibraryLock {
    param(
        [Parameter(Mandatory=$true)][string]$LibraryRoot,
        [Parameter(Mandatory=$true)][scriptblock]$Action,
        [int]$TimeoutSeconds = 30
    )

    $root = Resolve-SvgLibraryRoot $LibraryRoot
    $mutex = New-Object Threading.Mutex($false, (Get-SvgLibraryMutexName $root))
    $acquired = $false
    try {
        try { $acquired = $mutex.WaitOne([TimeSpan]::FromSeconds($TimeoutSeconds)) }
        catch [Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { throw "Timed out waiting for the SVG library lock: $root" }
        return & $Action
    } finally {
        if ($acquired) { try { $mutex.ReleaseMutex() } catch {} }
        $mutex.Dispose()
    }
}

function Read-SvgLibraryManifest {
    param([Parameter(Mandatory=$true)][string]$LibraryRoot)

    $paths = Initialize-SvgLibraryInternal $LibraryRoot
    try { $raw = Get-Content -LiteralPath $paths.manifest -Raw -Encoding UTF8 | ConvertFrom-Json }
    catch { throw "Invalid SVG library manifest '$($paths.manifest)': $($_.Exception.Message)" }

    $entries = if ($raw -is [System.Array]) {
        @($raw)
    } elseif ($null -ne $raw.PSObject.Properties['entries']) {
        @($raw.entries)
    } elseif ($null -ne $raw.PSObject.Properties['assets']) {
        @($raw.assets)
    } else {
        @()
    }
    return [pscustomobject]@{ paths = $paths; entries = $entries }
}

function Write-SvgLibraryManifestAtomic {
    param(
        [Parameter(Mandatory=$true)][string]$LibraryRoot,
        [Parameter(Mandatory=$true)][object[]]$Entries
    )

    $paths = Initialize-SvgLibraryInternal $LibraryRoot
    $library = [ordered]@{
        libraryVersion = 1
        root = '.'
        updatedAt = (Get-Date).ToString('s')
        entries = @($Entries | Sort-Object file)
    }
    $json = $library | ConvertTo-Json -Depth 12
    $temp = Join-Path $paths.root ('.library-' + [Guid]::NewGuid().ToString('N') + '.tmp')
    [IO.File]::WriteAllText($temp, $json, (New-Object Text.UTF8Encoding($false)))
    try {
        if (Test-Path -LiteralPath $paths.manifest) {
            [IO.File]::Replace($temp, $paths.manifest, $paths.backup, $true)
        } else {
            [IO.File]::Move($temp, $paths.manifest)
        }
    } finally {
        if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force }
    }
    return $paths
}

function Get-SvgTagsFromName {
    param([Parameter(Mandatory=$true)][string]$FileName)

    $stem = [IO.Path]::GetFileNameWithoutExtension($FileName)
    return @([regex]::Split($stem, '[^\p{L}\p{Nd}]+') | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() })
}

function Resolve-SvgEntryPath {
    param(
        [Parameter(Mandatory=$true)][string]$LibraryRoot,
        [Parameter(Mandatory=$true)]$Entry
    )

    $path = [string]$Entry.path
    if ([string]::IsNullOrWhiteSpace($path)) { $path = Join-Path 'icons' ([string]$Entry.file) }
    if ([IO.Path]::IsPathRooted($path)) { return [IO.Path]::GetFullPath($path) }
    return [IO.Path]::GetFullPath((Join-Path (Resolve-SvgLibraryRoot $LibraryRoot) $path))
}

function Get-SvgLibraryState {
    param([string]$LibraryRoot)

    $manifest = Read-SvgLibraryManifest (Resolve-SvgLibraryRoot $LibraryRoot)
    $files = @(Get-ChildItem -LiteralPath $manifest.paths.icons -Filter '*.svg' -File | Sort-Object Name)
    $entryByFile = @{}
    foreach ($entry in @($manifest.entries)) {
        if ($entry.file) { $entryByFile[[string]$entry.file] = $entry }
    }

    $unindexed = @($files | Where-Object { -not $entryByFile.ContainsKey($_.Name) } | ForEach-Object { $_.Name })
    $missing = New-Object System.Collections.Generic.List[string]
    $hashMismatch = New-Object System.Collections.Generic.List[string]
    foreach ($entry in @($manifest.entries)) {
        $fullPath = Resolve-SvgEntryPath $manifest.paths.root $entry
        if (-not (Test-Path -LiteralPath $fullPath)) {
            $missing.Add([string]$entry.file)
            continue
        }
        if ($entry.sha256) {
            $actual = (Get-FileHash -LiteralPath $fullPath -Algorithm SHA256).Hash.ToLowerInvariant()
            if ($actual -ne ([string]$entry.sha256).ToLowerInvariant()) { $hashMismatch.Add([string]$entry.file) }
        }
    }

    return [pscustomobject]@{
        root = $manifest.paths.root
        manifest = $manifest.paths.manifest
        iconsDirectory = $manifest.paths.icons
        svgFileCount = $files.Count
        manifestEntryCount = @($manifest.entries).Count
        unindexedFiles = @($unindexed)
        missingFiles = @($missing)
        hashMismatchFiles = @($hashMismatch)
        entries = @($manifest.entries)
        consistent = ($unindexed.Count -eq 0 -and $missing.Count -eq 0 -and $hashMismatch.Count -eq 0)
    }
}

function Get-SvgFileAssessment {
    param([Parameter(Mandatory=$true)][string]$SvgPath)

    $resolved = (Resolve-Path -LiteralPath $SvgPath).Path
    $warnings = New-Object System.Collections.Generic.List[string]
    $external = New-Object System.Collections.Generic.List[string]
    $risky = New-Object System.Collections.Generic.List[string]
    $validXml = $false
    $root = $null
    $xml = $null
    $raw = ''
    try {
        $raw = Get-Content -LiteralPath $resolved -Raw -Encoding UTF8
        $xml = [xml]$raw
        $root = $xml.DocumentElement
        if ($null -eq $root -or $root.LocalName -ne 'svg') { throw 'Root element is not svg.' }
        $validXml = $true
        if (-not $root.GetAttribute('xmlns')) { $warnings.Add('Missing xmlns attribute.') }
        if (-not $root.GetAttribute('viewBox')) { $warnings.Add('Missing viewBox attribute.') }

        $ids = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($node in @($xml.SelectNodes('//*[@id]'))) { [void]$ids.Add([string]$node.GetAttribute('id')) }
        foreach ($node in @($xml.SelectNodes('//*[local-name()="image" or local-name()="use" or local-name()="script"]'))) {
            if ($node.LocalName -eq 'script') { $external.Add('script element'); continue }
            $href = $node.GetAttribute('href')
            if (-not $href) { $href = $node.GetAttribute('xlink:href') }
            if ($href -match '^#(.+)$') {
                if (-not $ids.Contains($matches[1])) { $external.Add('unresolved local reference: ' + $href) }
            } elseif ($href -and $href -notmatch '^data:') {
                $external.Add($href)
            }
        }
        foreach ($name in @('filter','mask','foreignObject','animate','animateMotion','animateTransform','set')) {
            if (@($xml.SelectNodes("//*[local-name()='$name']")).Count -gt 0) { $risky.Add($name) }
        }
        if ($external.Count -gt 0) { $warnings.Add('External, scripted, or unresolved asset references were found.') }
        if ($risky.Count -gt 0) { $warnings.Add('Visio compatibility-risk elements were found: ' + (@($risky) -join ', ') + '.') }
    } catch {
        $warnings.Add($_.Exception.Message)
    }

    return [pscustomobject]@{
        path = $resolved
        raw = $raw
        xml = if ($validXml) { $xml } else { $null }
        validXml = $validXml
        root = if ($root) { $root.LocalName } else { $null }
        xmlns = if ($root) { $root.GetAttribute('xmlns') } else { $null }
        viewBox = if ($root) { $root.GetAttribute('viewBox') } else { $null }
        externalReferences = @($external)
        compatibilityRisks = @($risky)
        warnings = @($warnings)
        readyForVisioImport = ($validXml -and $warnings.Count -eq 0)
    }
}
