# fix_media_paths.ps1 -- relocate image parts that landed at the PACKAGE ROOT (media/) into word/media/
# and make their relationship targets relative, which is the conventional OPC layout.
#
# Why: officecli's picture insertion on some legacy documents writes word/media content as "media/xxx"
# with Target="/media/xxx". Word renders it fine, but LibreOffice, python-docx and strict package
# validators may not resolve a package-root part.
#
# Pure ASCII on purpose (PowerShell 5.1 reads a BOM-less UTF-8 script as ANSI).
param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [string]$Out = ""
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not $Out) { $Out = $Docx }
if (-not (Test-Path -LiteralPath $Docx)) { throw "docx not found: $Docx" }

$zip = [System.IO.Compression.ZipFile]::OpenRead($Docx)
$entries = @{}
$order = New-Object System.Collections.ArrayList
foreach ($e in $zip.Entries) {
    if ([string]::IsNullOrEmpty($e.Name)) { continue }
    $ms = New-Object System.IO.MemoryStream; $s = $e.Open(); $s.CopyTo($ms); $s.Close()
    $entries[$e.FullName] = $ms.ToArray(); [void]$order.Add($e.FullName)
}
$zip.Dispose()

# Parts sitting at the package root that belong under word/
$moved = @{}
$newnames = New-Object System.Collections.ArrayList
foreach ($k in $order) {
    if ($k -match '^media/(.+)$') {
        $target = 'word/media/' + $Matches[1]
        $moved[$k] = $target
        if (-not $entries.ContainsKey($target)) { $entries[$target] = $entries[$k] }
        [void]$newnames.Add($target)
    } else {
        [void]$newnames.Add($k)
    }
}
if ($moved.Count -eq 0) {
    [pscustomobject]@{ ok = $true; moved = 0; out = $Out; note = "nothing to fix" } | ConvertTo-Json -Compress
    exit 0
}

# Rewrite relationship targets: "/media/x" or "media/x" -> "media/x" (relative to word/)
$relsPath = 'word/_rels/document.xml.rels'
if ($entries.ContainsKey($relsPath)) {
    $rels = [System.Text.Encoding]::UTF8.GetString($entries[$relsPath])
    $rels = $rels -replace 'Target="/media/', 'Target="media/'
    $entries[$relsPath] = [System.Text.Encoding]::UTF8.GetBytes($rels)
}

$tmp = $Out + '.tmp'
if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force }
$fs = [IO.File]::Open($tmp, [IO.FileMode]::CreateNew)
$za = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($k in $newnames) {
        $ze = $za.CreateEntry($k, [System.IO.Compression.CompressionLevel]::Optimal)
        $os = $ze.Open(); $os.Write($entries[$k], 0, $entries[$k].Length); $os.Close()
    }
} finally { $za.Dispose(); $fs.Close() }
if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out -Force }
Move-Item -LiteralPath $tmp -Destination $Out

[pscustomobject]@{
    ok     = $true
    moved  = $moved.Count
    out    = $Out
    parts  = @($moved.Keys)
} | ConvertTo-Json -Compress
