# office_convert.ps1 -- MS Office COM batch format converter (doc/rtf/xls/ppt -> modern).
# Invoked as a child process by the Office Toolkit app (tool: office_convert, route: /convert).
# Prints one result JSON object to stdout.
#
# NOTE: this file is intentionally ASCII-only. It may be run by Windows PowerShell 5.1
# (which decodes BOM-less UTF-8 as ANSI), so any non-ASCII character would corrupt parsing.
#
# Job JSON: {
#   "input":  "<absolute path of file or directory>",
#   "output": "<output dir, may be empty: empty = alongside each source>",
#   "mode":   "keep|backup|replace",
#   "recursive": true,
#   "keepTimestamps": true
# }
# Enumeration happens inside this script: the app process is confined by the Node
# permission model and must not touch user directories directly.

param(
    [Parameter(Mandatory = $true)][string]$JobFile
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# The app child process can have a lean environment; without PATHEXT PowerShell will not treat
# ".exe" as an executable and the call operator fails with "cannot run a document".
if (-not $env:PATHEXT) { $env:PATHEXT = ".COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC" }

$job = Get-Content -LiteralPath $JobFile -Raw -Encoding UTF8 | ConvertFrom-Json
$inputPath = [string]$job.input
$outDir = if ($job.output) { [string]$job.output } else { "" }
$mode = if ($job.mode) { [string]$job.mode } else { "keep" }
$recursive = if ($null -ne $job.recursive) { [bool]$job.recursive } else { $true }
$keepTs = [bool]$job.keepTimestamps
$enginePref = if ($job.engine) { [string]$job.engine } else { "auto" }
# Optional explicit target format (pdf, html, csv, txt, doc, xls, ppt, odt, ods, odp, rtf, epub, md).
# Empty means the legacy default mapping (doc/rtf -> docx, xls -> xlsx, ppt -> pptx).
$toFormat = if ($job.to) { ([string]$job.to).TrimStart('.').ToLower() } else { "" }

$legacyMap = @{
    ".doc" = "doc"; ".rtf" = "doc"
    ".xls" = "xls"
    ".ppt" = "ppt"
}

# Inputs LibreOffice can read, grouped into families. Used only when an explicit target is given.
$familyOf = @{
    ".doc" = "doc"; ".rtf" = "doc"; ".docx" = "doc"; ".odt" = "doc"; ".txt" = "doc"; ".html" = "doc"; ".htm" = "doc"; ".wps" = "doc"; ".wpd" = "doc"; ".md" = "doc"
    ".xls" = "xls"; ".xlsx" = "xls"; ".ods" = "xls"; ".csv" = "xls"; ".tsv" = "xls"
    ".ppt" = "ppt"; ".pptx" = "ppt"; ".odp" = "ppt"
}
$defaultTarget = @{ "doc" = "docx"; "xls" = "xlsx"; "ppt" = "pptx" }

# Authoritative LibreOffice output filter names (LibreOffice help: "File Conversion Filter Names").
function Get-LoFilter($family, $toExt) {
    switch ($toExt) {
        "pdf"  { if ($family -eq "xls") { return "calc_pdf_Export" } elseif ($family -eq "ppt") { return "impress_pdf_Export" } else { return "writer_pdf_Export" } }
        "docx" { return "MS Word 2007 XML" }
        "doc"  { return "MS Word 97" }
        "odt"  { return "writer8" }
        "rtf"  { return "Rich Text Format" }
        "txt"  { return "Text (encoded):UTF8" }
        "html" { if ($family -eq "xls") { return "HTML (StarCalc)" } else { return "HTML (StarWriter)" } }
        "epub" { return "EPUB" }
        "md"   { return "Markdown" }
        "xlsx" { return "Calc MS Excel 2007 XML" }
        "xls"  { return "MS Excel 97" }
        "ods"  { return "calc8" }
        "csv"  { return "Text - txt - csv (StarCalc):44,34,76" }
        "pptx" { return "Impress MS PowerPoint 2007 XML" }
        "ppt"  { return "MS PowerPoint 97" }
        "odp"  { return "impress8" }
        default { return "" }
    }
}

$results = New-Object System.Collections.ArrayList

function Write-ProgressJson($phase, $current, $total, $file, $officePid) {
    $fileB64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$file))
    if (-not $officePid) { $officePid = 0 }
    $payload = [pscustomobject]@{ phase = $phase; current = $current; total = $total; fileB64 = $fileB64; officePid = [int]$officePid } | ConvertTo-Json -Compress
    [Console]::Error.WriteLine("OFFICE_PROGRESS:$payload")
}

# The Office GUI process is launched by COM outside this process tree, so taskkill /T on
# PowerShell cannot reach it. Report the PID of the instance we caused to start so the app
# can terminate exactly that process on cancel/stall (never the user's own pre-existing Office).
function Get-NewOfficePid($processName, $beforeIds) {
    $candidate = Get-Process -Name $processName -ErrorAction SilentlyContinue |
        Where-Object { $beforeIds -notcontains $_.Id } |
        Sort-Object StartTime | Select-Object -First 1
    if ($candidate) { return [int]$candidate.Id }
    return 0
}

function Write-Result($obj) {
    $obj | ConvertTo-Json -Depth 6 -Compress
}

function Set-DestTimes($dest, $src) {
    if (-not $keepTs) { return }
    try {
        $s = Get-Item -LiteralPath $src
        $d = Get-Item -LiteralPath $dest
        $d.CreationTime = $s.CreationTime
        $d.LastWriteTime = $s.LastWriteTime
        $d.LastAccessTime = $s.LastAccessTime
    } catch {}
}

function Handle-Source($src, $backupDir) {
    switch ($mode) {
        "backup" {
            if ($backupDir) {
                $relative = $src.Substring($srcRoot.TrimEnd('\\').Length).TrimStart('\\')
                $backupTarget = Join-Path $backupDir $relative
                $backupParent = Split-Path -Parent $backupTarget
                if (-not (Test-Path -LiteralPath $backupParent)) { New-Item -ItemType Directory -Path $backupParent -Force | Out-Null }
                Move-Item -LiteralPath $src -Destination $backupTarget
            }
        }
        "replace" { Remove-Item -LiteralPath $src -Force }
        default { } # keep
    }
}

if (-not (Test-Path -LiteralPath $inputPath)) {
    Write-Result ([pscustomobject]@{ engine = "ms-office"; error = "input not found: $inputPath"; results = @() })
    exit 0
}

$isDir = (Get-Item -LiteralPath $inputPath).PSIsContainer
$srcRoot = if ($isDir) { $inputPath } else { Split-Path -Parent $inputPath }

$items = New-Object System.Collections.ArrayList
function Add-ConvertItem($src, $family) {
    $toExt = if ($toFormat) { $toFormat } else { $defaultTarget[$family] }
    if (-not $toExt) { return }
    if (([System.IO.Path]::GetExtension($src)).TrimStart('.').ToLower() -eq $toExt) { return }
    [void]$items.Add([pscustomobject]@{ src = $src; type = $family; to = $toExt; filter = (Get-LoFilter $family $toExt) })
}

if ($isDir) {
    $files = Get-ChildItem -LiteralPath $inputPath -File -Recurse:$recursive
    foreach ($f in $files) {
        $ext = $f.Extension.ToLower()
        if ($toFormat) {
            if ($familyOf.ContainsKey($ext)) { Add-ConvertItem $f.FullName $familyOf[$ext] }
        } elseif ($legacyMap.ContainsKey($ext)) {
            Add-ConvertItem $f.FullName $legacyMap[$ext]
        }
    }
} else {
    $ext = ([System.IO.Path]::GetExtension($inputPath)).ToLower()
    if ($toFormat) {
        if (-not $familyOf.ContainsKey($ext)) {
            Write-Result ([pscustomobject]@{ engine = "libreoffice"; error = "Unsupported input for target '$toFormat': $inputPath"; results = @() })
            exit 0
        }
        Add-ConvertItem $inputPath $familyOf[$ext]
    } else {
        if (-not $legacyMap.ContainsKey($ext)) {
            Write-Result ([pscustomobject]@{ engine = "libreoffice"; error = "not a legacy format (.doc/.rtf/.xls/.ppt): $inputPath"; results = @() })
            exit 0
        }
        Add-ConvertItem $inputPath $legacyMap[$ext]
    }
}

function Get-DestPath($src, $newExt) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($src)
    if ($outDir) {
        if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
        return (Join-Path $outDir ($base + $newExt))
    }
    return ([System.IO.Path]::ChangeExtension($src, $newExt))
}

$backupDir = ""
if ($mode -eq "backup") {
    $bkRoot = if ($outDir) { $outDir } else { $srcRoot }
    $backupDir = Join-Path $bkRoot "old-format-files"
}

$totalItems = $items.Count
if ($totalItems -eq 0) {
    Write-Result ([pscustomobject]@{ engine = "ms-office"; error = "No legacy Office files found."; results = @() })
    exit 0
}
# One-shot diagnostics: where does the child think TEMP/USERPROFILE point, and can it write there?
$tempProbe = "n/a"
try {
    $probeFile = Join-Path $env:TEMP ("officecli_probe_" + [Guid]::NewGuid().ToString() + ".tmp")
    Set-Content -LiteralPath $probeFile -Value "ok" -Encoding ASCII
    Remove-Item -LiteralPath $probeFile -Force
    $tempProbe = "writable"
} catch { $tempProbe = "FAILED: " + $_.Exception.Message }
$envInfo = "TEMP=$($env:TEMP) | USERPROFILE=$($env:USERPROFILE) | SESSIONNAME=$($env:SESSIONNAME) | ProgramFiles=$($env:ProgramFiles) | probe=$tempProbe"
Write-ProgressJson "env" 0 $totalItems $envInfo 0
Write-ProgressJson "enumerated" 0 $totalItems "" 0
$script:progressCurrent = 0
$script:officePid = 0
$script:engineName = "none"

function Get-Pids($names) {
    $ids = @()
    foreach ($n in $names) { $ids += @(Get-Process -Name $n -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) }
    return $ids
}

function Get-NewPid($names, $beforeIds) {
    foreach ($n in $names) {
        $c = Get-Process -Name $n -ErrorAction SilentlyContinue |
            Where-Object { $beforeIds -notcontains $_.Id } |
            Sort-Object StartTime | Select-Object -First 1
        if ($c) { return [int]$c.Id }
    }
    return 0
}

# Engine preference (from job.engine): "auto" tries WPS first then MS Office;
# "wps" forces WPS; "office" forces Microsoft Office.
function New-OfficeApp($wpsProgIds, $msProgIds) {
    $order = switch ($enginePref) {
        "wps" { $wpsProgIds }
        "office" { $msProgIds }
        default { @($wpsProgIds) + @($msProgIds) }
    }
    foreach ($prog in $order) {
        try {
            $app = New-Object -ComObject $prog -ErrorAction Stop
            $script:engineName = if ($prog -like "K*") { "wps" } else { "ms-office" }
            return $app
        } catch { }
    }
    return $null
}

function Convert-Word($list) {
    if ($list.Count -eq 0) { return }
    $word = $null
    try {
        $beforeIds = Get-Pids @("wps", "WINWORD")
        $word = New-OfficeApp @("KWPS.Application") @("Word.Application")
        if (-not $word) { throw "No Word automation engine available (WPS or MS Office)." }
        try { $word.Visible = $false } catch { }
        try { $word.DisplayAlerts = 0 } catch { }
        try { $word.AutomationSecurity = 3 } catch { }
        try { $word.Options.ConfirmConversions = $false } catch { }
        $script:officePid = Get-NewPid @("wps", "WINWORD") $beforeIds
        foreach ($it in $list) {
            $src = $it.src
            $script:progressCurrent++
            Write-ProgressJson "word" $script:progressCurrent $totalItems $src $script:officePid
            $dest = Get-DestPath $src $(if ($script:comPdf) { ".pdf" } else { ".docx" })
            try {
                if (Test-Path -LiteralPath $dest) { throw "Destination already exists: $dest" }
                Write-ProgressJson "word-open" $script:progressCurrent $totalItems $src $script:officePid
                $doc = $word.Documents.Open($src, $false, $true, $false)
                Write-ProgressJson "word-opened" $script:progressCurrent $totalItems $src $script:officePid
                Write-ProgressJson "word-saving" $script:progressCurrent $totalItems $src $script:officePid
                if ($script:comPdf) {
                    # wdExportFormatPDF = 17. May prompt/hang on some hosts; the app's stall timer guards it.
                    $doc.ExportAsFixedFormat($dest, 17)
                } else {
                    try { $doc.SaveAs2($dest, 16) } catch { $doc.SaveAs($dest, 16) }
                }
                Write-ProgressJson "word-saved" $script:progressCurrent $totalItems $src $script:officePid
                $doc.Close($false)
                Write-ProgressJson "word-closed" $script:progressCurrent $totalItems $src $script:officePid
                [void]$results.Add([pscustomobject]@{ src = $src; dest = $dest; ok = $true; error = $null; engine = $script:engineName })
                Set-DestTimes $dest $src
                Handle-Source $src $backupDir
            } catch {
                [void]$results.Add([pscustomobject]@{ src = $src; dest = $dest; ok = $false; error = $_.Exception.Message })
            }
        }
    } finally { if ($word) { try { $word.Quit() } catch { } } }
}

function Convert-Excel($list) {
    if ($list.Count -eq 0) { return }
    $excel = $null
    try {
        $beforeIds = Get-Pids @("et", "EXCEL")
        $excel = New-OfficeApp @("KET.Application") @("Excel.Application")
        if (-not $excel) { throw "No Excel automation engine available (WPS or MS Office)." }
        try { $excel.Visible = $false } catch { }
        try { $excel.DisplayAlerts = $false } catch { }
        $script:officePid = Get-NewPid @("et", "EXCEL") $beforeIds
        foreach ($it in $list) {
            $src = $it.src
            $script:progressCurrent++
            Write-ProgressJson "excel" $script:progressCurrent $totalItems $src $script:officePid
            $dest = Get-DestPath $src $(if ($script:comPdf) { ".pdf" } else { ".xlsx" })
            try {
                if (Test-Path -LiteralPath $dest) { throw "Destination already exists: $dest" }
                $wb = $excel.Workbooks.Open($src, 0, $true)
                if ($script:comPdf) { $wb.ExportAsFixedFormat(0, $dest) } else { $wb.SaveAs($dest, 51) }
                $wb.Close($false)
                [void]$results.Add([pscustomobject]@{ src = $src; dest = $dest; ok = $true; error = $null; engine = $script:engineName })
                Set-DestTimes $dest $src
                Handle-Source $src $backupDir
            } catch {
                [void]$results.Add([pscustomobject]@{ src = $src; dest = $dest; ok = $false; error = $_.Exception.Message })
            }
        }
    } finally { if ($excel) { try { $excel.Quit() } catch { } } }
}

function Convert-Ppt($list) {
    if ($list.Count -eq 0) { return }
    $ppt = $null
    try {
        $beforeIds = Get-Pids @("wpp", "POWERPNT")
        $ppt = New-OfficeApp @("KWPP.Application") @("PowerPoint.Application")
        if (-not $ppt) { throw "No PowerPoint automation engine available (WPS or MS Office)." }
        $script:officePid = Get-NewPid @("wpp", "POWERPNT") $beforeIds
        foreach ($it in $list) {
            $src = $it.src
            $script:progressCurrent++
            Write-ProgressJson "powerpoint" $script:progressCurrent $totalItems $src $script:officePid
            $dest = Get-DestPath $src $(if ($script:comPdf) { ".pdf" } else { ".pptx" })
            try {
                if (Test-Path -LiteralPath $dest) { throw "Destination already exists: $dest" }
                $pres = $ppt.Presentations.Open($src, $true, $false, $false)
                if ($script:comPdf) { $pres.SaveAs($dest, 32) } else { $pres.SaveAs($dest, 24) }
                $pres.Close()
                [void]$results.Add([pscustomobject]@{ src = $src; dest = $dest; ok = $true; error = $null; engine = $script:engineName })
                Set-DestTimes $dest $src
                Handle-Source $src $backupDir
            } catch {
                [void]$results.Add([pscustomobject]@{ src = $src; dest = $dest; ok = $false; error = $_.Exception.Message })
            }
        }
    } finally { if ($ppt) { try { $ppt.Quit() } catch { } } }
}

# ---------------------------------------------------------------- engine selection
# LibreOffice headless is the preferred, desktop-independent path when present.
# The app child process has a lean environment (SESSIONNAME empty, ProgramFiles may be unset),
# so do not rely on $env:ProgramFiles alone.
$sofficePath = ""
$sofficeCandidates = @()
foreach ($root in @($env:ProgramFiles, $env:ProgramW6432, [Environment]::GetFolderPath('ProgramFiles'), "C:\Program Files", "C:\Program Files (x86)", $env:LOCALAPPDATA)) {
    if ($root) { $sofficeCandidates += (Join-Path $root "LibreOffice\program\soffice.exe") }
}
foreach ($cand in $sofficeCandidates) {
    if ($cand -and (Test-Path -LiteralPath $cand)) { $sofficePath = $cand; break }
}
if (-not $sofficePath) {
    $cmd = Get-Command soffice.exe -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source) { $sofficePath = $cmd.Source }
}
$useLibreOffice = ($sofficePath -ne "") -and ($enginePref -eq "auto" -or $enginePref -eq "libreoffice")

function Convert-LibreOffice($list) {
    if ($list.Count -eq 0) { return }
    $script:engineName = "libreoffice"
    $loProfile = Join-Path $env:TEMP ("lo_profile_" + [Guid]::NewGuid().ToString())
    $loProfileUrl = "file:///" + ($loProfile -replace '\\', '/')
    foreach ($it in $list) {
        $src = $it.src
        $targetExt = "." + $it.to
        $phase = if ($it.type -eq "xls") { "excel" } elseif ($it.type -eq "ppt") { "powerpoint" } else { "word" }
        $script:progressCurrent++
        Write-ProgressJson $phase $script:progressCurrent $totalItems $src $script:officePid
        $dest = Get-DestPath $src $targetExt
        try {
            if (Test-Path -LiteralPath $dest) { throw "Destination already exists: $dest" }
            $targetDir = Split-Path -Parent $dest
            if (-not (Test-Path -LiteralPath $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
            $convertTo = if ($it.filter) { $it.to + ":" + $it.filter } else { $it.to }
            $beforeIds = Get-Pids @("soffice.bin", "soffice")
            $loArgs = @(
                "--headless", "--norestore", "--nolockcheck", "--nodefault", "--nofirststartwizard",
                ("-env:UserInstallation=" + $loProfileUrl),
                "--convert-to", $convertTo,
                "--outdir", $targetDir, $src
            )
            $quoted = foreach ($a in $loArgs) { if ($a -match '[ \t"]') { '"' + ($a -replace '"', '\"') + '"' } else { $a } }
            $proc = Start-Process -FilePath $sofficePath -ArgumentList ($quoted -join ' ') -PassThru -WindowStyle Hidden
            $script:officePid = Get-NewPid @("soffice.bin", "soffice") $beforeIds
            try { $proc.WaitForExit(600000) | Out-Null } catch { }
            # soffice.exe may return while soffice.bin keeps working; wait briefly for the output.
            $waited = 0
            while (-not (Test-Path -LiteralPath $dest) -and $waited -lt 120) { Start-Sleep -Seconds 1; $waited++ }
            if (-not (Test-Path -LiteralPath $dest)) { throw "LibreOffice did not produce the target file (exit $($proc.ExitCode))." }
            [void]$results.Add([pscustomobject]@{ src = $src; dest = $dest; ok = $true; error = $null; engine = "libreoffice"; target = $it.to })
            Set-DestTimes $dest $src
            Handle-Source $src $backupDir
        } catch {
            [void]$results.Add([pscustomobject]@{ src = $src; dest = $dest; ok = $false; error = $_.Exception.Message })
        }
    }
    try { if (Test-Path -LiteralPath $loProfile) { Remove-Item -LiteralPath $loProfile -Recurse -Force -ErrorAction SilentlyContinue } } catch { }
}

if ($useLibreOffice) {
    Write-ProgressJson "engine" 0 $totalItems ("using libreoffice: $sofficePath") 0
    Convert-LibreOffice $items
} else {
    if ($toFormat) {
        if ($toFormat -ne "pdf") {
            Write-Result ([pscustomobject]@{ engine = "none"; error = ("Target format '" + $toFormat + "' requires LibreOffice, which was not found. With only Office/WPS installed, the COM fallback can produce .docx/.xlsx/.pptx and .pdf. Install LibreOffice for the full format matrix: https://www.libreoffice.org/download/"); results = @() })
            exit 0
        }
        $script:comPdf = $true
    } else {
        $script:comPdf = $false
    }
    Write-ProgressJson "engine" 0 $totalItems "using office com (wps/ms)" 0
    Convert-Word  @($items | Where-Object { $_.type -eq "doc" })
    Convert-Ppt   @($items | Where-Object { $_.type -eq "ppt" })
    Convert-Excel @($items | Where-Object { $_.type -eq "xls" })
}
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

$reportedOutDir = if ($outDir) { $outDir } else { $srcRoot }
Write-Result ([pscustomobject]@{
    engine = $script:engineName
    input = $inputPath
    output = $reportedOutDir
    backupDir = $backupDir
    total = $results.Count
    results = $results
})
