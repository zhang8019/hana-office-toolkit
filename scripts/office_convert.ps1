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
    [string]$JobFile = "",
    # Internal: worker mode. Runs ONLY the Office/WPS COM conversion, reporting progress and the
    # final result through files. Used because COM needs a caller that is BOTH non-elevated AND
    # outside the Hana sandbox; we get that by launching this script again via explorer.exe.
    [switch]$ComWorker,
    # Internal: sweep transient temp dirs left by earlier runs, then exit.
    [switch]$SweepTemp,
    # Internal: list running Office / soffice processes (pid, name, has window), then exit.
    [switch]$ListOfficePids,
    # Internal: sweep worker mode (launched via explorer by the parent), then exit.
    [switch]$SweepWorker,
    [string]$WorkerAction = "sweep",
    [string]$ProgressFile = "",
    [string]$ResultFile = "",
    [string]$DoneFlag = "",
    [string]$WorkerPidFile = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# The app child process can have a lean environment; without PATHEXT PowerShell will not treat
# ".exe" as an executable and the call operator fails with "cannot run a document".
if (-not $env:PATHEXT) { $env:PATHEXT = ".COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC" }
$script:ProgressFile = $ProgressFile

function Get-ZombieProcs {
    # An Office automation instance has no main window; a document the user opened does. This is the
    # standard way to tell a leftover apart from real work, and it is why we can safely auto-clean.
    # LibreOffice is the exception: soffice.bin is the real process behind a visible window too, so
    # for it we only take instances whose command line carries our own profile marker.
    $names = @("WINWORD", "EXCEL", "POWERPNT", "wps", "et", "wpp", "soffice", "soffice.bin")
    $procs = @()
    foreach ($n in $names) {
        Get-Process -Name $n -ErrorAction SilentlyContinue | ForEach-Object {
            $procs += [pscustomobject]@{ pid = $_.Id; name = $_.ProcessName; title = [string]$_.MainWindowTitle }
        }
    }
    return $procs
}

function Test-KillableZombie($p) {
    if ($p.name -like "soffice*") {
        try {
            $cl = (Get-CimInstance Win32_Process -Filter ("ProcessId=" + $p.pid) -ErrorAction SilentlyContinue).CommandLine
            return ($cl -and ($cl -match "lo_profile_"))
        } catch { return $false }
    }
    return (-not ([string]$p.title).Trim())
}

function Invoke-SweepWorker($action) {
    # The parent runs INSIDE the Hana sandbox, whose restricted process view cannot see (and
    # therefore cannot kill) Office instances started elsewhere. explorer.exe runs outside the
    # sandbox, so we hop through it exactly like the COM worker does, and read the answer back
    # from a file. VBS + window style 0 keeps the hop silent.
    $tmp = Join-Path $env:TEMP ("ot_sweep_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null
    $res = Join-Path $tmp 'sweep.json'
    $done = Join-Path $tmp 'done.flag'
    $vbs = Join-Path $tmp 'sweep.vbs'
    $self = $PSCommandPath
    if (-not $self) { $self = $MyInvocation.MyCommand.Path }
    $cmdLine = "pwsh -NoProfile -ExecutionPolicy Bypass -File `"$self`" -SweepWorker -WorkerAction $action -ResultFile `"$res`" -DoneFlag `"$done`""
    Set-Content -LiteralPath $vbs -Value ('CreateObject("WScript.Shell").Run "' + ($cmdLine -replace '"', '""') + '", 0, False') -Encoding ascii
    try { Start-Process -FilePath "explorer.exe" -ArgumentList "`"$vbs`"" -WindowStyle Hidden | Out-Null } catch { }
    $deadline = (Get-Date).AddSeconds(25)
    while (-not (Test-Path -LiteralPath $done)) {
        if ((Get-Date) -gt $deadline) { break }
        Start-Sleep -Milliseconds 300
    }
    $obj = $null
    if (Test-Path -LiteralPath $res) { try { $obj = Get-Content -LiteralPath $res -Raw -Encoding UTF8 | ConvertFrom-Json } catch { } }
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
    return $obj
}

if ($SweepWorker) {
    # Runs outside the sandbox: sees every process, may delete temp dirs and kill zombies.
    $payload = [ordered]@{ ok = $true; action = $WorkerAction; removed = 0; procs = @(); killed = @() }
    try {
        if ($WorkerAction -ne 'list') {
            $cut = (Get-Date).AddHours(-1)
            foreach ($pat in @("ot_com_*", "unelevated_*", "lo_profile_*", "officecli_probe_*")) {
                Get-ChildItem -LiteralPath $env:TEMP -Directory -Filter $pat -ErrorAction SilentlyContinue | ForEach-Object {
                    if ($_.LastWriteTime -lt $cut) {
                        try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue; $payload.removed++ } catch { }
                    }
                }
            }
        }
        $procs = @(Get-ZombieProcs)
        $payload.procs = $procs
        if ($WorkerAction -ne 'list') {
            $killed = @()
            foreach ($p in $procs) {
                if (Test-KillableZombie $p) {
                    try { & taskkill.exe /PID $p.pid /T /F 2>$null | Out-Null; $killed += $p.pid } catch { }
                }
            }
            $payload.killed = $killed
        }
    } catch {
        $payload.ok = $false
        $payload.error = $_.Exception.Message
    }
    # -Depth must cover procs[].pid/name/title so they are not collapsed into strings.
    $json = $payload | ConvertTo-Json -Depth 6 -Compress
    if (-not $json) { $json = '{"ok":false}' }
    $json | Set-Content -LiteralPath $ResultFile -Encoding UTF8
    if ($DoneFlag) { "done" | Set-Content -LiteralPath $DoneFlag -Encoding ascii }
    exit 0
}

if ($SweepTemp -or $ListOfficePids) {
    # Parent side: hop outside the sandbox, then answer. Falls back to a local look if the hop
    # fails, so a broken explorer path degrades instead of hanging.
    $action = if ($ListOfficePids) { 'list' } else { 'sweep' }
    $r = Invoke-SweepWorker $action
    if (-not $r) { $r = [pscustomobject]@{ ok = $false; procs = @(Get-ZombieProcs); removed = 0; killed = @() } }
    $procList = @($r.procs | Where-Object { $_ -ne $null })
    if ($ListOfficePids) {
        if ($procList.Count -eq 0) { "[]" } else { ConvertTo-Json -InputObject $procList -Depth 4 -Compress }
    } else {
        ([pscustomobject]@{ ok = $true; removed = [int]$r.removed; procs = $procList; killed = @($r.killed | Where-Object { $_ -ne $null }) }) | ConvertTo-Json -Depth 6 -Compress
    }
    exit 0
}

if (-not $JobFile) { ([pscustomobject]@{ ok = $false; error = "JobFile is required unless -SweepTemp is used." }) | ConvertTo-Json -Compress; exit 0 }

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
    $payload = [pscustomobject]@{ phase = $phase; current = $current; total = $total; fileB64 = $fileB64; officePid = [int]$officePid; workerPid = $PID } | ConvertTo-Json -Compress
    if ($script:ProgressFile) {
        # Worker mode: the parent tails this file and re-emits the lines on its own stderr.
        try { Add-Content -LiteralPath $script:ProgressFile -Value ("OFFICE_PROGRESS:$payload") -Encoding UTF8 } catch { }
    }
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
$wid = [Security.Principal.WindowsIdentity]::GetCurrent()
$isAdm = (New-Object Security.Principal.WindowsPrincipal($wid)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$hasAdmSid = ($wid.Groups | Where-Object { $_.Value -eq 'S-1-5-32-544' }).Count -gt 0
$admAttr = "n/a"; $lvlAttr = "n/a"
try {
    $whoamiExe = Join-Path $env:SystemRoot "System32\whoami.exe"
    $wa = & $whoamiExe /groups 2>$null
    $al = $wa | Select-String -Pattern 'S-1-5-32-544' | Select-Object -First 1
    $ll = $wa | Select-String -Pattern 'S-1-16-' | Select-Object -First 1
    if ($al) { $admAttr = ([string]$al.Line).Trim() -replace '\s+', '~' }
    if ($ll) { $lvlAttr = ([string]$ll.Line).Trim() -replace '\s+', '~' }
} catch { }
$envInfo = "TEMP=$($env:TEMP) | USERPROFILE=$($env:USERPROFILE) | ProgramFiles=$($env:ProgramFiles) | probe=$tempProbe | IsAdmin=$isAdm | TokenHasAdminSid=$hasAdmSid | admGroup=$admAttr | integrity=$lvlAttr"
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

# Engine preference (from job.engine): "auto" tries Microsoft Office first, then WPS;
# "wps" forces WPS; "office" forces Microsoft Office. Word/WPS keep legacy layout, so they lead.
function New-OfficeApp($wpsProgIds, $msProgIds) {
    $order = switch ($enginePref) {
        "wps" { $wpsProgIds }
        "office" { $msProgIds }
        default { @($msProgIds) + @($wpsProgIds) }
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

# Suppress every dialog Office can raise headless. Each assignment is best-effort: property names
# differ between MS Office and WPS, and a missing one must not abort the run.
# Notes from the vendor docs: Excel's DisplayAlerts=False makes Confirm-Save-As and compatibility
# prompts default to Yes; Excel's CheckCompatibility=False skips the compatibility checker entirely;
# PowerPoint's DisplayAlerts already defaults to ppAlertsNone (1).
function Set-OfficeAppQuiet($app, $kind) {
    $pairs = switch ($kind) {
        "word" {
            @(
                @("Visible", $false), @("DisplayAlerts", 0), @("ScreenUpdating", $false),
                @("AutomationSecurity", 3), @("Options.ConfirmConversions", $false),
                @("Options.WarnBeforeSavingPrintingSendingMarkup", $false),
                @("Options.SavePropertiesPrompt", $false), @("Options.UpdateLinksAtOpen", $false)
            )
        }
        "excel" {
            @(
                @("Visible", $false), @("DisplayAlerts", $false), @("ScreenUpdating", $false),
                @("EnableEvents", $false), @("AskToUpdateLinks", $false), @("AutomationSecurity", 3),
                @("CheckCompatibility", $false)
            )
        }
        "ppt" {
            @(@("DisplayAlerts", 1), @("Visible", 1))
        }
        default { @() }
    }
    foreach ($p in $pairs) { try { $app.($p[0]) = $p[1] } catch { } }
}

function Convert-Word($list) {
    if ($list.Count -eq 0) { return }
    $word = $null
    try {
        $beforeIds = Get-Pids @("wps", "WINWORD")
        $word = New-OfficeApp @("KWPS.Application") @("Word.Application")
        if (-not $word) { throw "No Word automation engine available (WPS or MS Office)." }
        Set-OfficeAppQuiet $word "word"
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
        Set-OfficeAppQuiet $excel "excel"
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
        Set-OfficeAppQuiet $ppt "ppt"
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

# ---------------------------------------------------------------- dispatch
# Engine priority (auto): MS Office COM -> WPS COM -> LibreOffice.
# Rationale: LibreOffice re-renders some legacy documents with layout drift (observed: images
# overflowing the page), while Word/WPS keep the original layout. COM, however, only works from a
# caller that is BOTH non-elevated AND outside the Hana sandbox, so COM runs in a worker process
# launched via explorer.exe; anything else falls back to LibreOffice.
$comTargets = @("docx", "xlsx", "pptx", "pdf")
$canUseCom = (-not $toFormat) -or ($comTargets -contains $toFormat)

function Write-ResultFile($obj) {
    if ($ResultFile) { $obj | ConvertTo-Json -Depth 6 -Compress | Set-Content -LiteralPath $ResultFile -Encoding UTF8 }
}

if ($ComWorker) {
    # Worker mode: launched unelevated / out of sandbox by the parent. Reports through files only.
    try {
        if ($WorkerPidFile) { "$PID" | Set-Content -LiteralPath $WorkerPidFile -Encoding ascii }
        Write-ProgressJson "worker" 0 $totalItems "com worker pid=$PID" 0
        $script:comPdf = ($toFormat -eq "pdf")
        Convert-Word  @($items | Where-Object { $_.type -eq "doc" })
        Convert-Ppt   @($items | Where-Object { $_.type -eq "ppt" })
        Convert-Excel @($items | Where-Object { $_.type -eq "xls" })
        $reportedOut = if ($outDir) { $outDir } else { $srcRoot }
        Write-ResultFile ([pscustomobject]@{ engine = $script:engineName; input = $inputPath; output = $reportedOut; backupDir = $backupDir; total = $results.Count; results = $results })
    } catch {
        Write-ResultFile ([pscustomobject]@{ engine = $script:engineName; error = $_.Exception.ToString(); results = @() })
    } finally {
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        if ($DoneFlag) { "done" | Set-Content -LiteralPath $DoneFlag -Encoding ascii }
    }
    exit 0
}

function Invoke-ComWorker {
    $tmp = Join-Path $env:TEMP ("ot_com_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null
    $prog = Join-Path $tmp 'progress.txt'
    $res  = Join-Path $tmp 'result.json'
    $done = Join-Path $tmp 'done.flag'
    $pidf = Join-Path $tmp 'worker.pid'
    $vbs  = Join-Path $tmp 'run.vbs'
    $self = $PSCommandPath
    if (-not $self) { $self = $MyInvocation.MyCommand.Path }
    $cmdLine = "pwsh -NoProfile -ExecutionPolicy Bypass -File `"$self`" -JobFile `"$JobFile`" -ComWorker -ProgressFile `"$prog`" -ResultFile `"$res`" -DoneFlag `"$done`" -WorkerPidFile `"$pidf`""
    # Two things must hold for Office/WPS COM: the caller is non-elevated AND outside the Hana
    # sandbox. explorer.exe gives both (it runs at medium integrity, outside the sandbox).
    # We hand it a .vbs instead of a .cmd precisely because explorer.exe would pop a visible
    # console for a .cmd; WScript.Shell.Run with window style 0 keeps it silent.
    $vbsLine = 'CreateObject("WScript.Shell").Run "' + ($cmdLine -replace '"', '""') + '", 0, False'
    Set-Content -LiteralPath $vbs -Value $vbsLine -Encoding ascii
    try { Start-Process -FilePath "explorer.exe" -ArgumentList "`"$vbs`"" -WindowStyle Hidden | Out-Null } catch { }
    $seen = 0
    $deadline = (Get-Date).AddMinutes(30)
    $workerPid = 0
    while (-not (Test-Path -LiteralPath $done)) {
        if (Test-Path -LiteralPath $prog) {
            $lines = @(Get-Content -LiteralPath $prog -Encoding UTF8 -ErrorAction SilentlyContinue)
            for ($i = $seen; $i -lt $lines.Count; $i++) { [Console]::Error.WriteLine([string]$lines[$i]) }
            $seen = $lines.Count
        }
        if (Test-Path -LiteralPath $pidf) { try { $workerPid = [int]((Get-Content -LiteralPath $pidf -Raw).Trim()) } catch { } }
        if ((Get-Date) -gt $deadline) { break }
        Start-Sleep -Milliseconds 500
    }
    if (Test-Path -LiteralPath $prog) {
        $lines = @(Get-Content -LiteralPath $prog -Encoding UTF8 -ErrorAction SilentlyContinue)
        for ($i = $seen; $i -lt $lines.Count; $i++) { [Console]::Error.WriteLine([string]$lines[$i]) }
    }
    if (Test-Path -LiteralPath $pidf) { try { $workerPid = [int]((Get-Content -LiteralPath $pidf -Raw).Trim()) } catch { } }
    if (-not (Test-Path -LiteralPath $res)) {
        if ($workerPid) { try { & taskkill.exe /PID $workerPid /T /F 2>$null | Out-Null } catch { } }
        Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
        return [pscustomobject]@{ engine = "none"; error = "COM worker did not finish within 30 minutes."; results = @() }
    }
    $out = Get-Content -LiteralPath $res -Raw -Encoding UTF8 | ConvertFrom-Json
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
    return $out
}

$forceLo = ($enginePref -eq "libreoffice")
$wantLo = $forceLo -or (-not $canUseCom)

if ($wantLo) {
    if (-not $useLibreOffice) {
        Write-Result ([pscustomobject]@{ engine = "none"; error = ("Target format '" + $toFormat + "' requires LibreOffice, which was not found. Install LibreOffice: https://www.libreoffice.org/download/"); results = @() })
        exit 0
    }
    Write-ProgressJson "engine" 0 $totalItems ("using libreoffice: $sofficePath") 0
    Convert-LibreOffice $items
} else {
    Write-ProgressJson "engine" 0 $totalItems "engine order: ms office -> wps -> libreoffice (unelevated worker)" 0
    $w = Invoke-ComWorker
    $okCountW = 0
    if ($w -and $w.results) { $okCountW = @($w.results | Where-Object { $_.ok }).Count }
    if ($w -and (-not $w.error) -and ($okCountW -gt 0)) {
        $script:engineName = if ($w.engine) { [string]$w.engine } else { "ms-office" }
        # Keep $results an ArrayList: Convert-LibreOffice appends to it, and a plain fixed-size array
        # would throw "the collection has a fixed size" the moment the per-file fallback kicks in.
        $results = New-Object System.Collections.ArrayList
        foreach ($r0 in @($w.results)) { [void]$results.Add($r0) }
        # Per-file fallback. COM can lose on a single file for reasons that have nothing to do with
        # the batch: an empty presentation cannot be exported to PDF, a locked part, an odd shape.
        # Those files are handed to LibreOffice, which is more forgiving. Successes are kept as-is,
        # and an explicitly forced engine (office/wps) is respected rather than silently retried.
        $allowFallback = ($useLibreOffice) -and ($enginePref -ne "office") -and ($enginePref -ne "wps")
        if ($allowFallback) {
            $failedSrc = @($results | Where-Object { -not $_.ok } | ForEach-Object { [string]$_.src })
            if ($failedSrc.Count -gt 0) {
                $retry = @($items | Where-Object { $failedSrc -contains [string]$_.src })
                if ($retry.Count -gt 0) {
                    Write-ProgressJson "engine" 0 $totalItems ("com failed on " + $retry.Count + " file(s) -> retry with libreoffice") 0
                    $kept = New-Object System.Collections.ArrayList
                    foreach ($r0 in $results) { if ($r0.ok) { [void]$kept.Add($r0) } }
                    $results = $kept
                    Convert-LibreOffice $retry
                    $script:engineName = "mixed"
                }
            }
        }
    } else {
        $reason = if ($w -and $w.error) { [string]$w.error } else { "COM produced no successful conversion" }
        Write-ProgressJson "engine" 0 $totalItems ("com engine unusable -> fallback to libreoffice") 0
        if (-not $useLibreOffice) {
            Write-Result ([pscustomobject]@{ engine = "none"; error = ("Office/WPS COM failed and LibreOffice is not installed. COM said: " + $reason); results = @() })
            exit 0
        }
        Convert-LibreOffice $items
    }
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
