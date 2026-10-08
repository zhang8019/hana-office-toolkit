# env_probe.ps1 -- report which document-tooling dependencies exist on this machine.
# Called by the Office Toolkit app (tool: office_deps_status, route: /deps).
# Prints one JSON object to stdout.
#
# Learned the hard way:
#   1. ASCII-only (PS 5.1 decodes BOM-less UTF-8 as ANSI).
#   2. NEVER launch an external program here. `soffice --version` opens a console and blocks on
#      "Press Enter to continue". Read file version metadata instead.
#   3. `$null -ne ""` is TRUE in PowerShell, so every "found?" check must use
#      [string]::IsNullOrWhiteSpace, not `-ne ""`. Otherwise a missing dependency reports as
#      found with a null path.
#   4. The app child process has a lean environment: $env:ProgramFiles / $env:LOCALAPPDATA may be
#      empty, so every Join-Path against them must be guarded.

$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Join-Under($root, $rel) {
    if ([string]::IsNullOrWhiteSpace($root)) { return "" }
    try { return (Join-Path $root $rel) } catch { return "" }
}

function Find-First($paths) {
    foreach ($p in @($paths)) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        try { if (Test-Path -LiteralPath $p) { return [string]$p } } catch { }
    }
    return ""
}

function Cmd-Path($name) {
    $c = Get-Command $name -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($c -and (-not [string]::IsNullOrWhiteSpace($c.Source))) { return [string]$c.Source }
    return ""
}

# Version straight from file metadata: no process launch.
function File-Version($exe) {
    if ([string]::IsNullOrWhiteSpace($exe)) { return "" }
    try {
        $vi = (Get-Item -LiteralPath $exe -ErrorAction Stop).VersionInfo
        if ($vi.ProductVersion) { return ([string]$vi.ProductVersion).Trim() }
        if ($vi.FileVersion) { return ([string]$vi.FileVersion).Trim() }
    } catch { }
    return ""
}

# Program-files roots, from several sources because the child env may be minimal.
$pfRoots = @(
    $env:ProgramFiles,
    $env:ProgramW6432,
    [Environment]::GetEnvironmentVariable("ProgramFiles", "Machine"),
    [Environment]::GetFolderPath("ProgramFiles"),
    (Join-Under $env:SystemDrive "Program Files"),
    "C:\Program Files"
) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

$localRoots = @(
    $env:LOCALAPPDATA,
    [Environment]::GetEnvironmentVariable("LOCALAPPDATA", "User"),
    [Environment]::GetFolderPath("LocalApplicationData"),
    (Join-Under $env:USERPROFILE "AppData\Local")
) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

# --- Microsoft Office (PRIMARY conversion engine when present; file metadata only, no launch) ---
$wordCandidates = @()
foreach ($r in $pfRoots) {
    $wordCandidates += (Join-Under $r "Microsoft Office\Root\Office16\WINWORD.EXE")
    $wordCandidates += (Join-Under $r "Microsoft Office\Office16\WINWORD.EXE")
}
$word = Find-First $wordCandidates
$msRegVer = ""
try { $msRegVer = [string](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue).VersionToReport } catch { }

# --- WPS Office (SECOND conversion engine) ---
$wps = ""
$wpsDirVer = ""
foreach ($r in @($localRoots) + @($pfRoots)) {
    $base = Join-Under $r "Kingsoft\WPS Office"
    if (-not $base) { continue }
    if (Test-Path -LiteralPath $base) {
        $subs = Get-ChildItem -LiteralPath $base -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending
        foreach ($s in $subs) {
            $cand = Join-Path $s.FullName "office6\wps.exe"
            if (Test-Path -LiteralPath $cand) { $wps = $cand; $wpsDirVer = $s.Name; break }
        }
    }
    if (-not $wps) {
        $cand2 = Join-Under $base "office6\wps.exe"
        if ($cand2 -and (Test-Path -LiteralPath $cand2)) { $wps = $cand2 }
    }
    if ($wps) { break }
}

# --- LibreOffice (FALLBACK engine: only it can produce html/csv/txt/rtf/odt/ods/odp/epub/md) ---
$sofficeCandidates = @($env:SOFFICE_PATH)
foreach ($r in $pfRoots) { $sofficeCandidates += (Join-Under $r "LibreOffice\program\soffice.exe") }
foreach ($r in $localRoots) { $sofficeCandidates += (Join-Under $r "LibreOffice\program\soffice.exe") }
$sofficeCandidates += (Cmd-Path "soffice.exe")
$soffice = Find-First $sofficeCandidates

# --- uv / uvx (needed to run the bundled document MCP) ---
$uv = Find-First @((Cmd-Path "uv.exe"), (Cmd-Path "uv"))
$uvx = Find-First @((Cmd-Path "uvx.exe"), (Cmd-Path "uvx"))

# --- officecli (bridge for docx/xlsx/pptx read & write) ---
# OFFICECLI_PATH wins; otherwise the installer's default location under LOCALAPPDATA.
$officecli = Find-First @($env:OFFICECLI_PATH)
if (-not $officecli) {
    $cliCandidates = @()
    foreach ($r in $localRoots) { $cliCandidates += (Join-Under $r "OfficeCli\officecli.exe") }
    foreach ($r in $pfRoots) { $cliCandidates += (Join-Under $r "OfficeCli\officecli.exe") }
    $officecli = Find-First ($cliCandidates + @((Cmd-Path "officecli.exe"), (Cmd-Path "officecli")))
}

# --- COM registration visibility (registry lookup only, no activation) ---
function Has-ProgId($progId) {
    try { return ($null -ne [Type]::GetTypeFromProgID($progId, $false)) } catch { return $false }
}
$hasWps = (Has-ProgId "KWPS.Application") -or (Has-ProgId "KET.Application")
$hasMsOffice = (Has-ProgId "Word.Application") -or (Has-ProgId "Excel.Application")

$msVersion = File-Version $word
if (-not $msVersion) { $msVersion = $msRegVer }
$wpsVersion = File-Version $wps
if (-not $wpsVersion) { $wpsVersion = $wpsDirVer }

$result = [pscustomobject]@{
    officeMs    = [pscustomobject]@{ found = (-not [string]::IsNullOrWhiteSpace($word));   path = [string]$word;   version = $msVersion;  progId = $hasMsOffice }
    officeWps   = [pscustomobject]@{ found = (-not [string]::IsNullOrWhiteSpace($wps));    path = [string]$wps;    version = $wpsVersion; progId = $hasWps }
    libreoffice = [pscustomobject]@{ found = (-not [string]::IsNullOrWhiteSpace($soffice)); path = [string]$soffice; version = (File-Version $soffice) }
    officecli   = [pscustomobject]@{ found = (-not [string]::IsNullOrWhiteSpace($officecli)); path = [string]$officecli; version = (File-Version $officecli) }
    uv          = [pscustomobject]@{ found = (-not [string]::IsNullOrWhiteSpace($uv));        path = [string]$uv }
    uvx         = [pscustomobject]@{ found = (-not [string]::IsNullOrWhiteSpace($uvx));       path = [string]$uvx }
    officeCom   = [pscustomobject]@{ wps = $hasWps; ms = $hasMsOffice }
    roots       = [pscustomobject]@{ programFiles = ($pfRoots -join " | "); localAppData = ($localRoots -join " | ") }
}
$result | ConvertTo-Json -Depth 5 -Compress
