# install_officecli.ps1 -- download the official officecli Windows binary and install it locally.
# Run by the Office Toolkit app (tool: office_deps_install, target=officecli).
# Prints one JSON object to stdout.
#
# Rules: ASCII-only (may run under PowerShell 5.1); never launch a GUI or anything that waits
# for input; verify the published sha256 before replacing the installed binary.

param(
    [string]$OutDir = (Join-Path $env:LOCALAPPDATA "OfficeCli"),
    [string]$Repo = "iOfficeAI/OfficeCLI"
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ProgressPreference = "SilentlyContinue"

function Finish($obj) { $obj | ConvertTo-Json -Compress; exit 0 }

$arch = if ($env:PROCESSOR_ARCHITECTURE -eq "ARM64") { "arm64" } else { "x64" }
$assetName = "officecli-win-$arch.exe"

# Direct first, then common GitHub mirrors (same idea as the github-cli app).
$mirrors = @("", "https://gh-proxy.com/", "https://ghproxy.net/", "https://ghfast.top/")

try {
    $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/latest" `
        -Headers @{ "User-Agent" = "office-toolkit" } -TimeoutSec 60
} catch {
    Finish ([pscustomobject]@{ ok = $false; error = ("could not query latest release: " + $_.Exception.Message) })
}

$asset = $rel.assets | Where-Object { $_.name -eq $assetName } | Select-Object -First 1
if (-not $asset) {
    Finish ([pscustomobject]@{ ok = $false; error = ("release " + $rel.tag_name + " has no asset named " + $assetName) })
}
$url = $asset.browser_download_url
$expected = [string]$asset.digest -replace '^sha256:', ''

New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$tmp = Join-Path $env:TEMP ("officecli_" + [Guid]::NewGuid().ToString() + ".exe")

$ok = $false
$lastErr = ""
foreach ($m in $mirrors) {
    try {
        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add("User-Agent", "office-toolkit")
        $wc.DownloadFile(($m + $url), $tmp)
        $wc.Dispose()
        $ok = $true
        break
    } catch {
        $lastErr = $_.Exception.Message
        try { if ($wc) { $wc.Dispose() } } catch { }
    }
}
if (-not $ok) {
    Finish ([pscustomobject]@{ ok = $false; error = ("download failed from all sources: " + $lastErr) })
}

$actual = (Get-FileHash -LiteralPath $tmp -Algorithm SHA256).Hash.ToLower()
if ($expected -and ($actual -ne $expected)) {
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    Finish ([pscustomobject]@{ ok = $false; error = ("checksum mismatch: expected " + $expected + ", got " + $actual) })
}

$dest = Join-Path $OutDir "officecli.exe"
Move-Item -LiteralPath $tmp -Destination $dest -Force

Finish ([pscustomobject]@{ ok = $true; version = $rel.tag_name; path = $dest; sha256 = $actual })
