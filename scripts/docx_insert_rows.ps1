# docx_insert_rows.ps1 -- insert photo rows into a .docx by cloning the document's OWN photo block.
#
# Why this exists: the high-level APIs (officecli picture / the timeverse MCP word_add_image) cannot see
# or write legacy VML floating pictures, and copying geometry from a DIFFERENT document always drifts
# (each of these files has its own margins, box sizes and per-row offsets). So we clone the target
# document's own "picture paragraph / blank lines / caption paragraph" as the template and only swap the
# image bytes and the caption text.
#
# Design notes carried over from a working field implementation:
#   - Each row is wrapped in an invisible single-cell table with <w:cantSplit/> so a row can never be
#     split across a column/page boundary. w:keepNext/w:keepLines would prevent the split too, but Word
#     paints a visible black square in the left margin for those, which a recipient rejected as wrong.
#   - Image margin-top is normalised to 0: a NEGATIVE top margin is relative to the anchor paragraph, so
#     once the anchor lands at the top of a column the picture is thrown into the header.
#   - w:shapetype must appear only once per document, so it is stripped from cloned runs.
#   - Old photos are re-encoded as JPEG at a fixed width to keep the package from ballooning.
#
# Pure ASCII on purpose: PowerShell 5.1 decodes a BOM-less UTF-8 script as ANSI.
param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [Parameter(Mandatory = $true)][string]$SpecJson,
    [string]$AnchorText = "",
    [string]$Out = "",
    [int]$TargetW = 1040,
    [int]$Quality = 82,
    [string]$BorderColor = "none",
    [double]$PhotoGapPt = 0,
    [int]$RowGapExtra = 0,
    [int]$CaptionGapExtra = 0
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing

if (-not (Test-Path -LiteralPath $Docx)) { throw "docx not found: $Docx" }
if (-not (Test-Path -LiteralPath $SpecJson)) { throw "spec json not found: $SpecJson" }
if (-not $Out) {
    $dir = [IO.Path]::GetDirectoryName($Docx)
    $base = [IO.Path]::GetFileNameWithoutExtension($Docx)
    $Out = Join-Path $dir ($base + "_withphotos.docx")
}
$rows = Get-Content -LiteralPath $SpecJson -Raw -Encoding utf8 | ConvertFrom-Json
if ($rows -isnot [System.Array]) { $rows = @($rows) }

# ---------- read the package ----------
$zip = [System.IO.Compression.ZipFile]::OpenRead($Docx)
$entries = @{}
$order = New-Object System.Collections.ArrayList
foreach ($e in $zip.Entries) {
    if ([string]::IsNullOrEmpty($e.Name)) { continue }
    $ms = New-Object System.IO.MemoryStream; $s = $e.Open(); $s.CopyTo($ms); $s.Close()
    $entries[$e.FullName] = $ms.ToArray(); [void]$order.Add($e.FullName)
}
$zip.Dispose()

$docXml = [System.Text.Encoding]::UTF8.GetString($entries['word/document.xml'])
$relsXml = [System.Text.Encoding]::UTF8.GetString($entries['word/_rels/document.xml.rels'])
$ctXml = [System.Text.Encoding]::UTF8.GetString($entries['[Content_Types].xml'])
if (-not $docXml.StartsWith('<?xml')) { throw "document.xml does not start with the XML declaration" }

# ---------- page geometry helper (pt) ----------
function Get-PageGeo([string]$xml) {
    $sec = [regex]::Match($xml, '(?s)<w:sectPr.*?</w:sectPr>').Value
    $pgW = 0; $mL = 0; $mR = 0; $cols = 1; $csp = 0
    $m = [regex]::Match($sec, '<w:pgSz[^>]*w:w="(\d+)"'); if ($m.Success) { $pgW = [int]$m.Groups[1].Value }
    $m = [regex]::Match($sec, 'w:left="(\d+)"');  if ($m.Success) { $mL = [int]$m.Groups[1].Value }
    $m = [regex]::Match($sec, 'w:right="(\d+)"'); if ($m.Success) { $mR = [int]$m.Groups[1].Value }
    $m = [regex]::Match($sec, '<w:cols[^>]*w:num="(\d+)"');   if ($m.Success) { $cols = [int]$m.Groups[1].Value }
    $m = [regex]::Match($sec, '<w:cols[^>]*w:space="(\d+)"'); if ($m.Success) { $csp = [int]$m.Groups[1].Value }
    $colW = 0.0
    if ($pgW -gt 0 -and $cols -gt 0) { $colW = ($pgW - $mL - $mR - $csp * ($cols - 1)) / $cols / 20.0 }
    return @{ colW = $colW; cols = $cols }
}

# ---------- take this document's own picture run as the template ----------
# Do NOT split paragraphs with a regex to find it: a VML picture can carry a textbox whose
# <w:txbxContent> holds a NESTED <w:p>, and a non-greedy <w:p ...</w:p> match stops at that inner
# paragraph, so the picture runs are lost and the document looks picture-less. Locate runs directly.
$tplImgPPr = '<w:pPr><w:ind w:firstLineChars="0"/></w:pPr>'
$tplPictRuns = @()
$hasOwn = $false
$imgIdx = $docXml.IndexOf('<v:imagedata')
if ($imgIdx -ge 0) {
    $runs = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($docXml, '(?s)<w:r[ >][^>]*>.*?</w:r>')) {
        if ($m.Value -match '<v:imagedata') { [void]$runs.Add($m.Value) }
        if ($runs.Count -ge 2) { break }
    }
    if ($runs.Count -gt 0) {
        $tplPictRuns = @($runs)
        $hasOwn = $true
        # paragraph properties of the picture paragraph (for indent/style continuity)
        $pStart = $docXml.LastIndexOf('<w:p ', $imgIdx)
        if ($pStart -lt 0) { $pStart = $docXml.LastIndexOf('<w:p>', $imgIdx) }
        if ($pStart -ge 0) {
            $window = $docXml.Substring($pStart, [Math]::Min(3000, $docXml.Length - $pStart))
            $m = [regex]::Match($window, '(?s)<w:pPr>.*?</w:pPr>')
            if ($m.Success) { $tplImgPPr = $m.Value }
        }
    }
}
# Blank lines between picture and caption: a fixed run of empty paragraphs, matching the original layout.
$tplEmpties = ('<w:p><w:pPr><w:ind w:firstLineChars="0"/></w:pPr></w:p>' * 6)

# ---------- geometry: derive from THIS document, never from another one ----------
$geo = Get-PageGeo $docXml
$boxW = 224.0; $boxH = 168.0; $gapPt = 24.0
if ($geo.colW -gt 60) {
    $gapPt = if ($PhotoGapPt -gt 0) { $PhotoGapPt } else { 24.0 }
    $boxW = ($geo.colW - $gapPt) / 2.0 - 2.0
    if ($boxW -gt 224.0) { $boxW = 224.0 }   # field photos are 4:3 and were never larger than this
    if ($boxW -lt 60.0) { $boxW = 60.0 }
    $boxH = $boxW * 3.0 / 4.0
} elseif ($PhotoGapPt -gt 0) {
    $gapPt = $PhotoGapPt
}

if (-not $hasOwn) {
    # No pictures in this document to clone. Build the two runs from the document's OWN column width    # instead of borrowing another file's numbers (cross-document geometry always drifts).
    $tplImgPPr = '<w:pPr><w:ind w:firstLineChars="0"/></w:pPr>'
    $leftMl = [Math]::Round(($geo.colW - (2 * $boxW + $gapPt)) / 2.0, 2)
    if ($leftMl -lt 0) { $leftMl = 0 }
    $rightMl = [Math]::Round($leftMl + $boxW + $gapPt, 2)
    $wPt = [Math]::Round($boxW, 2)
    $hPt = [Math]::Round($boxH, 2)
    # Build each run in its own variable first. Do NOT mix + concatenation with , inside @():
    # @('a' + 'b', 'c' + 'd') collapses into ONE string in PowerShell, which silently produced a
    # single merged template and a corrupt XML block.
    $runLeftTpl = '<w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="_x0000_s9001" type="#_x0000_t75" style="position:absolute;left:0;text-align:left;margin-left:' + $leftMl + 'pt;margin-top:0pt;width:' + $wPt + 'pt;height:' + $hPt + 'pt;z-index:251656704;visibility:visible"><v:imagedata r:id="rId0" o:title=""/></v:shape></w:pict></w:r>'
    $runRightTpl = '<w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="_x0000_s9002" type="#_x0000_t75" style="position:absolute;left:0;text-align:left;margin-left:' + $rightMl + 'pt;margin-top:0pt;width:' + $wPt + 'pt;height:' + $hPt + 'pt;z-index:251655680;visibility:visible"><v:imagedata r:id="rId0" o:title=""/></v:shape></w:pict></w:r>'
    $tplPictRuns = @($runLeftTpl, $runRightTpl)
    $tplEmpties = ('<w:p><w:pPr><w:ind w:firstLineChars="0"/></w:pPr></w:p>' * 6)
}

# A document needs exactly one <v:shapetype>; drop it from the clones.
$tplPictRuns = $tplPictRuns | ForEach-Object { $_ -replace '(?s)<v:shapetype.*?</v:shapetype>', '' }

function Get-Geo([string]$run) {
    $s = [regex]::Match($run, 'style="([^"]*)"').Groups[1].Value
    $ml = 0.0; $w = $boxW; $h = $boxH
    $m1 = [regex]::Match($s, 'margin-left:([-0-9.]+)'); if ($m1.Success) { $ml = [double]$m1.Groups[1].Value }
    $m2 = [regex]::Match($s, 'width:([-0-9.]+)');      if ($m2.Success) { $w = [double]$m2.Groups[1].Value }
    $m3 = [regex]::Match($s, 'height:([-0-9.]+)');     if ($m3.Success) { $h = [double]$m3.Groups[1].Value }
    return @($ml, $w, $h)
}
# Map the template runs to left/right BY GEOMETRY. (The original field script hard-wired
# [last]=left / [first]=right, which was only correct because that document's first run happened to be
# the right-hand photo; deriving it from margin-left works for either order.)
$gA = Get-Geo $tplPictRuns[0]
$gB = Get-Geo $tplPictRuns[-1]
if ($gA[0] -le $gB[0]) { $gL = $gA; $gR = $gB; $runL = $tplPictRuns[0]; $runR = $tplPictRuns[-1] }
else { $gL = $gB; $gR = $gA; $runL = $tplPictRuns[-1]; $runR = $tplPictRuns[0] }

# Re-centre the pair in the column (and honour an explicit gap) when we know the column width.
if ($geo.colW -gt 60) {
    $total = $gL[1] + $gapPt + $gR[1]
    $leftMl = [Math]::Round(($geo.colW - $total) / 2.0, 2)
    if ($leftMl -lt 0) { $leftMl = 0 }
    $rightMl = [Math]::Round($leftMl + $gL[1] + $gapPt, 2)
    $runL = $runL -replace 'margin-left:[^;]*', ('margin-left:' + $leftMl + 'pt')
    $runR = $runR -replace 'margin-left:[^;]*', ('margin-left:' + $rightMl + 'pt')
    $gL[0] = $leftMl; $gR[0] = $rightMl
}
$tabL = [int](($gL[0] + $gL[1] / 2) * 20)
$tabR = [int](($gR[0] + $gR[1] / 2) * 20)
if ($tabL -lt 100) { $tabL = 3280 }
if ($tabR -le $tabL) { $tabR = 7870 }
$imgH = [Math]::Max($gL[2], $gR[2])   # same height for both photos in a row, otherwise they look off

# Character width estimate for a 12pt Song face: half-width ~6pt, full-width ~12pt.
function TxtW([string]$s) { $w = 0.0; foreach ($ch in $s.ToCharArray()) { if ([int][char]$ch -lt 128) { $w += 6.0 } else { $w += 12.0 } }; return $w }

# ---------- allocate fresh ids ----------
$maxRid = 0; foreach ($m in [regex]::Matches($relsXml, 'Id="rId(\d+)"')) { $n = [int]$m.Groups[1].Value; if ($n -gt $maxRid) { $maxRid = $n } }
$maxImg = 0; foreach ($k in $entries.Keys) { if ($k -match '^word/media/[^/]*?(\d+)\.[A-Za-z]+$') { $n = [int]$Matches[1]; if ($n -gt $maxImg) { $maxImg = $n } } }
# Shape ids must be unique document-wide. Seed from the largest _x0000_sN already present.
$maxSpid = 0; foreach ($m in [regex]::Matches($docXml, '_x0000_s(\d+)')) { $n = [int]$m.Groups[1].Value; if ($n -gt $maxSpid) { $maxSpid = $n } }

$newMedia = @{}; $relAdds = New-Object System.Text.StringBuilder
$block = New-Object System.Text.StringBuilder
$cnt = 0
$tblOpen = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="9940"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="9940" w:type="dxa"/></w:tcPr>'
$tblClose = '</w:tc></w:tr></w:tbl>'

function Normalize-Geo([string]$run, [double]$h) {
    $r = $run -replace 'margin-top:[^;]*', 'margin-top:0pt'
    $r = $r -replace 'height:[^;]*', ('height:' + $h + 'pt')
    return $r
}

function Resize-Jpeg([string]$path, [int]$tw, [int]$q) {
    $img = [System.Drawing.Image]::FromFile($path)
    try {
        $hh = [int][math]::Round($img.Height * ($tw / $img.Width))
        $bmp = New-Object System.Drawing.Bitmap($tw, $hh)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.DrawImage($img, 0, 0, $tw, $hh); $g.Dispose()
        $ms = New-Object System.IO.MemoryStream
        $enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
        $ps = New-Object System.Drawing.Imaging.EncoderParameters(1)
        $ps.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]$q)
        $bmp.Save($ms, $enc, $ps); $bmp.Dispose()
        return $ms.ToArray()
    } finally { $img.Dispose() }
}

foreach ($row in $rows) {
    $left = $row.left; $right = $row.right
    if ($null -eq $left -and $null -eq $right) { continue }
    $pairs = @()
    if ($null -ne $right) { $pairs += @{ tpl = $runR; node = $right } }
    if ($null -ne $left)  { $pairs += @{ tpl = $runL; node = $left } }
    if ($pairs.Count -eq 1) { $pairs[0].tpl = $runL }

    $newRuns = @()
    foreach ($p in $pairs) {
        $script:cnt++
        $imgPath = [string]$p.node.img
        if ([string]::IsNullOrWhiteSpace($imgPath)) { throw "row $cnt : image path is empty" }
        if (-not (Test-Path -LiteralPath $imgPath)) { throw "row $cnt : image not found: $imgPath" }
        $maxImg++; $maxRid++
        $name = "image$maxImg.jpeg"
        $newMedia["word/media/$name"] = Resize-Jpeg $imgPath $TargetW $Quality
        [void]$relAdds.Append("<Relationship Id=`"rId$maxRid`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`" Target=`"media/$name`"/>")
        $r = $p.tpl -replace 'r:id="[^"]*"', "r:id=`"rId$maxRid`""
        $r = Normalize-Geo $r $imgH
        # Fresh unique shape id per emitted run. The lookbehind keeps this from touching r:id / paraId.
        $maxSpid++
        $r = $r -replace '(?<![:\w])id="[^"]*"', ('id="pic' + $maxSpid + '"')
        $r = $r -replace 'o:spid="_x0000_s\d+"', ('o:spid="_x0000_s' + $maxSpid + '"')
        if ([string]::IsNullOrWhiteSpace($BorderColor) -or $BorderColor -eq 'none') {
            $r = $r -replace ' o:bordertopcolor="[^"]*" o:borderleftcolor="[^"]*" o:borderbottomcolor="[^"]*" o:borderrightcolor="[^"]*" stroked="t" strokecolor="[^"]*"', ''
        } else {
            # Add a VML stroke in the requested colour. The attribute order below matches what Word writes.
            $r = $r -replace ' stroked="[^"]*"', ''
            $r = $r -replace ' strokecolor="[^"]*"', ''
            $r = $r -replace ' strokeweight="[^"]*"', ''
            $r = $r -replace 'style="', ('stroked="t" strokecolor="' + $BorderColor + '" strokeweight="1pt" style="')
        }
        $newRuns += $r
    }
    [void]$block.Append($tblOpen)
    [void]$block.Append('<w:p>' + $tplImgPPr + ($newRuns -join '') + '</w:p>')
    [void]$block.Append($tplEmpties)
    for ($g = 0; $g -lt $CaptionGapExtra; $g++) { [void]$block.Append('<w:p><w:pPr><w:ind w:firstLineChars="0" w:firstLine="0"/></w:pPr></w:p>') }
    $capList = New-Object System.Collections.ArrayList
    if ($null -ne $left -and $left.cap) { [void]$capList.Add($left.cap) }
    if ($null -ne $right -and $right.cap) { [void]$capList.Add($right.cap) }
    if ($capList.Count -eq 0) { [void]$capList.Add('') }
    $cap1 = $capList[0]
    $w1 = TxtW $cap1
    if ($capList.Count -ge 2) {
        $cap2 = $capList[1]
        $w2 = TxtW $cap2
        $indentPt = [Math]::Max(0.0, ($tabL / 20.0) - $w1 / 2.0)
        $gapSpacePt = ($tabR - $tabL) / 20.0 - $w1 / 2.0 - $w2 / 2.0
        $nsp = [Math]::Max(2, [int][Math]::Floor($gapSpacePt / 6.0))
        $joined = $cap1 + (' ' * $nsp) + $cap2
    } else {
        $centerPt = if ($null -ne $left -and $null -ne $right) { ($tabL + $tabR) / 40.0 } else { $tabL / 20.0 }
        $indentPt = [Math]::Max(0.0, $centerPt - $w1 / 2.0)
        $joined = $cap1
    }
    $capPPr = '<w:pPr><w:ind w:firstLineChars="0" w:firstLine="' + [int]($indentPt * 20) + '"/></w:pPr>'
    $capRun = '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t xml:space="preserve">' + [System.Security.SecurityElement]::Escape($joined) + '</w:t></w:r>'
    [void]$block.Append('<w:p>' + $capPPr + $capRun + '</w:p>')
    for ($g = 0; $g -lt $RowGapExtra; $g++) { [void]$block.Append('<w:p><w:pPr><w:ind w:firstLineChars="0" w:firstLine="0"/></w:pPr></w:p>') }
    [void]$block.Append($tblClose)
}

# ---------- insert before the anchor paragraph ----------
if ($AnchorText) {
    # Match on each paragraph's CONCATENATED text, not on the raw XML: Word routinely splits a phrase
    # across several <w:r> runs, so a literal search for ">phrase<" misses the very paragraphs it should
    # find. Concatenating <w:t> first is the reliable way.
    $pStart = -1
    $paraMatches = [regex]::Matches($docXml, '(?s)<w:p[ >].*?</w:p>')
    foreach ($pm in $paraMatches) {
        $ptxt = (([regex]::Matches($pm.Value, '(?s)<w:t[^>]*>(.*?)</w:t>') | ForEach-Object { $_.Groups[1].Value }) -join '')
        if ($ptxt -like ('*' + $AnchorText + '*')) { $pStart = $pm.Index; break }
    }
    if ($pStart -lt 0) { throw "anchor text not found in any paragraph: $AnchorText" }
} else {
    # No anchor: append right before the final paragraph that carries the section properties.
    $pStart = $docXml.LastIndexOf('<w:p ')
    if ($pStart -lt 0) { $pStart = $docXml.LastIndexOf('<w:p>') }
    if ($pStart -lt 0) { throw "could not locate an insertion point" }
}
$docXml = $docXml.Substring(0, $pStart) + $block.ToString() + $docXml.Substring($pStart)
if (-not $docXml.StartsWith('<?xml')) { throw "document.xml lost its XML declaration after the insert" }

$relsXml = $relsXml.Replace('</Relationships>', $relAdds.ToString() + '</Relationships>')
$entries['word/_rels/document.xml.rels'] = [System.Text.Encoding]::UTF8.GetBytes($relsXml)
if ($ctXml -notmatch 'Extension="jpeg"') {
    # Default entries must come before Override entries in [Content_Types].xml.
    $ctXml = $ctXml.Replace('<Default Extension="png"', '<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png"')
    if ($ctXml -notmatch 'Extension="jpeg"') {
        $ctXml = $ctXml.Replace('<Default Extension="xml"', '<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="xml"')
    }
    $entries['[Content_Types].xml'] = [System.Text.Encoding]::UTF8.GetBytes($ctXml)
}
$entries['word/document.xml'] = [System.Text.Encoding]::UTF8.GetBytes($docXml)

$tmp = $Out + '.tmp'
if (Test-Path $tmp) { Remove-Item $tmp -Force }
$fs = [IO.File]::Open($tmp, [IO.FileMode]::CreateNew)
$za = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($k in $order) { $ze = $za.CreateEntry($k, [System.IO.Compression.CompressionLevel]::Optimal); $os = $ze.Open(); $os.Write($entries[$k], 0, $entries[$k].Length); $os.Close() }
    foreach ($k in $newMedia.Keys) { $ze = $za.CreateEntry($k, [System.IO.Compression.CompressionLevel]::Optimal); $os = $ze.Open(); $os.Write($newMedia[$k], 0, $newMedia[$k].Length); $os.Close() }
} finally { $za.Dispose(); $fs.Close() }
if (Test-Path $Out) { Remove-Item $Out -Force }
Move-Item $tmp $Out

[pscustomobject]@{
    ok        = $true
    out       = $Out
    photos    = $cnt
    rows      = @($rows).Count
    template  = $(if ($hasOwn) { "cloned-from-document" } else { "derived-from-own-page-geometry" })
    rowHeight = $imgH
    colWidth  = [Math]::Round($geo.colW, 2)
    boxW      = [Math]::Round($boxW, 2)
    tabs      = "$tabL/$tabR"
} | ConvertTo-Json -Compress
