# compose_caption.ps1 -- burn a caption line into an image so it can be inserted as ONE picture.
# Optional path: use when the caption does not need to stay searchable/editable text.
# Prints one JSON object to stdout.
#
# Font sizing maths: the caption must LOOK like SizePt points once the composed image is placed at
# InsertWidthCm in the document. The image has Px pixels across that width, so 1 px = InsertWidthCm/Px
# cm. A SizePt glyph is SizePt/72 inch = SizePt*2.54/72 cm, i.e. that many px. GDI+ renders a font of
# S points with an em box of S*(dpi/72) px (Bitmap defaults to 96 dpi), so S = emPx*72/96.
param(
    [Parameter(Mandatory = $true)][string]$Image,
    [Parameter(Mandatory = $true)][string]$Out,
    [Parameter(Mandatory = $true)][string]$CaptionFile,
    [double]$InsertWidthCm = 0,
    [double]$SizePt = 9,
    [string]$FontName = 'SimSun',
    [string]$Color = '#000000',
    [string]$Bg = '#FFFFFF',
    [int]$GapPx = 12,
    [int]$PadPx = 8,
    [string]$Align = 'center',
    [double]$LineHeight = 1.45,
    [double]$DefaultFontPx = 30
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function Hex2Color($hex) {
    $h = ([string]$hex).Trim().TrimStart('#')
    if ($h.Length -lt 6) { return [System.Drawing.Color]::Black }
    return [System.Drawing.Color]::FromArgb(
        [Convert]::ToInt32($h.Substring(0, 2), 16),
        [Convert]::ToInt32($h.Substring(2, 2), 16),
        [Convert]::ToInt32($h.Substring(4, 2), 16))
}

$caption = ''
if (Test-Path -LiteralPath $CaptionFile) {
    $caption = (Get-Content -LiteralPath $CaptionFile -Raw -Encoding UTF8)
    if ($null -eq $caption) { $caption = '' }
    $caption = $caption.TrimEnd("`r", "`n")
}

# Parameter names are $CaptionFile/$Image: avoid $input, it is an automatic variable.
$src = [System.Drawing.Image]::FromFile((Resolve-Path -LiteralPath $Image).Path)
try {
    $pxW = $src.Width
    $pxH = $src.Height

    $fontPx = $DefaultFontPx
    $mapped = $false
    if ($InsertWidthCm -gt 0 -and $pxW -gt 0) {
        $emPx = ($SizePt / 72.0 * 2.54) * ($pxW / $InsertWidthCm)
        if ($emPx -gt 4) { $fontPx = $emPx * (72.0 / 96.0); $mapped = $true }
    }
    if ($fontPx -lt 6) { $fontPx = 6 }

    $font = New-Object System.Drawing.Font($FontName, [single]$fontPx, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
    # Measure on a scratch bitmap so the strip height is right before we allocate the real one.
    $probe = New-Object System.Drawing.Bitmap 8, 8
    $pg = [System.Drawing.Graphics]::FromImage($probe)
    $pg.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $size = $pg.MeasureString($caption, $font, [single]([Math]::Max(16, $pxW - 2 * $PadPx)))
    $pg.Dispose(); $probe.Dispose()

    $stripH = [int][Math]::Ceiling($size.Height * $LineHeight) + 2 * $PadPx
    $totalH = $pxH + $GapPx + $stripH

    $bmp = New-Object System.Drawing.Bitmap $pxW, $totalH
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.Clear((Hex2Color $Bg))
    $g.DrawImage($src, 0, 0, $pxW, $pxH)

    $brush = New-Object System.Drawing.SolidBrush((Hex2Color $Color))
    $textY = [single]($pxH + $GapPx + $PadPx)
    if ($Align -eq 'left') {
        $g.DrawString($caption, $font, $brush, [single]$PadPx, $textY)
    } else {
        $textX = [single]([Math]::Max($PadPx, ($pxW - $size.Width) / 2))
        $g.DrawString($caption, $font, $brush, $textX, $textY)
    }
    $g.Dispose()

    $outDir = Split-Path -Parent $Out
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
    $bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose(); $brush.Dispose(); $font.Dispose()

    [pscustomobject]@{
        ok        = $true
        out       = $Out
        width     = $pxW
        height    = $totalH
        stripH    = $stripH
        fontPt    = [math]::Round($fontPx, 2)
        mappedPt  = $mapped
        insertCm  = $InsertWidthCm
        sizePt    = $SizePt
    } | ConvertTo-Json -Compress
} finally {
    $src.Dispose()
}
