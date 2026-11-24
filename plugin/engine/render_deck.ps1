# Render a pptx into per-slide PNGs (PowerPoint COM).
# ASCII only on purpose: Windows PowerShell 5.1 reads .ps1 as ANSI without a BOM,
# so non-ASCII comments corrupt the script.
# Usage: powershell -NoProfile -File render_deck.ps1 -Pptx <file> -OutDir <dir> [-Width 1600] [-Height 900]
param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [int]$Width = 1600,
    [int]$Height = 900
)

$ErrorActionPreference = "Stop"
if (-not (Test-Path -LiteralPath $Pptx)) { throw "pptx not found: $Pptx" }
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    # Two-argument Open only: Open(..., WithWindow=$false) fails on this build.
    $pres = $ppt.Presentations.Open($Pptx, $true)
    $n = $pres.Slides.Count
    for ($i = 1; $i -le $n; $i++) {
        $png = Join-Path $OutDir ("p{0:d2}.PNG" -f $i)
        if (Test-Path -LiteralPath $png) { Remove-Item -LiteralPath $png -Force }
        $pres.Slides.Item($i).Export($png, "PNG", $Width, $Height)
    }
    $pres.Close()
    $pres = $null
    Write-Output "rendered $n slides -> $OutDir"
}
finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($ppt) { try { $ppt.Quit() } catch {} }
}
