# Native geometry oracle: own Visio instance, no attachment to an open document.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('visio-filled-arrows-' + [guid]::NewGuid())), [switch]$BothEnds)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$app = New-Object -ComObject Visio.InvisibleApp
$rows = @()
try {
    $document = $app.Documents.Add('')
    $page = $document.Pages.Item(1)
    foreach ($code in @(2, 4, 5, 6)) {
    foreach ($weight in @(0.005, 0.01, 0.02)) {
        foreach ($size in 0..6) {
            $shape = $page.DrawLine(0, 0, 2, 0)
            $shape.CellsU('EndArrow').FormulaU = [string]$code
            $shape.CellsU('EndArrowSize').FormulaU = [string]$size
            if ($BothEnds) {
                $shape.CellsU('BeginArrow').FormulaU = [string]$code
                $shape.CellsU('BeginArrowSize').FormulaU = [string]$size
            }
            $shape.CellsU('LineWeight').FormulaU = $weight.ToString([cultureinfo]::InvariantCulture) + ' in'
            $path = Join-Path $OutputDirectory "arrow-$code-$weight-$size.svg"
            $shape.Export($path)
            $xml = Get-Content -LiteralPath $path -Raw
            $glyphMatch = [regex]::Match($xml, "(?s)<g id=`"lend$code`">.*?<path\s+d=`"([^`"]+)`"")
            if (-not $glyphMatch.Success) { throw 'Missing native open-arrow geometry' }
            $glyph = [regex]::Replace($glyphMatch.Groups[1].Value.Trim(), '\s+', ' ')
            $scale = [double]::Parse([regex]::Match($xml, 'scale\(-([^,]+)').Groups[1].Value, [cultureinfo]::InvariantCulture)
            $endSetback = [regex]::Match($xml, 'refX="-([^"]+)"').Groups[1].Value
            $beginSetback = if ($BothEnds) { [double]::Parse([regex]::Match($xml, 'refX="(\d[^" ]*)"').Groups[1].Value, [cultureinfo]::InvariantCulture) * $weight } else { 0 }
            $rows += [ordered]@{ code = $code; size = $size; lineWidth = $weight; nativeScale = $scale; extent = $scale * $weight; nativeGlyph = $glyph; setback = [double]::Parse($endSetback, [cultureinfo]::InvariantCulture) * $weight; beginSetback = $beginSetback; linePath = [regex]::Match($xml, '<path d="(M[^"]+)" class="st1"').Groups[1].Value }
        }
    }
    }
    $document.SaveAs((Join-Path $OutputDirectory 'filled-arrows.vsdx'))
    [ordered]@{ application = 'Microsoft Visio'; version = $app.Version; drawingScale = 1; cases = $rows } |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
} finally {
    if ($document) { $document.Saved = $true; $document.Close() }
    $app.Quit()
}
Write-Output $OutputDirectory
