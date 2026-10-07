# Native geometry oracle: own Visio instance, no attachment to an open document.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('visio-open-arrows-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$app = New-Object -ComObject Visio.InvisibleApp
$rows = @()
try {
    $document = $app.Documents.Add('')
    $page = $document.Pages.Item(1)
    foreach ($code in @(1, 3, 7, 9)) {
    foreach ($weight in @(0.005, 0.01, 0.02)) {
        foreach ($size in 0..6) {
            $shape = $page.DrawLine(0, 0, 2, 0)
            $shape.CellsU('EndArrow').FormulaU = [string]$code
            $shape.CellsU('EndArrowSize').FormulaU = [string]$size
            $shape.CellsU('LineWeight').FormulaU = $weight.ToString([cultureinfo]::InvariantCulture) + ' in'
            $path = Join-Path $OutputDirectory "arrow-$code-$weight-$size.svg"
            $shape.Export($path)
            $xml = Get-Content -LiteralPath $path -Raw
            $glyphMatch = [regex]::Match($xml, "(?s)<g id=`"lend$code`">.*?<path\s+d=`"([^`"]+)`"")
            if (-not $glyphMatch.Success) { throw 'Missing native open-arrow geometry' }
            $glyph = [regex]::Replace($glyphMatch.Groups[1].Value.Trim(), '\s+', ' ')
            $scale = [double]::Parse([regex]::Match($xml, 'scale\(-([^,]+)').Groups[1].Value, [cultureinfo]::InvariantCulture)
            $rows += [ordered]@{ code = $code; size = $size; lineWidth = $weight; nativeScale = $scale; extent = $scale * $weight; nativeGlyph = $glyph }
        }
    }
    }
    $document.SaveAs((Join-Path $OutputDirectory 'open-arrows.vsdx'))
    [ordered]@{ application = 'Microsoft Visio'; version = $app.Version; drawingScale = 1; cases = $rows } |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
} finally {
    if ($document) { $document.Saved = $true; $document.Close() }
    $app.Quit()
}
Write-Output $OutputDirectory
