# Native geometry oracle: own Visio instance, no attachment to an open document.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('visio-tick-arrows-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$app = New-Object -ComObject Visio.InvisibleApp
$rows = @()
try {
    $document = $app.Documents.Add('')
    $page = $document.Pages.Item(1)
    foreach ($code in @(1, 3, 9)) {
    foreach ($weight in @(0.005, 0.01, 0.02)) {
        foreach ($size in 0..6) {
            $shape = $page.DrawLine(0, 0, 2, 0)
            $shape.CellsU('EndArrow').FormulaU = [string]$code
            $shape.CellsU('EndArrowSize').FormulaU = [string]$size
            $shape.CellsU('LineWeight').FormulaU = $weight.ToString([cultureinfo]::InvariantCulture) + ' in'
            $path = Join-Path $OutputDirectory "arrow-$code-$weight-$size.svg"
            $shape.Export($path)
            $xml = Get-Content -LiteralPath $path -Raw
            $glyph = switch ($code) { 1 { 'M 1 -1 L 0 0 L 1 1' } 3 { 'M 2 1 L 0 0 L 2 -1' } 9 { 'M 1 -1 L -1 1' } }
            if (-not $xml.Contains($glyph)) { throw 'Unexpected native open-arrow geometry' }
            $scale = [double]::Parse([regex]::Match($xml, 'scale\(-([^,]+)').Groups[1].Value, [cultureinfo]::InvariantCulture)
            $rows += [ordered]@{ code = $code; size = $size; lineWidth = $weight; nativeScale = $scale; extent = $scale * $weight; nativeGlyph = $glyph }
            $shape.Delete()
        }
    }
    }
    [ordered]@{ application = 'Microsoft Visio'; version = $app.Version; drawingScale = 1; cases = $rows } |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
} finally {
    if ($document) { $document.Saved = $true; $document.Close() }
    $app.Quit()
}
Write-Output $OutputDirectory
