# Native short, one-ended connector oracle. Own invisible Visio instance only.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('visio-short-arrows-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$app = New-Object -ComObject Visio.InvisibleApp
$cases = @()
try {
    $document = $app.Documents.Add('')
    $page = $document.Pages.Item(1)
    foreach ($code in @(2, 4, 5, 6)) {
        foreach ($size in @(0, 2, 6)) {
            foreach ($weight in @(0.005, 0.01)) {
                $base = @(0.02, 0.025, 0.035, 0.045, 0.055, 0.125, 0.25)[$size] + $weight
                $factor = switch ($code) { 2 {1} 4 {2} 5 {1.75} 6 {2.25} }
                foreach ($ratio in @(0.25, 0.5, 0.99, 1, 1.01, 1.25, 2)) {
                    $length = $base * $factor * $ratio
                    $shape = $page.DrawLine(0, 0, $length, 0)
                    $shape.CellsU('EndArrow').FormulaU = [string]$code
                    $shape.CellsU('EndArrowSize').FormulaU = [string]$size
                    $shape.CellsU('LineWeight').FormulaU = $weight.ToString([cultureinfo]::InvariantCulture) + ' in'
                    $file = Join-Path $OutputDirectory "arrow-$code-$size-$weight-$ratio.svg"
                    $shape.Export($file)
                    $xml = Get-Content -LiteralPath $file -Raw
                    $reference = [regex]::Match($xml, 'refX="-([^"]+)"').Groups[1].Value
                    $setback = [double]::Parse($reference, [cultureinfo]::InvariantCulture) * $weight
                    $path = [regex]::Match($xml, '<path d="(M[^"]+)" class="st1"').Groups[1].Value
                    $cases += [ordered]@{ code=$code; size=$size; lineWidth=$weight; length=$length; ratio=$ratio; setback=$setback; path=$path }
                }
            }
        }
    }
    $document.SaveAs((Join-Path $OutputDirectory 'short-arrows.vsdx')) | Out-Null
    [ordered]@{ application='Microsoft Visio'; version=$app.Version; drawingScale=1; cases=$cases } |
        ConvertTo-Json -Depth 5 | Set-Content (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
} finally {
    if ($document) { $document.Saved=$true; $document.Close() }
    $app.Quit()
}
Write-Output $OutputDirectory
