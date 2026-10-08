# Native SVG references from an owned invisible Visio instance. Does not use the clipboard.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-line-patterns-' + [guid]::NewGuid().ToString('N'))),
    [double[]]$Weights = @(1, 3),
    [double[]]$Scales = @(1, 0.5, 2)
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
if (Test-Path -LiteralPath (Join-Path $directory 'paint.vsdx')) { throw 'Reference output already exists.' }
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
try {
    $application.AlertResponse = 7
    $application.EventsEnabled = 0
    $document = $application.Documents.Add('')
    $cases = [Collections.Generic.List[object]]::new()
    foreach ($scale in $Scales) {
        foreach ($cap in @(0, 1, 2)) {
            foreach ($weight in $Weights) {
                foreach ($pattern in 0..23) {
                    $page = if ($cases.Count -eq 0) { $document.Pages.Item(1) } else { $document.Pages.Add() }
                    $page.Name = 'Line-' + $cases.Count
                    $page.PageSheet.CellsU('DrawingScale').FormulaU = '1 in'
                    $page.PageSheet.CellsU('PageScale').FormulaU = $scale.ToString([Globalization.CultureInfo]::InvariantCulture) + ' in'
                    $shape = $page.DrawRectangle(1 / $scale, 1 / $scale, 4 / $scale, 2 / $scale)
                    $shape.CellsU('LineWeight').FormulaU = $weight.ToString([Globalization.CultureInfo]::InvariantCulture) + ' pt'
                    $shape.CellsU('LineCap').FormulaU = [string]$cap
                    $shape.CellsU('LinePattern').FormulaU = [string]$pattern
                    $shape.CellsU('LineColor').FormulaU = 'RGB(0,0,0)'
                    $shape.CellsU('FillPattern').FormulaU = '0'
                    $path = Join-Path $directory ($page.Name + '.svg')
                    $page.Export($path)
                    $svg = [Xml.XmlDocument]::new()
                    $svg.XmlResolver = $null
                    $svg.Load($path)
                    $cases.Add([ordered]@{
                        pattern = $pattern; cap = $cap; weight = $weight; scale = $scale
                        pageId = [string]$page.ID; shapeId = [string]$shape.ID
                        lineWidth = [double]$shape.CellsU('LineWeight').ResultIU
                        svgStyle = ($svg.SelectNodes('//*[local-name()="style"]') | ForEach-Object { $_.InnerText }) -join ''
                    })
                }
            }
        }
    }
    $document.SaveAs((Join-Path $directory 'paint.vsdx')) | Out-Null
    [ordered]@{ version = $application.Version; lines = @($cases.ToArray()) } |
        ConvertTo-Json -Depth 8 |
        Set-Content -LiteralPath (Join-Path $directory 'evidence.json') -Encoding utf8
    Write-Output $directory
} finally {
    if ($document) { $document.Saved = $true; $document.Close() }
    $application.Quit()
    if ($document) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
