# Native vector references only. Raster hairlines use a different device policy.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-svg-zero-stroke-' + [guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
if (Test-Path -LiteralPath (Join-Path $directory 'zero-stroke.vsdx')) { throw 'Reference output already exists.' }
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
try {
    $application.AlertResponse = 7
    $application.EventsEnabled = 0
    $document = $application.Documents.Add('')
    $cases = [Collections.Generic.List[object]]::new()
    foreach ($cap in @(0, 1, 2, 3)) {
        foreach ($weight in @(0, 0.001, 0.75)) {
            $page = if ($cases.Count -eq 0) { $document.Pages.Item(1) } else { $document.Pages.Add() }
            $page.Name = 'Zero-' + $cases.Count
            $shape = $page.DrawLine(1, 2, 3, 2)
            $arrow = $cap -eq 3
            $shape.CellsU('LineCap').FormulaU = [string]$(if ($arrow) { 0 } else { $cap })
            $shape.CellsU('LineWeight').FormulaU = $weight.ToString([Globalization.CultureInfo]::InvariantCulture) + ' pt'
            $shape.CellsU('LinePattern').FormulaU = $(if ($arrow) { '1' } else { '3' })
            $shape.CellsU('LineColor').FormulaU = 'RGB(0,0,0)'
            $shape.CellsU('BeginArrow').FormulaU = $(if ($arrow) { '4' } else { '0' })
            $shape.CellsU('EndArrow').FormulaU = $(if ($arrow) { '4' } else { '0' })
            $shape.CellsU('BeginArrowSize').FormulaU = '2'
            $shape.CellsU('EndArrowSize').FormulaU = '2'
            $path = Join-Path $directory ($page.Name + '.svg')
            $page.Export($path)
            $svg = [Xml.XmlDocument]::new()
            $svg.XmlResolver = $null
            $svg.Load($path)
            $markers = @($svg.SelectNodes('//*[local-name()="marker"]') | ForEach-Object {
                [ordered]@{
                    reference = [double]::Parse($_.GetAttribute('refX'), [Globalization.CultureInfo]::InvariantCulture)
                    transform = $_.SelectSingleNode('*[local-name()="use"]').GetAttribute('transform')
                }
            })
            $cases.Add([ordered]@{
                pageId = [string]$page.ID; shapeId = [string]$shape.ID
                weight = $weight; cap = $(if ($arrow) { 0 } else { $cap }); arrow = $arrow
                lineWidth = [double]$shape.CellsU('LineWeight').ResultIU
                svgStyle = ($svg.SelectNodes('//*[local-name()="style"]') | ForEach-Object { $_.InnerText }) -join ''
                markers = $markers
            })
        }
    }
    $document.SaveAs((Join-Path $directory 'zero-stroke.vsdx')) | Out-Null
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
