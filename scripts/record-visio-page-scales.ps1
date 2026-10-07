# Drawing/page scale oracle using a separate invisible Visio instance.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('visio-page-scales-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$app = New-Object -ComObject Visio.InvisibleApp
$cases = @()
try {
    $document = $app.Documents.Add('')
    foreach ($pair in @(@(0.5,1), @(1,1), @(2,1), @(4,1), @(2,0.5), @(2,2))) {
        $drawing = $pair[0]; $paper = $pair[1]
        $page = $document.Pages.Add()
        $page.PageSheet.CellsU('DrawingScale').FormulaU = "$drawing in"
        $page.PageSheet.CellsU('PageScale').FormulaU = "$paper in"
        $line = $page.DrawLine(1,1,3,1)
        $line.CellsU('EndArrow').FormulaU = '4'
        $line.CellsU('EndArrowSize').FormulaU = '2'
        $line.CellsU('LineWeight').FormulaU = '0.01 in'
        $rectangle = $page.DrawRectangle(1,2,5,4)
        $rectangle.CellsU('LineWeight').FormulaU = '0.01 in'
        $rectangle.CellsU('Char.Size').FormulaU = '9 pt'
        $rectangle.CellsU('Char.Font').FormulaU = 'FONT("Arial")'
        $rectangle.CellsU('Para.HorzAlign').FormulaU = '0'
        $rectangle.CellsU('Para.IndLeft').FormulaU = '0.15 in'
        $rectangle.CellsU('LeftMargin').FormulaU = '0.1 in'
        $rectangle.Text = 'Scale'
        $rounded = $page.DrawRectangle(0,0,4,4)
        $rounded.DeleteSection(10); $rounded.AddSection(10) | Out-Null
        $points = @(@(0,0), @(4,0), @(4,4))
        for ($i=0; $i -lt 3; $i++) {
            $tag = if ($i -eq 0) {138} else {139}
            $row = $rounded.AddRow(10,$i+1,$tag)
            $rounded.CellsSRC(10,$row,0).FormulaU = "$($points[$i][0]) in"
            $rounded.CellsSRC(10,$row,1).FormulaU = "$($points[$i][1]) in"
        }
        $rounded.CellsU('Geometry1.NoFill').FormulaU = '1'
        $rounded.CellsU('FillPattern').FormulaU = '0'
        $rounded.CellsU('Rounding').FormulaU = '0.2 in'
        $roundedFile = Join-Path $OutputDirectory "rounded-$drawing-$paper.svg"
        $rounded.Export($roundedFile)
        $roundedSvg = Get-Content $roundedFile -Raw
        $roundedPath = [regex]::Matches($roundedSvg, '<path d="([^"]+)"') | Where-Object { -not $_.Groups[1].Value.EndsWith('Z') } | Select-Object -Last 1
        $lineFile = Join-Path $OutputDirectory "line-$drawing-$paper.svg"
        $rectangleFile = Join-Path $OutputDirectory "rectangle-$drawing-$paper.svg"
        $line.Export($lineFile); $rectangle.Export($rectangleFile)
        $pageFile = Join-Path $OutputDirectory "page-$drawing-$paper.svg"
        $page.Export($pageFile)
        $lineSvg = Get-Content $lineFile -Raw
        $rectangleSvg = Get-Content $rectangleFile -Raw
        $pageSvg = Get-Content $pageFile -Raw
        $nativeRect = [regex]::Match($rectangleSvg, '<rect [^>]*width="([^"]+)" height="([^"]+)"')
        $nativeOrigin = [regex]::Match($pageSvg, "<g id=`"shape$($rectangle.ID)-[^`"]+`"[^>]*transform=`"translate\(([^,]+),([^\)]+)\)")
        $cases += [ordered]@{
            pageId=[string]$page.ID; drawingScale=$drawing; pageScale=$paper
            lineId=[string]$line.ID; rectangleId=[string]$rectangle.ID
            roundedId=[string]$rounded.ID
            paperWidth=$page.PageSheet.CellsU('PageWidth').ResultIU
            lineWidth=$line.CellsU('Width').ResultIU
            rectangleWidth=$rectangle.CellsU('Width').ResultIU; rectangleHeight=$rectangle.CellsU('Height').ResultIU
            fontSize=$rectangle.CellsU('Char.Size').ResultIU
            leftMargin=$rectangle.CellsU('LeftMargin').ResultIU; indentLeft=$rectangle.CellsU('Para.IndLeft').ResultIU
            nativeLinePath=[regex]::Match($lineSvg, '<path d="([^"]+)" class="st1"').Groups[1].Value
            nativeLineStroke=[double]::Parse([regex]::Match($lineSvg, 'st1 \{[^}]*stroke-width:([^}]+)').Groups[1].Value, [cultureinfo]::InvariantCulture) / 72
            nativeRectangleWidth=[double]::Parse($nativeRect.Groups[1].Value, [cultureinfo]::InvariantCulture) / 72
            nativeRectangleHeight=[double]::Parse($nativeRect.Groups[2].Value, [cultureinfo]::InvariantCulture) / 72
            nativeRectangleX=[double]::Parse($nativeOrigin.Groups[1].Value, [cultureinfo]::InvariantCulture) / 72
            nativeRectangleY=-[double]::Parse($nativeOrigin.Groups[2].Value, [cultureinfo]::InvariantCulture) / 72
            nativePaperWidth=[double]::Parse([regex]::Match($pageSvg, 'width="([0-9.]+)in"').Groups[1].Value, [cultureinfo]::InvariantCulture)
            nativeTextX=[double]::Parse([regex]::Match($rectangleSvg, '<text x="([^"]+)"').Groups[1].Value, [cultureinfo]::InvariantCulture) / 72
            nativeRoundedRadius=[double]::Parse([regex]::Match($roundedPath.Groups[1].Value, 'A([0-9.]+) ').Groups[1].Value, [cultureinfo]::InvariantCulture) / 72
        }
    }
    $document.SaveAs((Join-Path $OutputDirectory 'page-scales.vsdx')) | Out-Null
    [ordered]@{application='Microsoft Visio';version=$app.Version;cases=$cases} |
        ConvertTo-Json -Depth 5 | Set-Content (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
} finally {
    if ($document) { $document.Saved=$true; $document.Close() }
    $app.Quit()
}
Write-Output $OutputDirectory
