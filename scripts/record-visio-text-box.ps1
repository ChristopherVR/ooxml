# Records fixed text-box API references in an owned, invisible Visio instance. No clipboard access.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-text-box-' + [guid]::NewGuid().ToString('N'))),
    [switch]$CustomTextStyle,
    [string]$CoreOutputPath
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
function Read-Shape($page) {
    $shape = $page.Shapes.Item(1)
    $values = [ordered]@{}
    foreach ($name in @('PinX','PinY','Width','Height','LocPinX','LocPinY','Angle',
                        'LinePattern','FillPattern','VerticalAlign','Para.HorzAlign',
                        'Char.Size','Char.Style','LeftMargin','RightMargin','TopMargin','BottomMargin')) {
        $cell = $shape.CellsU($name)
        $values[$name] = [ordered]@{ value=[double]$cell.ResultIU; formula=$cell.FormulaU }
    }
    $x0 = 0.0; $y0 = 0.0; $xx = 0.0; $yx = 0.0; $xy = 0.0; $yy = 0.0
    $shape.XYToPage(0,0,[ref]$x0,[ref]$y0)
    $shape.XYToPage(1,0,[ref]$xx,[ref]$yx)
    $shape.XYToPage(0,1,[ref]$xy,[ref]$yy)
    [ordered]@{
        pageId=[string]$page.ID; shapeId=[string]$shape.ID; text=$shape.Text
        values=$values; transform=@(($xx-$x0),($yx-$y0),($xy-$x0),($yy-$y0),$x0,$y0)
        lineStyle=$shape.LineStyle; fillStyle=$shape.FillStyle; textStyle=$shape.TextStyle
        font=$document.Fonts.ItemFromID([int]$shape.CellsU('Char.Font').ResultIU).Name
    }
}
try {
    $application.AlertResponse = 7
    if ($CoreOutputPath) {
        $reference = Get-Content -LiteralPath (Join-Path $directory 'evidence.json') -Raw | ConvertFrom-Json
        $document = $application.Documents.OpenEx((Resolve-Path -LiteralPath $CoreOutputPath).Path,202)
        if ($document.Pages.Count -ne $reference.cases.Count) { throw 'Reopened page count differs.' }
        $actual = @()
        for ($index = 0; $index -lt $document.Pages.Count; $index++) {
            $observed = Read-Shape $document.Pages.Item($index + 1)
            $expected = $reference.cases[$index].native
            foreach ($name in @('pageId','shapeId','text','font','lineStyle','fillStyle','textStyle')) {
                if ($observed[$name] -cne $expected.$name) { throw "Native reopen mismatch: $index $name" }
            }
            foreach ($name in $observed.values.Keys) {
                if ([Math]::Abs($observed.values[$name].value - $expected.values.$name.value) -gt 1e-10) {
                    throw "Native reopen mismatch: $index $name"
                }
            }
            for ($component = 0; $component -lt 6; $component++) {
                if ([Math]::Abs($observed.transform[$component] - $expected.transform[$component]) -gt 1e-10) {
                    throw "Native reopen mismatch: $index transform"
                }
            }
            $actual += $observed
        }
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; accepted=$actual } |
            ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'reopen-evidence.json') -Encoding utf8
    } else {
        $document = $application.Documents.Add('')
        if ($CustomTextStyle) {
            $style = $document.Styles.ItemU('Text Only')
            $style.CellsU('Char.Size').FormulaU = '18 pt'
            $style.CellsU('Char.Font').FormulaU = 'FONT("Arial")'
            $document.DefaultTextStyle = 'Text Only'
        }
        $cases = @()
        foreach ($drawingScale in @(1,2,0.5)) {
            foreach ($text in @('Fixed text box',"First line`nSecond line","Trailing`n`n",'')) {
                $page = if ($cases.Count -eq 0) { $document.Pages.Item(1) } else { $document.Pages.Add() }
                $page.Name = 'Text-' + [string]$drawingScale + '-' + [string]$cases.Count
                $page.PageSheet.CellsU('PageScale').FormulaU = '1 in'
                $page.PageSheet.CellsU('DrawingScale').FormulaU = ([string]$drawingScale) + ' in'
                $cases += [ordered]@{ pageId=[string]$page.ID; drawingScale=$drawingScale; text=$text }
            }
        }
        $document.SaveAs((Join-Path $directory 'original.vsdx'))
        for ($index = 0; $index -lt $cases.Count; $index++) {
            $page = $document.Pages.Item($index + 1)
            $shape = $page.DrawRectangle(1,1,4,2)
            $shape.CellsU('LinePattern').FormulaU = '0'
            $shape.CellsU('FillPattern').FormulaU = '0'
            $shape.Text = $cases[$index].text
            $cases[$index].native = Read-Shape $page
            $page.Export((Join-Path $directory "native-$index.svg"))
        }
        $document.SaveAs((Join-Path $directory 'native.vsdx'))
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; customText=[bool]$CustomTextStyle; cases=$cases } |
            ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'evidence.json') -Encoding utf8
    }
} finally {
    if ($document) { $document.Close() }
    $application.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
Write-Output $directory
