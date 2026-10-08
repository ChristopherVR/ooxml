# Native mixed-run formatting reference. Own and close only this invisible Visio instance.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-rich-formatting-' + [guid]::NewGuid().ToString('N'))),
    [string]$CoreOutputPath
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
function Read-RichText($doc) {
    @($doc.Pages | ForEach-Object {
        $page = $_
        $shape = $page.Shapes.ItemFromID(1)
        $characters = @()
        for ($index = 0; $index -lt $shape.Text.Length; $index++) {
            $range = $shape.Characters
            $range.Begin = $index
            $range.End = $index + 1
            $characterRow = $range.CharPropsRow(0)
            $paragraphRow = $range.ParaPropsRow(0)
            $values = [ordered]@{ text=$range.Text }
            foreach ($entry in @(@('style',2),@('size',7),@('color',1),@('strike',10))) {
                $values[$entry[0]] = [double]$shape.CellsSRC(3,$characterRow,$entry[1]).ResultIU
            }
            $values.font = $doc.Fonts.ItemFromID([int]$shape.CellsSRC(3,$characterRow,0).ResultIU).Name
            foreach ($entry in @(@('indentLeft',1),@('alignment',6),@('bullet',7))) {
                $values[$entry[0]] = [double]$shape.CellsSRC(4,$paragraphRow,$entry[1]).ResultIU
            }
            $characters += $values
        }
        [ordered]@{ pageId=[string]$page.ID; shapeId='1'; text=$shape.Text; characters=$characters }
    })
}
try {
    $application.AlertResponse = 7
    if ($CoreOutputPath) {
        $reference = Get-Content -LiteralPath (Join-Path $directory 'evidence.json') -Raw | ConvertFrom-Json
        $document = $application.Documents.OpenEx((Resolve-Path -LiteralPath $CoreOutputPath).Path,202)
        $actual = @(Read-RichText $document)
        if ($actual.Count -ne $reference.native.Count) { throw 'Reopened page count differs.' }
        for ($pageIndex = 0; $pageIndex -lt $actual.Count; $pageIndex++) {
            $observed = $actual[$pageIndex]
            $expected = $reference.native[$pageIndex]
            if ($observed.pageId -cne $expected.pageId -or $observed.text -cne $expected.text -or
                $observed.characters.Count -ne $expected.characters.Count) {
                throw "Native reopen mismatch: page $pageIndex text/identity"
            }
            for ($index = 0; $index -lt $observed.characters.Count; $index++) {
                $character = $observed.characters[$index]
                $native = $expected.characters[$index]
                foreach ($name in @('text','font')) {
                    if ($character[$name] -cne $native.$name) { throw "Character mismatch: page $pageIndex/$index/$name" }
                }
                foreach ($name in @('style','size','color','strike','indentLeft','alignment','bullet')) {
                    if ([Math]::Abs($character[$name] - $native.$name) -gt 1e-10) {
                        throw "Native reopen mismatch: page $pageIndex/$index/$name"
                    }
                }
            }
        }
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; accepted=$actual } |
            ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'reopen-evidence.json') -Encoding utf8
    } else {
        $document = $application.Documents.Add('')
        foreach ($scale in @(1,2,0.5)) {
            $page = if ($document.Pages.Item(1).Shapes.Count -eq 0) { $document.Pages.Item(1) } else { $document.Pages.Add() }
            $page.Name = 'Rich-' + [string]$scale
            $page.PageSheet.CellsU('PageScale').FormulaU = '1 in'
            $page.PageSheet.CellsU('DrawingScale').FormulaU = ([string]$scale) + ' in'
            $shape = $page.DrawRectangle(1,1,5,4)
            $shape.Text = "Alpha Beta`nSecond line"
            $shape.CellsU('Char.Font').FormulaU = 'FONT("Arial")'
            $shape.CellsU('Char.Size').FormulaU = '12 pt'
            $shape.CellsU('Char.Style').FormulaU = '0'
            $range = $shape.Characters
            $range.Begin = 6; $range.End = 10
            $range.CharProps(2) = 2
            $range.CharProps(7) = 20
            $range = $shape.Characters
            $range.Begin = 11; $range.End = 22
            $range.CharProps(2) = 4
            $range.ParaProps(6) = 2
            $range.ParaProps(1) = 12
            $fontSample = $page.DrawRectangle(6,1,7,2)
            $fontSample.Text = 'Font registration'
            $fontSample.CellsU('Char.Font').FormulaU = 'FONT("Calibri")'
        }
        [void]$document.SaveAs((Join-Path $directory 'original.vsdx'))
        $original = @(Read-RichText $document)
        foreach ($page in $document.Pages) {
            $shape = $page.Shapes.ItemFromID(1)
            # Whole-shape formatting includes the saved terminal paragraph style.
            # Characters.CharCount excludes that terminal marker, so address its row too.
            for ($row = 0; $row -lt $shape.RowCount(3); $row++) {
                $style = $shape.CellsSRC(3,$row,2)
                $style.FormulaU = [string]([int]$style.ResultIU -bor 1)
            }
            $range = $shape.Characters
            $range.CharProps(0) = [int]$page.Shapes.ItemFromID(2).CellsU('Char.Font').ResultIU
            $range.CharProps(7) = 18
            for ($row = 0; $row -lt $shape.RowCount(3); $row++) {
                $shape.CellsSRC(3,$row,1).FormulaU = 'RGB(128,0,128)'
            }
            $range.CharProps(10) = 1
            $range.ParaProps(1) = 18
            $range.ParaProps(6) = 3
            $range.ParaProps(7) = 1
        }
        [void]$document.SaveAs((Join-Path $directory 'native.vsdx'))
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; original=$original; native=@(Read-RichText $document) } |
            ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'evidence.json') -Encoding utf8
    }
    Write-Output $directory
} finally {
    if ($null -ne $document) { $document.Saved = $true; $document.Close() }
    $application.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
