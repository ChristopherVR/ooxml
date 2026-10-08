# Creates owned reference drawings or reopens core output in a separate invisible Visio instance.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-formatting-' + [guid]::NewGuid().ToString('N'))),
    [string]$CoreOutputPath,
    [ValidateSet('front','back','forward','backward')][string]$Order = 'front',
    [switch]$Extended
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
function Read-Formatting($doc) {
    @($doc.Pages | ForEach-Object {
        $page = $_
        $shape = $page.Shapes.ItemU('Format target')
        $values = [ordered]@{}
        foreach ($cell in @('Char.Size','Char.Style','Char.Strikethru','Char.Color',
                            'Para.Bullet','Para.IndLeft','Para.HorzAlign','VerticalAlign',
                            'FillForegnd','FillPattern','LineColor','LineWeight')) {
            $values[$cell] = [double]$shape.CellsU($cell).ResultIU
        }
        [ordered]@{
            pageId = [string]$page.ID
            shapeId = [string]$shape.ID
            text = $shape.Text
            font = $doc.Fonts.ItemFromID([int]$shape.CellsU('Char.Font').ResultIU).Name
            values = $values
            order = @($page.Shapes | ForEach-Object { [string]$_.ID })
        }
    })
}
try {
    $application.AlertResponse = 7
    if ($CoreOutputPath) {
        $reference = Get-Content -LiteralPath (Join-Path $directory 'evidence.json') -Raw | ConvertFrom-Json
        # Read-only, hidden, omitted from recent files, and with macros disabled.
        $document = $application.Documents.OpenEx((Resolve-Path -LiteralPath $CoreOutputPath).Path,202)
        $actual = @(Read-Formatting $document)
        if ($actual.Count -ne $reference.native.Count) { throw 'Reopened page count differs.' }
        for ($index = 0; $index -lt $actual.Count; $index++) {
            $expected = $reference.native[$index]
            $observed = $actual[$index]
            foreach ($key in @('pageId','shapeId','text','font')) {
                if ($observed[$key] -cne $expected.$key) { throw "Native reopen mismatch: page $index $key" }
            }
            foreach ($cell in $observed.values.Keys) {
                if ([Math]::Abs($observed.values[$cell] - $expected.values.$cell) -gt 1e-9) {
                    throw "Native reopen mismatch: page $index $cell"
                }
            }
            if (($observed.order -join ',') -ne ($expected.order -join ',')) {
                throw "Native reopen mismatch: page $index stacking order"
            }
            $document.Pages.Item($index + 1).Export((Join-Path $directory "core-reopened-$index.svg"))
        }
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; accepted=$actual } |
            ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'reopen-evidence.json') -Encoding utf8
    } else {
        $document = $application.Documents.Add('')
        foreach ($scale in @(1,2,0.5)) {
            $page = if ($document.Pages.Item(1).Shapes.Count -eq 0) { $document.Pages.Item(1) } else { $document.Pages.Add() }
            $page.Name = 'Scale-' + [string]$scale
            $page.PageSheet.CellsU('PageScale').FormulaU = '1 in'
            $page.PageSheet.CellsU('DrawingScale').FormulaU = ([string]$scale) + ' in'
            $shape = $page.DrawRectangle(1,1,4,3)
            $shape.NameU = 'Format target'
            $shape.Text = "First paragraph`nSecond paragraph"
            $shape.CellsU('Char.Font').FormulaU = 'FONT("Arial")'
            $shape.CellsU('Char.Size').FormulaU = '10 pt'
            $shape.CellsU('Char.Style').FormulaU = '0'
            $other = $page.DrawRectangle(2,2,5,4)
            $other.Text = 'Calibri font registration'
            $other.CellsU('Char.Font').FormulaU = 'FONT("Calibri")'
            $page.DrawRectangle(3,3,6,5) | Out-Null
            $shape.BringForward()
        }
        $source = @(Read-Formatting $document)
        $document.SaveAs((Join-Path $directory 'original.vsdx')) | Out-Null
        foreach ($page in $document.Pages) {
            $shape = $page.Shapes.ItemU('Format target')
            $shape.CellsU('Char.Font').FormulaU = 'FONT("Calibri")'
            $shape.CellsU('Char.Size').FormulaU = '18 pt'
            $shape.CellsU('Char.Style').FormulaU = '7'
            $shape.CellsU('Para.HorzAlign').FormulaU = '2'
            $shape.CellsU('VerticalAlign').FormulaU = '0'
            $shape.CellsU('FillForegnd').FormulaU = 'RGB(255,0,0)'
            $shape.CellsU('FillPattern').FormulaU = '1'
            $shape.CellsU('LineColor').FormulaU = 'RGB(0,0,255)'
            $shape.CellsU('LineWeight').FormulaU = '2.25 pt'
            if ($Extended) {
                $shape.CellsU('Char.Strikethru').FormulaU = '1'
                $shape.CellsU('Char.Color').FormulaU = 'RGB(128,0,128)'
                $shape.CellsU('Para.Bullet').FormulaU = '1'
                $shape.CellsU('Para.IndLeft').FormulaU = '18 pt'
                $shape.CellsU('Para.HorzAlign').FormulaU = '3'
            }
            switch ($Order) {
                'front' { $shape.BringToFront() }
                'back' { $shape.SendToBack() }
                'forward' { $shape.BringForward() }
                'backward' { $shape.SendBackward() }
            }
            $page.Export((Join-Path $directory ('native-' + [string]$page.ID + '.svg')))
        }
        $native = @(Read-Formatting $document)
        $document.SaveAs((Join-Path $directory 'native.vsdx')) | Out-Null
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; order=$Order;
                    extended=[bool]$Extended; source=$source; native=$native } |
            ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'evidence.json') -Encoding utf8
    }
} finally {
    if ($document) { $document.Saved = $true; $document.Close() }
    $application.Quit()
}
Write-Output $directory
