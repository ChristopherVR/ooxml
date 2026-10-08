# Reopen only owned factory fixtures; no existing application or clipboard is used.
# First run: bun scripts/create-visio-new-drawing-fixtures.ts <OutputDirectory>
param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$cases = @(Get-Content -LiteralPath (Join-Path $directory 'cases.json') -Raw | ConvertFrom-Json)
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
function Assert-Near($actual, $expected, $label) {
    if ([Math]::Abs([double]$actual - [double]$expected) -gt 1e-10) { throw "Mismatch: $label" }
}
function Read-Shape($doc, $shape) {
    $values = [ordered]@{ text=$shape.Text; font=$doc.Fonts.ItemFromID([int]$shape.CellsU('Char.Font').ResultIU).Name }
    foreach ($name in @('PinX','PinY','Width','Height','Char.Size','LineWeight','LineColor','FillForegnd','FillPattern','LeftMargin')) {
        $values[$name] = [double]$shape.CellsU($name).ResultIU
    }
    $values
}
try {
    $application.AlertResponse = 7
    $observations = @()
    foreach ($case in $cases) {
        foreach ($kind in @('blank','drawn')) {
            $name = $case.name + '-' + $kind
            $document = $application.Documents.OpenEx((Join-Path $directory ($name + '.vsdx')),202)
            if ($document.Pages.Count -ne 1) { throw 'Expected one editable page.' }
            $page = $document.Pages.Item(1)
            Assert-Near $page.PageSheet.CellsU('PageWidth').ResultIU $case.width 'page width'
            Assert-Near $page.PageSheet.CellsU('PageHeight').ResultIU $case.height 'page height'
            Assert-Near $page.PageSheet.CellsU('PageScale').ResultIU 1 'page scale'
            Assert-Near $page.PageSheet.CellsU('DrawingScale').ResultIU 1 'drawing scale'
            $expectedCount = if ($kind -eq 'blank') { 0 } else { 1 }
            if ($page.Shapes.Count -ne $expectedCount) { throw 'Unexpected shape count.' }
            if ($kind -eq 'blank') {
                $shape = $page.DrawRectangle(1,2.5,3,3.5)
                $shape.Text = 'New diagram'
            } else { $shape = $page.Shapes.Item(1) }
            $observed = Read-Shape $document $shape
            if ($observed.font -cne 'Calibri' -or $observed.text -cne 'New diagram') { throw 'Default font or text differs.' }
            $expected = @{ PinX=2; PinY=3; Width=2; Height=1; 'Char.Size'=(12/72); LineWeight=(0.75/72); LineColor=0; FillForegnd=1; FillPattern=1; LeftMargin=(4/72) }
            foreach ($key in $expected.Keys) { Assert-Near $observed[$key] $expected[$key] $key }
            [void]$document.SaveAs((Join-Path $directory ($name + '-native.vsdx')))
            $observations += [ordered]@{ name=$name; width=$case.width; height=$case.height; initialShapes=$expectedCount; shape=$observed }
            $document.Saved = $true
            $document.Close()
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
            $document = $null
        }
    }
    [ordered]@{ application='Microsoft Visio'; version=$application.Version; cases=$observations } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'evidence.json') -Encoding utf8
    Write-Output $directory
} finally {
    if ($document) {
        $document.Saved = $true
        $document.Close()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    $application.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
