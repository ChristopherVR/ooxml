# Native source-preserving duplication reference, using only an owned invisible instance.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-duplicate-' + [guid]::NewGuid().ToString('N'))),
    [string]$CoreOutputPath
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
function Read-Shapes($doc) {
    @($doc.Pages | ForEach-Object {
        $page = $_
        [ordered]@{ pageId=[string]$page.ID; shapes=@($page.Shapes | ForEach-Object {
            $shape = $_
            $values = [ordered]@{ id=[string]$shape.ID; name=$shape.NameU; text=$shape.Text }
            foreach ($name in @('PinX','PinY','Width','Height','Angle','User.Link','User.Self','LockMoveX','LockMoveY')) {
                $values[$name] = [double]$shape.CellsU($name).ResultIU
            }
            $values
        }) }
    })
}
try {
    $application.AlertResponse = 7
    if ($CoreOutputPath) {
        $reference = Get-Content -LiteralPath (Join-Path $directory 'evidence.json') -Raw | ConvertFrom-Json
        $document = $application.Documents.OpenEx((Resolve-Path -LiteralPath $CoreOutputPath).Path,202)
        $actual = @(Read-Shapes $document)
        if ($actual.Count -ne $reference.native.Count) { throw 'Reopened page count differs.' }
        for ($pageIndex = 0; $pageIndex -lt $actual.Count; $pageIndex++) {
            $observed = $actual[$pageIndex]
            $expected = $reference.native[$pageIndex]
            if ($observed.pageId -cne $expected.pageId -or $observed.shapes.Count -ne $expected.shapes.Count) {
                throw "Native reopen mismatch: page $pageIndex identity/count"
            }
            for ($index = 0; $index -lt $observed.shapes.Count; $index++) {
                $shape = $observed.shapes[$index]
                $native = $expected.shapes[$index]
                foreach ($name in @('id','name','text')) {
                    if ($shape[$name] -cne $native.$name) { throw "Shape mismatch: page $pageIndex/$index/$name" }
                }
                foreach ($name in @('PinX','PinY','Width','Height','Angle','User.Link','User.Self','LockMoveX','LockMoveY')) {
                    if ([Math]::Abs($shape[$name] - $native.$name) -gt 1e-10) {
                        throw "Native reopen mismatch: page $pageIndex/$index/$name"
                    }
                }
            }
        }
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; accepted=$actual } |
            ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'reopen-evidence.json') -Encoding utf8
    } else {
        $document = $application.Documents.Add('')
        $cases = @()
        foreach ($scale in @(1,2,0.5)) {
            foreach ($mode in @('single','reverse','locked')) {
                $page = if ($document.Pages.Item(1).Shapes.Count -eq 0) { $document.Pages.Item(1) } else { $document.Pages.Add() }
                $page.Name = "$mode-$scale"
                $page.PageSheet.CellsU('PageScale').FormulaU = '1 in'
                $page.PageSheet.CellsU('DrawingScale').FormulaU = ([string]$scale) + ' in'
                $first = $page.DrawRectangle(1,1,3,2)
                $first.NameU = 'Custom.47'
                $first.Text = 'Alpha Beta'
                $first.CellsU('Angle').FormulaU = '15 deg'
                $range = $first.Characters
                $range.Begin = 6; $range.End = 10; $range.CharProps(2) = 2
                [void]$first.UniqueID(1)
                $second = $page.DrawOval(4,3,6,4)
                $second.NameU = 'Source beta'
                $second.Text = 'Beta'
                $third = $page.DrawRectangle(7,5,8,7)
                $third.Text = 'Untouched'
                foreach ($shape in @($first,$second,$third)) {
                    [void]$shape.AddNamedRow(242,'Link',0)
                    [void]$shape.AddNamedRow(242,'Self',0)
                    $shape.CellsU('User.Link').FormulaU = 'Sheet.1!PinX'
                    $shape.CellsU('User.Self').FormulaU = 'Sheet.' + [string]$shape.ID + '!PinY'
                    $shape.CellsU('LockMoveX').FormulaU = '0'
                    $shape.CellsU('LockMoveY').FormulaU = '0'
                }
                if ($mode -eq 'locked') {
                    $first.CellsU('LockMoveX').FormulaU = '1'
                    $first.CellsU('LockMoveY').FormulaU = '1'
                }
                $selectionIds = if ($mode -eq 'single') { @('1') } else { @('2','1') }
                $cases += [ordered]@{ pageId=[string]$page.ID; selection=@($selectionIds) }
            }
        }
        $source = @(Read-Shapes $document)
        [void]$document.SaveAs((Join-Path $directory 'original.vsdx'))
        foreach ($case in $cases) {
            $page = $document.Pages.ItemFromID([int]$case.pageId)
            $selection = $page.CreateSelection(0)
            foreach ($id in $case.selection) { $selection.Select($page.Shapes.ItemFromID([int]$id),2) }
            if ([string]$selection.PrimaryItem.ID -ne $case.selection[0]) { throw 'Native selection primary differs.' }
            [void]$selection.Duplicate()
        }
        [void]$document.SaveAs((Join-Path $directory 'native.vsdx'))
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; cases=$cases; source=$source; native=@(Read-Shapes $document) } |
            ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'evidence.json') -Encoding utf8
    }
    Write-Output $directory
} finally {
    if ($null -ne $document) { $document.Saved = $true; $document.Close() }
    $application.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
