# Owned native Selection.Delete oracle. This script never uses the system clipboard.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-batch-delete-' + [guid]::NewGuid().ToString('N'))),
    [string]$CoreOutputPath,
    [switch]$Restored
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
        if ($Restored) {
            # The captured pair is restored at its original pins with fresh IDs after the control.
            foreach ($page in $reference.source) {
                $control = $page.shapes[2]
                $nextId = [int]$control.id
                foreach ($shape in @($page.shapes[0],$page.shapes[1])) {
                    $nextId++
                    $shape.id = [string]$nextId
                    $shape.name = ($shape.name -replace '\.\d+$','') + '.' + [string]$nextId
                }
                $page.shapes = @($control,$page.shapes[0],$page.shapes[1])
            }
            $reference.native = $reference.source
        }
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
        $evidenceName = if ($Restored) { 'cut-paste-reopen-evidence.json' } else { 'reopen-evidence.json' }
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; accepted=$actual } |
            ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory $evidenceName) -Encoding utf8
    } else {
        $document = $application.Documents.Add('')
        $cases = @()
        foreach ($scale in @(1,2,0.5)) {
            foreach ($mode in @('forward','reverse','movement-locked')) {
                $page = if ($document.Pages.Item(1).Shapes.Count -eq 0) { $document.Pages.Item(1) } else { $document.Pages.Add() }
                $page.Name = "$mode-$scale"
                $page.PageSheet.CellsU('PageScale').FormulaU = '1 in'
                $page.PageSheet.CellsU('DrawingScale').FormulaU = ([string]$scale) + ' in'
                $first = $page.DrawRectangle(1,1,3,2)
                $first.Text = 'Alpha'
                $second = $page.DrawOval(4,3,6,4)
                $second.Text = 'Beta'
                $third = $page.DrawRectangle(7,5,8,7)
                $third.Text = 'Untouched rich text'
                $range = $third.Characters
                $range.Begin = 10; $range.End = 14; $range.CharProps(2) = 2
                foreach ($shape in @($first,$second,$third)) {
                    [void]$shape.AddNamedRow(242,'Link',0)
                    [void]$shape.AddNamedRow(242,'Self',0)
                    $targetId = if ($shape.ID -eq $first.ID) { $second.ID } elseif ($shape.ID -eq $second.ID) { $first.ID } else { $third.ID }
                    $shape.CellsU('User.Link').FormulaU = 'Sheet.' + [string]$targetId + '!PinX'
                    $shape.CellsU('User.Self').FormulaU = 'Sheet.' + [string]$shape.ID + '!PinY'
                    $shape.CellsU('LockMoveX').FormulaU = '0'
                    $shape.CellsU('LockMoveY').FormulaU = '0'
                }
                if ($mode -eq 'movement-locked') {
                    $first.CellsU('LockMoveX').FormulaU = '1'
                    $first.CellsU('LockMoveY').FormulaU = '1'
                }
                $selectionIds = if ($mode -eq 'reverse') { @([string]$second.ID,[string]$first.ID) } else { @([string]$first.ID,[string]$second.ID) }
                $cases += [ordered]@{ pageId=[string]$page.ID; selection=@($selectionIds); control=[string]$third.ID }
            }
        }
        $source = @(Read-Shapes $document)
        [void]$document.SaveAs((Join-Path $directory 'original.vsdx'))
        foreach ($case in $cases) {
            $page = $document.Pages.ItemFromID([int]$case.pageId)
            $selection = $page.CreateSelection(0)
            foreach ($id in $case.selection) { $selection.Select($page.Shapes.ItemFromID([int]$id),2) }
            $selection.Delete()
            if ($page.Shapes.Count -ne 1 -or [string]$page.Shapes.Item(1).ID -cne $case.control) { throw 'Native closed-set deletion differs.' }
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
