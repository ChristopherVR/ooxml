# Native align/distribute reference and core reopen checks, using only an owned invisible instance.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-arrangement-' + [guid]::NewGuid().ToString('N'))),
    [string]$CoreOutputPath
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
function Read-Positions($doc) {
    @($doc.Pages | ForEach-Object {
        [ordered]@{ pageId=[string]$_.ID; name=$_.Name; shapes=@($_.Shapes | ForEach-Object {
            [ordered]@{ id=[string]$_.ID; pinX=$_.CellsU('PinX').ResultIU; pinY=$_.CellsU('PinY').ResultIU;
                        width=$_.CellsU('Width').ResultIU; height=$_.CellsU('Height').ResultIU;
                        angle=$_.CellsU('Angle').ResultIU }
        }) }
    })
}
try {
    $application.AlertResponse = 7
    if ($CoreOutputPath) {
        $reference = Get-Content -LiteralPath (Join-Path $directory 'evidence.json') -Raw | ConvertFrom-Json
        $document = $application.Documents.OpenEx((Resolve-Path -LiteralPath $CoreOutputPath).Path,202)
        $actual = @(Read-Positions $document)
        if ($actual.Count -ne $reference.native.Count) { throw 'Reopened page count differs.' }
        for ($index = 0; $index -lt $actual.Count; $index++) {
            $expected = $reference.native[$index]
            $observed = $actual[$index]
            if ($observed.pageId -ne $expected.pageId) { throw 'Reopened page identity differs.' }
            # Native alignment extents round local dimensions to float32. Bound the two boxes'
            # rounding analytically in drawing inches; keep core coordinates at full precision.
            $rounding = 0.0
            foreach ($shape in $reference.source[$index].shapes) {
                $widthError = [Math]::Abs([double][single]$shape.width - $shape.width)
                $heightError = [Math]::Abs([double][single]$shape.height - $shape.height)
                $cos = [Math]::Abs([Math]::Cos($shape.angle)); $sin = [Math]::Abs([Math]::Sin($shape.angle))
                $rounding = [Math]::Max($rounding,[Math]::Max($cos*$widthError+$sin*$heightError,$sin*$widthError+$cos*$heightError))
            }
            for ($shape = 0; $shape -lt $observed.shapes.Count; $shape++) {
                foreach ($key in @('pinX','pinY')) {
                    if ([Math]::Abs($observed.shapes[$shape][$key] - $expected.shapes[$shape].$key) -gt (2*$rounding+1e-11)) {
                        throw "Native reopen mismatch: page $index shape $shape $key"
                    }
                }
            }
        }
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; accepted=$actual } |
            ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'reopen-evidence.json') -Encoding utf8
    } else {
        $document = $application.Documents.Add('')
        $cases = @()
        foreach ($scale in @(1,2,0.5)) {
            foreach ($action in @('left','center','right','top','middle','bottom','horizontal','vertical')) {
                $page = if ($document.Pages.Item(1).Shapes.Count -eq 0) { $document.Pages.Item(1) } else { $document.Pages.Add() }
                $page.Name = "$action-$scale"
                $page.PageSheet.CellsU('PageScale').FormulaU = '1 in'
                $page.PageSheet.CellsU('DrawingScale').FormulaU = ([string]$scale) + ' in'
                $first = $page.DrawRectangle(1,1,3,2)
                $first.CellsU('Angle').FormulaU = '-15 deg'
                $second = $page.DrawRectangle(4,3,5.6,4.2)
                $second.CellsU('Angle').FormulaU = '25 deg'
                $third = $page.DrawRectangle(7,5,8,7)
                $cases += [ordered]@{ pageId=[string]$page.ID; action=$action; selection=@('2','1','3') }
            }
        }
        $source = @(Read-Positions $document)
        $document.SaveAs((Join-Path $directory 'original.vsdx')) | Out-Null
        foreach ($case in $cases) {
            $page = $document.Pages.ItemFromID([int]$case.pageId)
            $selection = $page.CreateSelection(0)
            foreach ($id in $case.selection) { $selection.Select($page.Shapes.ItemFromID([int]$id),2) }
            if ([string]$selection.PrimaryItem.ID -ne $case.selection[0]) { throw 'Native selection primary differs.' }
            switch ($case.action) {
                'left' { $selection.Align(1,0,$false) }
                'center' { $selection.Align(2,0,$false) }
                'right' { $selection.Align(3,0,$false) }
                'top' { $selection.Align(0,1,$false) }
                'middle' { $selection.Align(0,2,$false) }
                'bottom' { $selection.Align(0,3,$false) }
                'horizontal' { $selection.Distribute(0,$false) }
                'vertical' { $selection.Distribute(4,$false) }
            }
        }
        $native = @(Read-Positions $document)
        $document.SaveAs((Join-Path $directory 'native.vsdx')) | Out-Null
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; cases=$cases; source=$source; native=$native } |
            ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'evidence.json') -Encoding utf8
    }
} finally {
    if ($document) { $document.Saved = $true; $document.Close() }
    $application.Quit()
}
Write-Output $directory
