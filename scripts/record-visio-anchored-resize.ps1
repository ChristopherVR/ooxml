# Native handle-resize measurements in owned invisible documents, without clipboard access.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-anchored-resize-' + [guid]::NewGuid().ToString('N'))),
    [string]$CoreOutputPath,
    [ValidateSet('half','quarter','constant','guard-locpin','guard-constant-locpin','offset-locpin','rotated','flip-x','flip-y','guard-pin','guard-width','guard-flip','pin-formula','lock-move','lock-width','lock-aspect','pin-dependent-locpin','dimension-angle','dimension-flip')]
    [string[]]$Modes = @('half','quarter','constant','guard-locpin','guard-constant-locpin','offset-locpin','rotated','flip-x','flip-y','guard-pin','guard-width','guard-flip','pin-formula','lock-move','lock-width','lock-aspect','pin-dependent-locpin','dimension-angle','dimension-flip'),
    [ValidateRange(0,7)][int[]]$Directions = @(0),
    [ValidateSet('rectangle','ellipse')][string[]]$ShapeKinds = @('rectangle'),
    [double[]]$DrawingScales = @(1),
    [double]$Distance = 0.25
)
$ErrorActionPreference = 'Stop'
$directory = [IO.Path]::GetFullPath($OutputDirectory)
if ($CoreOutputPath) {
    if (-not (Test-Path -LiteralPath $directory)) { throw 'Reopen requires an existing oracle directory.' }
} else {
    if (Test-Path -LiteralPath $directory) { throw 'Use a fresh output directory.' }
    New-Item -ItemType Directory -Path $directory | Out-Null
}
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
. (Join-Path $PSScriptRoot 'visio-capture-geometry.ps1')
$cellNames = @('Width','Height','PinX','PinY','LocPinX','LocPinY','Angle','FlipX','FlipY','LockMoveX','LockMoveY','LockWidth','LockHeight','LockAspect')
try {
    $application.AlertResponse = 7
    if ($CoreOutputPath) {
        $cases = Get-Content -LiteralPath (Join-Path $directory 'accepted-cases.json') -Raw | ConvertFrom-Json
        $document = $application.Documents.OpenEx((Resolve-Path -LiteralPath $CoreOutputPath).Path,202)
        $accepted = @()
        foreach ($case in $cases) {
            $shape = $document.Pages.ItemFromID([int]$case.pageId).Shapes.ItemFromID([int]$case.shapeId)
            $actual = Get-ShapeCells $shape $cellNames
            foreach ($name in $cellNames) {
                if ([Math]::Abs($actual[$name].value - $case.after.$name.value) -gt 1e-10) {
                    throw "Reopened native cell differs: $($case.pageId)/$name"
                }
            }
            foreach ($name in @('User.WidthCache','User.PinCache')) {
                if ([Math]::Abs($shape.CellsU($name).ResultIU - $case.afterDependents.$name.value) -gt 1e-10) {
                    throw "Reopened dependent cache differs: $($case.pageId)/$name"
                }
            }
            $matrix = Get-LineTransform $shape
            for ($index = 0; $index -lt 6; $index++) {
                if ([Math]::Abs($matrix[$index] - $case.afterTransform[$index]) -gt 1e-10) {
                    throw "Reopened native transform differs: $($case.pageId)/$index"
                }
            }
            $accepted += [ordered]@{pageId=$case.pageId;mode=$case.mode;direction=$case.direction;cells=$actual;transform=$matrix}
        }
        [IO.File]::WriteAllText((Join-Path $directory 'reopen-evidence.json'),($accepted | ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
        Write-Output $directory
        return
    }
    $document = $application.Documents.Add('')
    $cases = @()
    foreach ($scale in $DrawingScales) {
        if ([double]::IsNaN($scale) -or [double]::IsInfinity($scale) -or $scale -le 0) { throw 'Drawing scale must be finite and positive.' }
        foreach ($mode in $Modes) {
          foreach ($kind in $ShapeKinds) {
            foreach ($direction in $Directions) {
                $page = if ($document.Pages.Item(1).Shapes.Count -eq 0) { $document.Pages.Item(1) } else { $document.Pages.Add() }
                $page.Name = "$mode-$kind-$direction-$scale"
                $page.PageSheet.CellsU('PageScale').FormulaU = '1 in'
                $page.PageSheet.CellsU('DrawingScale').FormulaU = ([string]$scale) + ' in'
                $shape = if ($kind -eq 'ellipse') { $page.DrawOval(1,1,3,2) } else { $page.DrawRectangle(1,1,3,2) }
                $shape.Text = $mode
                switch ($mode) {
                    'quarter' {
                        $shape.CellsU('LocPinX').FormulaU = 'Width*0.25'
                        $shape.CellsU('LocPinY').FormulaU = 'Height*0.75'
                    }
                    'constant' {
                        $shape.CellsU('LocPinX').FormulaU = '0.25 in'
                        $shape.CellsU('LocPinY').FormulaU = '0.75 in'
                    }
                    'guard-locpin' {
                        $shape.CellsU('LocPinX').FormulaU = 'GUARD(Width*0.25)'
                        $shape.CellsU('LocPinY').FormulaU = 'GUARD(Height*0.75)'
                    }
                    'guard-constant-locpin' {
                        $shape.CellsU('LocPinX').FormulaU = 'GUARD(0.25 in)'
                        $shape.CellsU('LocPinY').FormulaU = 'GUARD(0.75 in)'
                    }
                    'offset-locpin' { $shape.CellsU('LocPinX').FormulaU = 'Width*0.25+0.1 in' }
                    'rotated' { $shape.CellsU('Angle').FormulaU = '30 deg' }
                    'flip-x' { $shape.CellsU('FlipX').FormulaU = '1' }
                    'flip-y' { $shape.CellsU('FlipY').FormulaU = '1' }
                    'guard-pin' {
                        $shape.CellsU('PinX').FormulaU = 'GUARD(2 in)'
                        $shape.CellsU('PinY').FormulaU = 'GUARD(1.5 in)'
                    }
                    'guard-width' { $shape.CellsU('Width').FormulaU = 'GUARD(2 in)' }
                    'guard-flip' { $shape.CellsU('FlipX').FormulaU = 'GUARD(1)' }
                    'pin-formula' { $shape.CellsU('PinX').FormulaU = 'Width' }
                    'lock-move' {
                        $shape.CellsU('LockMoveX').FormulaU = '1'
                        $shape.CellsU('LockMoveY').FormulaU = '1'
                    }
                    'lock-width' { $shape.CellsU('LockWidth').FormulaU = '1' }
                    'lock-aspect' { $shape.CellsU('LockAspect').FormulaU = '1' }
                    'pin-dependent-locpin' { $shape.CellsU('LocPinX').FormulaU = 'PinX*0.25' }
                    'dimension-angle' { $shape.CellsU('Angle').FormulaU = 'Width/(1 in)*15 deg' }
                    'dimension-flip' { $shape.CellsU('FlipX').FormulaU = 'IF(Width&gt;2 in,1,0)'.Replace('&gt;','>') }
                }
                [void]$shape.AddNamedRow(242,'WidthCache',0)
                $shape.CellsU('User.WidthCache').FormulaU = 'Width'
                [void]$shape.AddNamedRow(242,'PinCache',0)
                $shape.CellsU('User.PinCache').FormulaU = 'PinX'
                $cases += [ordered]@{
                    pageId=[string]$page.ID; shapeId=[string]$shape.ID; mode=$mode; kind=$kind; direction=$direction;
                    scale=$scale; distance=$Distance; before=(Get-ShapeCells $shape $cellNames);
                    beforeTransform=(Get-LineTransform $shape)
                }
            }
          }
        }
    }
    [void]$document.SaveAs((Join-Path $directory 'original.vsdx'))
    foreach ($case in $cases) {
        $page = $document.Pages.ItemFromID([int]$case.pageId)
        $shape = $page.Shapes.ItemFromID([int]$case.shapeId)
        $selection = $page.CreateSelection(0)
        $selection.Select($shape,2)
        $case['error'] = $null
        try { $selection.Resize($case.direction,$case.distance,65) } catch { $case['error'] = $_.Exception.Message }
        $case['after'] = Get-ShapeCells $shape $cellNames
        $case['afterTransform'] = Get-LineTransform $shape
        $case['afterDependents'] = Get-ShapeCells $shape @('User.WidthCache','User.PinCache')
    }
    [void]$document.SaveAs((Join-Path $directory 'native.vsdx'))
    $evidence = [ordered]@{ application='Microsoft Visio'; version=$application.Version; operation='Selection.Resize'; cases=$cases }
    [IO.File]::WriteAllText((Join-Path $directory 'evidence.json'),($evidence | ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
    Write-Output $directory
} finally {
    if ($null -ne $document) { $document.Saved = $true; $document.Close() }
    $application.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
