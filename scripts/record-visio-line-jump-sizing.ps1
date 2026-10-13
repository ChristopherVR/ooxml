param(
    [Parameter(Mandatory=$true)][string]$Output,
    [string]$Source
)
# Manual Windows probe. Owns and closes an invisible Visio instance, never saves the source.
# Factors and spacing are raw ShapeSheet values; exported SVG uses paper points.
$ErrorActionPreference = 'Stop'
$culture = [Globalization.CultureInfo]::InvariantCulture
$app = New-Object -ComObject Visio.InvisibleApp
$doc = $null
$temporary = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N') + '.svg')
try {
    if ($Source) { $doc = $app.Documents.OpenEx((Resolve-Path -LiteralPath $Source).Path, 66) }
    else {
        $doc = $app.Documents.Add('')
        $page = $doc.Pages.Item(1)
        $h = $page.Drop($app.ConnectorToolDataObject, 1, 1)
        $v = $page.Drop($app.ConnectorToolDataObject, 1, 1)
        foreach ($shape in @($h, $v)) { $shape.CellsU('ShapeRouteStyle').FormulaU = '16' }
        foreach ($entry in @(@('BeginX',1),@('BeginY',2),@('EndX',5),@('EndY',2))) {
            $h.CellsU($entry[0]).FormulaU = "$($entry[1]) in"
        }
        foreach ($entry in @(@('BeginX',3),@('BeginY',0),@('EndX',3),@('EndY',4))) {
            $v.CellsU($entry[0]).FormulaU = "$($entry[1]) in"
        }
    }
    $page = $doc.Pages.Item(1)
    $h = $page.Shapes.ItemFromID(1); $v = $page.Shapes.ItemFromID(2)
    foreach ($shape in @($h,$v)) {
        $shape.CellsU('ConLineJumpCode').FormulaU = '0'
        $shape.CellsU('ConLineJumpStyle').FormulaU = '0'
    }
    $page.PageSheet.CellsU('LineJumpStyle').FormulaU = '1'
    # direction,factorX,factorY,spacingX,spacingY,pageScale,drawingScale,weightPt
    $settings = @()
    foreach ($direction in @('h','v')) {
        foreach ($weight in @(1,3,6)) {
            foreach ($factor in @(0,0.5,(2.0/3.0),1,2)) {
                $settings += ,@($direction,$factor,$factor,0.125,0.125,1,1,$weight)
            }
        }
    }
    foreach ($scale in @(@(1,1),@(1,2),@(2,1))) {
        foreach ($spacing in @(0.0625,0.125,0.25)) {
            foreach ($weight in @(1,6)) {
                foreach ($direction in @('h','v')) {
                    $settings += ,@($direction,(2.0/3.0),(2.0/3.0),$spacing,$spacing,$scale[0],$scale[1],$weight)
                }
            }
        }
    }
    foreach ($direction in @('h','v')) { $settings += ,@($direction,0.5,2,0.125,0.25,1,1,6) }
    $cases = @()
    foreach ($setting in $settings) {
        $direction,$factorX,$factorY,$spacingX,$spacingY,$pageScale,$drawingScale,$weight = $setting
        foreach ($entry in @(@('PageScale',$pageScale),@('DrawingScale',$drawingScale),@('LineToLineX',$spacingX),@('LineToLineY',$spacingY))) {
            $page.PageSheet.CellsU($entry[0]).FormulaU = ([double]$entry[1]).ToString('R',$culture) + ' in'
        }
        $page.PageSheet.CellsU('LineJumpFactorX').FormulaU = ([double]$factorX).ToString('R',$culture)
        $page.PageSheet.CellsU('LineJumpFactorY').FormulaU = ([double]$factorY).ToString('R',$culture)
        foreach ($shape in @($h,$v)) { $shape.CellsU('LineWeight').FormulaU = "$weight pt" }
        $page.PageSheet.CellsU('LineJumpCode').FormulaU = $(if($direction -eq 'h'){'1'}else{'2'})
        $page.Export($temporary)
        $svg = Get-Content -LiteralPath $temporary -Raw
        $id = $(if($direction -eq 'h'){$h.ID}else{$v.ID})
        $path = [regex]::Match($svg, ('(?s)<g id="shape'+$id+'-[^"]*".*?<path d="([^"]*)"')).Groups[1].Value
        if (!$path) { throw 'The native SVG did not contain the connector path.' }
        $arc = [regex]::Match($path, '(?i)A\s*([\d.eE+-]+)\s+([\d.eE+-]+)')
        $rx = $(if($arc.Success){[double]::Parse($arc.Groups[1].Value,$culture)}else{$null})
        $ry = $(if($arc.Success){[double]::Parse($arc.Groups[2].Value,$culture)}else{$null})
        $factor = $(if($direction -eq 'h'){$factorX}else{$factorY})
        $spacing = $(if($direction -eq 'h'){$spacingX}else{$spacingY})
        $expected = 36 * $factor * $spacing * $pageScale / $drawingScale
        if ([Math]::Abs([double]$rx - $expected) -gt 0.00001 -or [Math]::Abs([double]$ry - $expected) -gt 0.00001) {
            throw "Unexpected native jump radius: expected $expected pt, got $rx,$ry."
        }
        $cases += ,@($direction,$factorX,$factorY,$spacingX,$spacingY,$pageScale,$drawingScale,$weight,$rx,$ry,$path)
    }
    [ordered]@{
        source = "Microsoft Visio $($app.Version) InvisibleApp COM SVG export, $(Get-Date -Format yyyy-MM-dd)"
        setup = 'Two crossing straight dynamic connectors; page arc style=1; connector style/policy=0'
        units = 'Spacing and page/drawing scale are raw ResultIU inches; arc radii are paper points'
        rule = 'Radius inches = axis factor * axis spacing * PageScale / DrawingScale / 2; zero removes the jump'
        columns = @('direction','factorX','factorY','lineToLineX','lineToLineY','pageScale','drawingScale','lineWeightPt','radiusXPt','radiusYPt','nativePath')
        cases = $cases
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Output
    Write-Output "Recorded $($cases.Count) native cases; all radii match the measured rule."
} finally {
    if ($doc) { $doc.Saved=$true; $doc.Close() }
    $app.Quit()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
    Remove-Item -LiteralPath $temporary -ErrorAction SilentlyContinue
}
