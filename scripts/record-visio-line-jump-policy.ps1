param([Parameter(Mandatory=$true)][string]$Output, [string]$FixtureOutput)
# Manual Windows acceptance probe. Owns an invisible Visio instance and creates no user documents.
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Visio.InvisibleApp
$svgPath = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString() + '.svg')
try {
    $doc = $app.Documents.Add('')
    $page = $doc.Pages.Item(1)
    $page.PageSheet.CellsU('LineJumpStyle').FormulaU = '1'
    $h = $page.Drop($app.ConnectorToolDataObject, 1, 1)
    $v = $page.Drop($app.ConnectorToolDataObject, 1, 1)
    foreach ($shape in @($h, $v)) { $shape.CellsU('ShapeRouteStyle').FormulaU = '16' }
    foreach ($entry in @(@('BeginX',1), @('BeginY',2), @('EndX',5), @('EndY',2))) {
        $h.CellsU($entry[0]).FormulaU = "$($entry[1]) in"
    }
    foreach ($entry in @(@('BeginX',3), @('BeginY',0), @('EndX',3), @('EndY',4))) {
        $v.CellsU($entry[0]).FormulaU = "$($entry[1]) in"
    }
    $rows = @()
    foreach ($pc in @(0,1,2,4,5)) {
        $page.PageSheet.CellsU('LineJumpCode').FormulaU = [string]$pc
        foreach ($hc in 0..4) {
            $h.CellsU('ConLineJumpCode').FormulaU = [string]$hc
            foreach ($vc in 0..4) {
                $v.CellsU('ConLineJumpCode').FormulaU = [string]$vc
                $page.Export($svgPath)
                $svg = Get-Content -LiteralPath $svgPath -Raw
                $hp = [regex]::Match($svg, '(?s)<g id="shape1-[^"]*".*?<path d="([^"]*)"').Groups[1].Value
                $vp = [regex]::Match($svg, '(?s)<g id="shape2-[^"]*".*?<path d="([^"]*)"').Groups[1].Value
                if (!$hp -or !$vp) { throw 'Native SVG did not contain both connector paths.' }
                $rows += ,@($pc, $hc, $vc, ($hp -match 'A'), ($vp -match 'A'))
            }
        }
    }
    [ordered]@{
        source = "Microsoft Visio $($app.Version) COM SVG export, $(Get-Date -Format yyyy-MM-dd)"
        setup = 'Two crossing straight dynamic connectors, horizontal drawn first; LineJumpStyle=1'
        columns = @('page','horizontal','vertical','hJump','vJump')
        cases = $rows
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $Output
    if ($FixtureOutput) {
        $page.PageSheet.CellsU('LineJumpCode').FormulaU = '1'
        $h.CellsU('ConLineJumpCode').FormulaU = '0'
        $v.CellsU('ConLineJumpCode').FormulaU = '0'
        $doc.SaveAs([IO.Path]::GetFullPath($FixtureOutput))
    }
    $doc.Saved = $true
    $doc.Close()
} finally {
    $app.Quit()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
    Remove-Item -LiteralPath $svgPath -ErrorAction SilentlyContinue
}
