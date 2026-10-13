param([Parameter(Mandatory=$true)][string]$Output)
# Manual native probe: owns its invisible application and never opens user documents.
$ErrorActionPreference = 'Stop'
$culture = [Globalization.CultureInfo]::InvariantCulture
$app = New-Object -ComObject Visio.InvisibleApp
$temporary = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N') + '.svg')
$doc = $null
try {
    $doc = $app.Documents.Add('')
    $page = $doc.Pages.Item(1)
    $a = $page.Drop($app.ConnectorToolDataObject, 1, 1)
    $b = $page.Drop($app.ConnectorToolDataObject, 1, 1)
    $page.PageSheet.CellsU('LineJumpStyle').FormulaU = '1'
    $page.PageSheet.CellsU('LineJumpFactorX').FormulaU = '0.5'
    $page.PageSheet.CellsU('LineJumpFactorY').FormulaU = '2'
    $cases = @()
    foreach ($curved in @($false,$true)) {
        foreach ($angle in @(0,15,30,45,60,75,90,120,135,150,165)) {
          foreach ($peerAngle in @(($angle+90), (-$angle)) | Select-Object -Unique) {
            if ([Math]::Abs([Math]::Sin(($peerAngle-$angle)*[Math]::PI/180)) -lt 0.01) { continue }
            $rad = $angle * [Math]::PI / 180
            $dx = 2 * [Math]::Cos($rad); $dy = 2 * [Math]::Sin($rad)
            $peerRad = $peerAngle * [Math]::PI / 180
            $px = 2 * [Math]::Cos($peerRad); $py = 2 * [Math]::Sin($peerRad)
            $ends = @(@((3-$dx),(3-$dy),(3+$dx),(3+$dy)),@((3-$px),(3-$py),(3+$px),(3+$py)))
            $i = 0
            foreach ($shape in @($a,$b)) {
                $shape.CellsU('ShapeRouteStyle').FormulaU = '16'
                $shape.CellsU('ConLineRouteExt').FormulaU = $(if($curved){'2'}else{'1'})
                $j = 0
                foreach ($name in @('BeginX','BeginY','EndX','EndY')) {
                    $shape.CellsU($name).FormulaU = ([double]$ends[$i][$j]).ToString('R',$culture) + ' in'
                    $j++
                }
                $i++
            }
            foreach ($code in @(1,2,4,5)) {
                $page.PageSheet.CellsU('LineJumpCode').FormulaU = [string]$code
                $page.Export($temporary)
                $svg = Get-Content -LiteralPath $temporary -Raw
                $paths = @()
                foreach ($shape in @($a,$b)) {
                    $path = [regex]::Match($svg, '(?s)<g id="shape' + $shape.ID + '-[^"]*".*?<path d="([^"]*)"').Groups[1].Value
                    if (!$path) { throw 'Missing native connector path.' }
                    $paths += $path
                }
                $cases += [ordered]@{curved=$curved;angle=$angle;peerAngle=$peerAngle;ends=$ends;pageCode=$code;paths=$paths}
            }
          }
        }
    }
    [ordered]@{source="Microsoft Visio $($app.Version) COM SVG export, $(Get-Date -Format yyyy-MM-dd)";factorX=0.5;factorY=2;cases=$cases} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Output
    "Recorded $($cases.Count) cases."
} finally {
    if ($doc) { $doc.Saved=$true; $doc.Close() }
    $app.Quit()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
    Remove-Item -LiteralPath $temporary -ErrorAction SilentlyContinue
}
