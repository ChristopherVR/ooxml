# Local native conformance probe. Requires desktop Microsoft Visio on Windows.
# Creates its own invisible instance and documents; never attaches to the user's session.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('visio-rounding-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$cases = @(
    @{ Name = 'shared'; Points = @(0,0,2,0,2,0.2,4,0.2); Radii = @(0.1,0.1) },
    @{ Name = 'endpoint'; Points = @(0,0,0.1,0,0.1,2); Radii = @(0.05) },
    @{ Name = 'endpoint-reservation'; Points = @(0,0,0.4,0,0.4,2); Radii = @(0.2) },
    @{ Name = 'unequal'; Points = @(0,0,0.05,0,0.05,0.2,2,0.2); Radii = @(0.025,0.175) },
    @{ Name = 'reversed'; Points = @(2,0.2,0.05,0.2,0.05,0,0,0); Radii = @(0.1,0.025) },
    @{ Name = 'both-endpoints'; Points = @(0,0,0.05,0,0.05,0.2,0.1,0.2); Radii = @(0.025,0.025) },
    @{ Name = 'three-corners'; Points = @(0,0,2,0,2,0.2,2.3,0.2,2.3,2); Radii = @(0.1,0.1,0.2) },
    @{ Name = 'collinear'; Points = @(0,0,1.95,0,2,0,2,1,4,1); Radii = @(0.05,0.25) },
    @{ Name = 'local'; Points = @(0,0,2,0,2,0.2,4,0.2,4,2,6,2); Radii = @(0.1,0.1,0.25,0.25) }
)
$native = New-Object -ComObject Visio.InvisibleApp
try {
    $document = $native.Documents.Add('')
    $page = $document.Pages.Item(1)
    foreach ($case in $cases) {
        # Explicit rows exercise the admitted MoveTo/LineTo subset. DrawPolyline
        # normally saves a compressed PolylineTo row, which is outside this slice.
        $shape = $page.DrawRectangle(0, 0, 6, 2)
        $shape.DeleteSection(10) # visSectionFirstComponent
        $shape.AddSection(10) | Out-Null
        $shape.CellsU('Geometry1.NoFill').FormulaU = '1'
        for ($vertex = 0; $vertex -lt $case.Points.Count / 2; $vertex++) {
            $tag = if ($vertex -eq 0) { 138 } else { 139 } # MoveTo / LineTo
            $row = $shape.AddRow(10, $vertex + 1, $tag)
            $shape.CellsSRC(10, $row, 0).FormulaU = ([double]$case.Points[2 * $vertex]).ToString([cultureinfo]::InvariantCulture) + ' in'
            $shape.CellsSRC(10, $row, 1).FormulaU = ([double]$case.Points[2 * $vertex + 1]).ToString([cultureinfo]::InvariantCulture) + ' in'
        }
        $shape.CellsU('Rounding').FormulaU = '0.25 in'
        $shape.CellsU('FillPattern').FormulaU = '0'
        $shape.CellsU('BeginArrow').FormulaU = '0'
        $shape.CellsU('EndArrow').FormulaU = '0'
        $file = Join-Path $OutputDirectory ($case.Name + '.svg')
        $shape.Export($file)
        [xml]$svg = Get-Content -LiteralPath $file -Raw
        # Visio may emit a separate implicitly closed fill path even with fill disabled.
        $paths = @($svg.SelectNodes('//*[local-name()="path"]') | Where-Object {
            $_.GetAttribute('d') -notmatch 'Z\s*$'
        })
        if ($paths.Count -ne 1) { throw ($case.Name + ': expected one native path') }
        $path = $paths[0].GetAttribute('d')
        $arcs = [regex]::Matches($path, 'A([\d.]+)\s+([\d.]+)')
        if ($arcs.Count -ne $case.Radii.Count) { throw ($case.Name + ': unexpected arc count') }
        for ($i = 0; $i -lt $arcs.Count; $i++) {
            foreach ($axis in @(1,2)) {
                $radius = [double]::Parse($arcs[$i].Groups[$axis].Value, [cultureinfo]::InvariantCulture) / 72
                if ([math]::Abs($radius - $case.Radii[$i]) -gt 0.000001) {
                    throw ($case.Name + ': native radius mismatch at corner ' + $i + ': ' + $radius)
                }
            }
        }
        $case.ShapeId = [string]$shape.ID
        $case.NativePath = $path
        Write-Output ($case.Name + ': native radii verified')
    }
    $document.SaveAs((Join-Path $OutputDirectory 'rounding.vsdx')) | Out-Null
    @{ Version = $native.Version; Cases = $cases } | ConvertTo-Json -Depth 5 |
        Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
    Write-Output ('Visio ' + $native.Version + '; evidence: ' + $OutputDirectory)
} finally {
    if ($null -ne $document) { $document.Saved = $true; $document.Close() }
    $native.Quit()
}
