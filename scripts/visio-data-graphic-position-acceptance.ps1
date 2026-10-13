param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
& bun (Join-Path $PSScriptRoot 'visio-data-graphic-position-acceptance.ts') $OutputDirectory
if($LASTEXITCODE){throw 'Core generation failed.'}
$OutputDirectory=(Resolve-Path -LiteralPath $OutputDirectory).Path
$expected=Get-Content -LiteralPath (Join-Path $OutputDirectory 'expected.json') -Raw|ConvertFrom-Json
$app=New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse=7
try {
    $results=@()
    foreach($name in @('authored','core-moved')) {
        $doc=$app.Documents.OpenEx((Join-Path $OutputDirectory "$name.vsdx"),66)
        try {
            if($name -eq 'authored') {
                $owner=$doc.Pages.Item(1).Shapes.ItemFromID(1)
                $owner.CellsU('PinX').ResultIU=4
                $owner.CellsU('PinY').ResultIU=5
            }
            foreach($entry in $expected) {
                $shape=$doc.Pages.Item(1).Shapes.ItemFromID($entry.id)
                $x=$shape.CellsU('PinX').ResultIU
                $y=$shape.CellsU('PinY').ResultIU
                if([Math]::Abs($x-$entry.x-2) -gt 0.0000001 -or [Math]::Abs($y-$entry.y-2) -gt 0.0000001){throw 'Native position differs.'}
                $results += [ordered]@{source=$name;shape=$entry.id;x=$x;y=$y;formula=$shape.CellsU('PinX').FormulaU}
            }
        } finally {$doc.Saved=$true;$doc.Close()}
    }
    [ordered]@{version=$app.Version;cases=$results}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json')
    $results|ConvertTo-Json
} finally {$app.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)|Out-Null}
