param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory=(Resolve-Path -LiteralPath $OutputDirectory).Path
$app=New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse=7
try {
    $doc=$app.Documents.Add('')
    try {
        $shape=$doc.Pages.Item(1).DrawRectangle(1,1,5,3)
        $shape.AddNamedRow(243,'Offset',0)|Out-Null
        $shape.CellsU('Prop.Offset.Type').FormulaU='2'
        $shape.CellsU('Prop.Offset').FormulaU='Width/1IN*0.25'
        $dependent=$doc.Pages.Item(1).DrawRectangle(0.5,4,1.5,5)
        $dependent.CellsU('PinX').FormulaU="Sheet.$($shape.ID)!Prop.Offset*1IN"
        $doc.SaveAs((Join-Path $OutputDirectory 'native.vsdx'))|Out-Null
    } finally {$doc.Saved=$true;$doc.Close()}
    & bun (Join-Path $PSScriptRoot 'visio-property-dependencies.ts') $OutputDirectory
    if($LASTEXITCODE){throw 'Core edit failed.'}
    $doc=$app.Documents.OpenEx((Join-Path $OutputDirectory 'edited.vsdx'),66)
    try {
        $shape=$doc.Pages.Item(1).Shapes.ItemFromID(1)
        $dependent=$doc.Pages.Item(1).Shapes.ItemFromID(2)
        $result=[ordered]@{version=$app.Version;width=$shape.CellsU('Width').ResultIU;property=$shape.CellsU('Prop.Offset').ResultIU;dependentPinX=$dependent.CellsU('PinX').ResultIU;formula=$dependent.CellsU('PinX').FormulaU}
        if($result.width -ne 8 -or $result.property -ne 2 -or $result.dependentPinX -ne 2){throw 'Native dependency values differ.'}
        $result|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json')
        $result|ConvertTo-Json
    } finally {$doc.Saved=$true;$doc.Close()}
} finally {$app.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)|Out-Null}
