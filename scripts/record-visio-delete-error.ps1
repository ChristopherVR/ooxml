# Probes Visio 16 through an owned invisible instance for two ShapeSheet behaviours the core edit
# guards depend on, and prints them as JSON:
#  1. deleting a shape that another shape's formula references (Sheet.2!Width): the formula text and
#     result Visio leaves behind;
#  2. a cell whose formula errors (1/0): its ResultIU, ResultStr and whether dependents follow.
# Run: pwsh scripts/record-visio-delete-error.ps1 [-EvidencePath <json>]
param([string]$EvidencePath)
$ErrorActionPreference = 'Stop'
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
try {
    $application.AlertResponse = 7
    $document = $application.Documents.Add('')
    $page = $document.Pages.Item(1)
    $one = $page.DrawRectangle(1, 1, 3, 2)
    $two = $page.DrawRectangle(4, 1, 6, 3)
    $one.CellsU('LocPinX').FormulaU = "Sheet.$($two.ID)!Width*0.25"
    $before = [ordered]@{ formula = $one.CellsU('LocPinX').FormulaU; result = $one.CellsU('LocPinX').ResultIU }
    $two.Delete()
    $afterCell = $one.CellsU('LocPinX')
    $after = [ordered]@{ formula = $afterCell.FormulaU; result = $afterCell.ResultIU; resultStr = $afterCell.ResultStr('') }
    $three = $page.DrawRectangle(1, 4, 3, 5)
    $three.CellsU('PinX').FormulaU = '1/0'
    $four = $page.DrawRectangle(5, 4, 7, 5)
    $four.CellsU('Width').FormulaU = 'Sheet.' + $three.ID + '!PinX+1'
    $errorProbe = [ordered]@{
        formula = $three.CellsU('PinX').FormulaU
        result = $three.CellsU('PinX').ResultIU
        resultStr = $three.CellsU('PinX').ResultStr('')
        dependentFormula = $four.CellsU('Width').FormulaU
        dependentResult = $four.CellsU('Width').ResultIU
        dependentStr = $four.CellsU('Width').ResultStr('')
    }
    $result = [ordered]@{ product = "Visio $($application.Version) build $($application.Build)"; deleteReferenced = [ordered]@{ before = $before; after = $after }; errorCell = $errorProbe }
    $json = $result | ConvertTo-Json -Depth 6
    if ($EvidencePath) { Set-Content -LiteralPath $EvidencePath -Value $json -Encoding utf8 }
    $json
}
finally {
    if ($document) { $document.Close() }
    $application.Quit()
}
