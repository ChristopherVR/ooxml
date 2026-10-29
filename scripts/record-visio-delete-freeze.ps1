# Native check of deleting a shape that another shape's formulas read. The core freezes the reading
# cells to their last result (Visio's own behaviour, see scripts/record-visio-delete-error.ps1).
# Opens before.vsdx (follower still formula-driven) and after.vsdx (rectangle deleted by the core) in an
# owned invisible Visio with AlertResponse = 7, and compares the follower's PinX/PinY with the core caches.
# First run: bun scripts/create-visio-delete-freeze-fixture.ts <OutputDirectory>
# Then:      pwsh scripts/record-visio-delete-freeze.ps1 -OutputDirectory <OutputDirectory> [-EvidencePath <json>]
param(
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [string]$EvidencePath
)
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$expected = Get-Content -LiteralPath (Join-Path $directory 'expected.json') -Raw | ConvertFrom-Json
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
function Read-Follower($doc, $id) {
    $shape = $doc.Pages.Item(1).Shapes.ItemFromID([int]$id)
    [ordered]@{
        pinXFormula = $shape.CellsU('PinX').FormulaU; pinX = $shape.CellsU('PinX').ResultIU
        pinYFormula = $shape.CellsU('PinY').FormulaU; pinY = $shape.CellsU('PinY').ResultIU
        shapeCount = $doc.Pages.Item(1).Shapes.Count
    }
}
try {
    $application.AlertResponse = 7
    $document = $application.Documents.OpenEx((Join-Path $directory 'before.vsdx'), 202)
    $before = Read-Follower $document $expected.shape
    $document.Close(); $document = $null
    $document = $application.Documents.OpenEx((Join-Path $directory 'after.vsdx'), 202)
    $after = Read-Follower $document $expected.shape
    $deltaX = [Math]::Abs($after.pinX - [double]$expected.pinX)
    $deltaY = [Math]::Abs($after.pinY - [double]$expected.pinY)
    $result = [ordered]@{
        product = "Visio $($application.Version) build $($application.Build)"
        before = $before; after = $after
        coreCache = [ordered]@{ pinX = $expected.pinX; pinY = $expected.pinY }
        maxDelta = [Math]::Max($deltaX, $deltaY)
        status = $(if ([Math]::Max($deltaX, $deltaY) -le 1e-6 -and $after.shapeCount -eq ($before.shapeCount - 1)) { 'pass' } else { 'fail' })
    }
    $json = $result | ConvertTo-Json -Depth 6
    if ($EvidencePath) { Set-Content -LiteralPath $EvidencePath -Value $json -Encoding utf8 }
    $json
}
finally {
    if ($document) { $document.Close() }
    $application.Quit()
}
