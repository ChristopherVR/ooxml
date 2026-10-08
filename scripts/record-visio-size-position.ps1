# Owned native cell-edit oracle for numeric pin, size and angle controls.
# First run: bun scripts/create-visio-size-position-fixtures.ts <OutputDirectory>
param([Parameter(Mandatory=$true)][string]$OutputDirectory, [switch]$ReopenCore)
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$cases = @(Get-Content -LiteralPath (Join-Path $directory 'cases.json') -Raw | ConvertFrom-Json)
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
$cells = @{ x='PinX'; y='PinY'; width='Width'; height='Height'; angle='Angle' }
function Read-Shape($shape) {
    $values = [ordered]@{ text=$shape.Text }
    foreach ($name in @('PinX','PinY','Width','Height','LocPinX','LocPinY','Angle')) {
        $values[$name] = [double]$shape.CellsU($name).ResultIU
    }
    $values
}
try {
    $application.AlertResponse = 7
    $observations = @()
    $reference = if ($ReopenCore) { Get-Content -LiteralPath (Join-Path $directory 'evidence.json') -Raw | ConvertFrom-Json } else { $null }
    foreach ($case in $cases) {
        $suffix = if ($ReopenCore) { '-core.vsdx' } else { '-original.vsdx' }
        $document = $application.Documents.OpenEx((Join-Path $directory ($case.name + $suffix)),202)
        $shape = $document.Pages.Item(1).Shapes.ItemFromID(1)
        if (!$ReopenCore) {
            $unit = if ($case.field -eq 'angle') { ' deg' } else { ' in' }
            $shape.CellsU($cells[$case.field]).FormulaU = ([double]$case.value).ToString('R',[Globalization.CultureInfo]::InvariantCulture) + $unit
            [void]$document.SaveAs((Join-Path $directory ($case.name + '-native.vsdx')))
        }
        $observed = Read-Shape $shape
        if ($ReopenCore) {
            $expected = @($reference.cases | Where-Object { $_.name -eq $case.name })[0].shape
            if ($observed.text -cne $expected.text) { throw "Text mismatch: $($case.name)" }
            foreach ($key in @('PinX','PinY','Width','Height','LocPinX','LocPinY','Angle')) {
                if ([Math]::Abs($observed[$key] - $expected.$key) -gt 1e-10) { throw "Cell mismatch: $($case.name)/$key" }
            }
        }
        $observations += [ordered]@{ name=$case.name; shape=$observed }
        $document.Saved = $true
        $document.Close()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
        $document = $null
    }
    $record = if ($ReopenCore) { 'reopen-evidence.json' } else { 'evidence.json' }
    [ordered]@{ application='Microsoft Visio'; version=$application.Version; cases=$observations } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory $record) -Encoding utf8
    Write-Output $directory
} finally {
    if ($document) {
        $document.Saved = $true
        $document.Close()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    $application.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
