# Native open, ShapeSheet read-back and save-as-copy check of core-edited VSDX files, using only an
# owned invisible Visio instance. Compares every cached cell value the core wrote (PinX, PinY, Width,
# Height, Angle, LocPinX, LocPinY, LineWeight) with the ResultIU Visio reports after opening, then
# saves a copy from Visio, reopens it and checks the values are unchanged.
# First run: bun scripts/create-visio-edit-family-fixtures.ts <OutputDirectory> [reference-flat.vsdx|reference-corpus.vsdx]
# Then:      pwsh scripts/record-visio-edit-families.ps1 -OutputDirectory <OutputDirectory> [-EvidencePath <json>]
param(
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [string]$EvidencePath
)
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$cases = @(Get-Content -LiteralPath (Join-Path $directory 'cases.json') -Raw | ConvertFrom-Json)
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
function Index-Shapes($shapes, $index) {
    foreach ($shape in $shapes) {
        $index[[string]$shape.ID] = $shape
        if ($shape.Type -eq 2) { Index-Shapes $shape.Shapes $index }
    }
}
function Read-Cells($doc, $pageName, $claimed) {
    $page = $doc.Pages.ItemU($pageName)
    $index = @{}
    Index-Shapes $page.Shapes $index
    $observed = @{}
    foreach ($shape in $claimed) {
        $live = $index[[string]$shape.id]
        if ($null -eq $live) { $observed[[string]$shape.id] = $null; continue }
        $values = @{}
        foreach ($name in $shape.cells.PSObject.Properties.Name) { $values[$name] = [double]$live.CellsU($name).ResultIU }
        foreach ($name in $shape.colors.PSObject.Properties.Name) { $values['color:' + $name] = ($live.CellsU($name).ResultStr('') -replace ' ', '') }
        $observed[[string]$shape.id] = $values
    }
    [ordered]@{ observed = $observed; total = $index.Count }
}
try {
    $application.AlertResponse = 7
    $results = @()
    foreach ($case in $cases) {
        $record = [ordered]@{ name = $case.name; family = $case.family; status = $case.status }
        if ($case.status -ne 'written') {
            $record.result = 'core-refused'; $record.detail = $case.error
            $results += $record; continue
        }
        $path = Join-Path $directory ($case.name + '.vsdx')
        $copy = Join-Path $directory ($case.name + '-visio.vsdx')
        $compared = 0; $maxDelta = 0.0; $mismatches = @(); $roundTripDrift = 0.0; $recalcCells = 0; $recalcRefused = 0; $recalcMaxDelta = 0.0; $recalcMismatches = @(); $opened = $false; $reopened = $false
        try {
            $document = $application.Documents.OpenEx($path, 202)
            $opened = $true
            $snapshots = @{}
            foreach ($page in $case.pages) {
                $read = Read-Cells $document $page.page $page.shapes
                $snapshots[$page.page] = $read.observed
                foreach ($shape in $page.shapes) {
                    $live = $read.observed[[string]$shape.id]
                    if ($null -eq $live) { $mismatches += "$($page.page)/$($shape.id): shape missing in Visio"; continue }
                    foreach ($name in $shape.colors.PSObject.Properties.Name) {
                        $compared++
                        if ($live['color:' + $name] -ine $shape.colors.$name) { $mismatches += "$($page.page)/$($shape.id)/$name core=$($shape.colors.$name) visio=$($live['color:' + $name])" }
                    }
                    foreach ($name in $shape.cells.PSObject.Properties.Name) {
                        $delta = [Math]::Abs($live[$name] - [double]$shape.cells.$name)
                        $compared++
                        if ($delta -gt $maxDelta) { $maxDelta = $delta }
                        if ($delta -gt 1e-6) { $mismatches += "$($page.page)/$($shape.id)/$name core=$($shape.cells.$name) visio=$($live[$name])" }
                    }
                }
            }
            [void]$document.SaveAsEx($copy, 0)
            # Force Visio to evaluate every formula it holds: re-enter each numeric cell's own formula.
            foreach ($page in $case.pages) {
                $index = @{}
                Index-Shapes $document.Pages.ItemU($page.page).Shapes $index
                foreach ($shape in $page.shapes) {
                    foreach ($name in $shape.cells.PSObject.Properties.Name) {
                        try { $cell = $index[[string]$shape.id].CellsU($name); $formula = $cell.FormulaU; $cell.FormulaU = $formula; $recalcCells++ }
                        catch { $recalcRefused++ }
                    }
                }
                $read = Read-Cells $document $page.page $page.shapes
                foreach ($shape in $page.shapes) {
                    foreach ($name in $shape.cells.PSObject.Properties.Name) {
                        $delta = [Math]::Abs($read.observed[[string]$shape.id][$name] - [double]$shape.cells.$name)
                        if ($delta -gt $recalcMaxDelta) { $recalcMaxDelta = $delta }
                        if ($delta -gt 1e-6) { $recalcMismatches += "$($page.page)/$($shape.id)/$name core=$($shape.cells.$name) visio-after-recalc=$($read.observed[[string]$shape.id][$name])" }
                    }
                }
            }
            $document.Saved = $true; $document.Close(); $document = $null
            $document = $application.Documents.OpenEx($copy, 202)
            $reopened = $true
            foreach ($page in $case.pages) {
                $read = Read-Cells $document $page.page $page.shapes
                foreach ($shape in $page.shapes) {
                    $before = $snapshots[$page.page][[string]$shape.id]; $after = $read.observed[[string]$shape.id]
                    if ($null -eq $after) { $mismatches += "reopen $($page.page)/$($shape.id): shape missing"; continue }
                    foreach ($name in @($before.Keys | Where-Object { $_ -notlike 'color:*' })) {
                        $drift = [Math]::Abs($after[$name] - $before[$name])
                        if ($drift -gt $roundTripDrift) { $roundTripDrift = $drift }
                        if ($drift -gt 1e-6) { $mismatches += "reopen $($page.page)/$($shape.id)/$name before=$($before[$name]) after=$($after[$name])" }
                    }
                }
            }
            $record.result = if ($mismatches.Count -eq 0 -and $recalcMismatches.Count -eq 0) { 'pass' } else { 'fail' }
        } catch {
            $record.result = 'fail'; $record.detail = $_.Exception.Message
        } finally {
            if ($null -ne $document) { try { $document.Saved = $true; $document.Close() } catch { }; $document = $null }
        }
        $record.openedWithoutRepair = $opened; $record.savedCopyReopened = $reopened
        $record.cellsCompared = $compared; $record.maxDelta = $maxDelta; $record.roundTripDrift = $roundTripDrift
        $record.recalcCells = $recalcCells; $record.recalcRefused = $recalcRefused; $record.recalcMaxDelta = $recalcMaxDelta
        $record.recalcMismatches = @($recalcMismatches | Select-Object -First 12)
        $record.mismatches = @($mismatches | Select-Object -First 12)
        $results += $record
    }
    $evidence = [ordered]@{ application = 'Microsoft Visio'; version = $application.Version; build = $application.Build
        recorded = (Get-Date -Format 'yyyy-MM-dd'); cases = $results }
    $json = $evidence | ConvertTo-Json -Depth 8
    if ($EvidencePath) { $json | Set-Content -LiteralPath $EvidencePath -Encoding utf8 }
    $json | Set-Content -LiteralPath (Join-Path $directory 'evidence.json') -Encoding utf8
    $results | ForEach-Object { '{0,-32} {1,-12} cells={2} maxDelta={3} drift={4} recalc={5} refused={6} recalcDelta={7}' -f $_.name, $_.result, $_.cellsCompared, $_.maxDelta, $_.roundTripDrift, $_.recalcCells, $_.recalcRefused, $_.recalcMaxDelta }
} finally {
    if ($null -ne $document) { try { $document.Saved = $true; $document.Close() } catch { } }
    $application.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
