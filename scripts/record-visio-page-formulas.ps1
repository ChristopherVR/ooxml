# Numeric page-function oracle in an owned, invisible Microsoft Visio instance.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP 'visio-page-formulas'),
    [string]$CoreOutputPath,
    [string]$CoreInsertedOutputPath
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$app = New-Object -ComObject Visio.InvisibleApp
$document = $null
function Read-PageCaches($doc) {
    foreach ($page in $doc.Pages) {
        $shape = $page.Shapes.Item(1)
        [ordered]@{
            id = [string]$page.ID; index = $page.Index; background = [bool]$page.Background
            number = $shape.CellsU('User.Number').ResultIU
            count = $shape.CellsU('User.Count').ResultIU
            combined = $shape.CellsU('User.Combined').ResultIU
        }
    }
}
try {
    $app.AlertResponse = 7
    $document = $app.Documents.Add('')
    $background = $document.Pages.Add()
    $background.NameU = 'Background'; $background.Background = -1
    $last = $document.Pages.Add()
    foreach ($page in @($document.Pages.Item(1), $background, $last)) {
        $shape = $page.DrawRectangle(1,1,3,2)
        foreach ($name in @('Number','Count','Combined')) {
            $shape.AddNamedRow(242,$name,0) | Out-Null
        }
        $shape.CellsU('User.Number').FormulaU = 'PAGENUMBER()'
        $shape.CellsU('User.Count').FormulaU = 'PAGECOUNT()'
        $shape.CellsU('User.Combined').FormulaU = 'User.Number+User.Count'
    }
    $initial = @(Read-PageCaches $document)
    $document.SaveAs((Join-Path $OutputDirectory 'page-formulas.vsdx')) | Out-Null
    $last.Index = 1
    $document.SaveAs((Join-Path $OutputDirectory 'reordered-native.vsdx')) | Out-Null
    $afterReorder = @(Read-PageCaches $document)
    # Reassigning the identical formula leaves a stale cache in Visio 16.0.
    # Toggle to a literal first to obtain the freshly evaluated native reference.
    foreach ($page in $document.Pages) {
        $shape = $page.Shapes.Item(1)
        $shape.CellsU('User.Number').FormulaU = '0'
        $shape.CellsU('User.Number').FormulaU = 'PAGENUMBER()'
    }
    $recalculated = @(Read-PageCaches $document)
    $document.SaveAs((Join-Path $OutputDirectory 'recalculated-native.vsdx')) | Out-Null
    $document.Saved = $true; $document.Close(); $document = $null
    $accepted = @()
    if ($CoreOutputPath) {
        $document = $app.Documents.OpenEx((Resolve-Path -LiteralPath $CoreOutputPath).Path,128)
        $accepted = @(Read-PageCaches $document)
        foreach ($actual in $accepted) {
            $expected = $recalculated | Where-Object { $_.id -eq $actual.id }
            foreach ($key in @('index','background','number','count','combined')) {
                if ($actual[$key] -ne $expected[$key]) { throw "Core page $($actual.id): incorrect $key" }
            }
        }
        $document.SaveAs((Join-Path $OutputDirectory 'core-native-reopened.vsdx')) | Out-Null
        $document.Saved = $true; $document.Close(); $document = $null
    }
    $insertedAccepted = @()
    if ($CoreInsertedOutputPath) {
        $document = $app.Documents.OpenEx((Resolve-Path -LiteralPath $CoreInsertedOutputPath).Path,128)
        if ($document.Pages.Count -ne 4) { throw 'Inserted drawing must have four pages.' }
        foreach ($page in $document.Pages) {
            if ($page.Shapes.Count -eq 0) { continue }
            $shape = $page.Shapes.Item(1)
            $expectedNumber = if ($page.Background) { 0 } else { $page.Index }
            if ($shape.CellsU('User.Number').ResultIU -ne $expectedNumber -or
                $shape.CellsU('User.Count').ResultIU -ne 3 -or
                $shape.CellsU('User.Combined').ResultIU -ne ($expectedNumber + 3)) {
                throw "Incorrect insertion caches on page $($page.ID)"
            }
            $insertedAccepted += [ordered]@{
                id = [string]$page.ID; number = $shape.CellsU('User.Number').ResultIU
                count = $shape.CellsU('User.Count').ResultIU
                combined = $shape.CellsU('User.Combined').ResultIU
            }
        }
        $document.SaveAs((Join-Path $OutputDirectory 'core-inserted-native-reopened.vsdx')) | Out-Null
    }
    [ordered]@{
        application = 'Microsoft Visio'; version = $app.Version
        initial = $initial; afterReorder = $afterReorder
        recalculated = $recalculated; accepted = $accepted; insertedAccepted = $insertedAccepted
    } | ConvertTo-Json -Depth 5 |
        Set-Content (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
} finally {
    if ($document) { $document.Saved = $true; $document.Close() }
    $app.Quit()
}
Write-Output $OutputDirectory
