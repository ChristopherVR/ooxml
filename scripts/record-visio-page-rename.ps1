# Local/universal page-name and reference oracle in a separate invisible Visio instance.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP 'visio-page-rename'),
    [string]$CoreOutputPath,
    [string]$CoreSecondOutputPath
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$app = New-Object -ComObject Visio.InvisibleApp
$document = $null
function Read-Rename($doc) {
    $page = $doc.Pages.Item(1)
    $shape = $doc.Pages.Item(2).Shapes.Item(1)
    $links = @($shape.Hyperlinks | ForEach-Object { $_.SubAddress })
    [ordered]@{
        id = [string]$page.ID; name = $page.Name; nameU = $page.NameU
        cachedName = $page.Shapes.Item(1).CellsU('User.Name').ResultStr('')
        formula = $shape.CellsU('User.PageWidth').FormulaU; subAddress = $links[0]
    }
}
try {
    $app.AlertResponse = 7
    $document = $app.Documents.Add('')
    $page = $document.Pages.Item(1)
    $named = $page.DrawRectangle(1,1,3,2)
    $named.AddNamedRow(242,'Name',0) | Out-Null
    $named.CellsU('User.Name').FormulaU = 'PAGENAME()'
    $other = $document.Pages.Add()
    $shape = $other.DrawRectangle(1,1,3,2)
    $shape.AddNamedRow(242,'PageWidth',0) | Out-Null
    $shape.CellsU('User.PageWidth').FormulaU = 'Pages[Page-1]!ThePage!PageWidth'
    $link = $shape.AddHyperlink(); $link.SubAddress = 'Page-1'
    $document.SaveAs((Join-Path $OutputDirectory 'original.vsdx')) | Out-Null
    $page.Name = 'Renamed & Page'
    $first = Read-Rename $document
    $document.SaveAs((Join-Path $OutputDirectory 'renamed-native.vsdx')) | Out-Null
    $page.Name = 'Second rename'
    $second = Read-Rename $document
    $document.SaveAs((Join-Path $OutputDirectory 'renamed-again-native.vsdx')) | Out-Null
    $document.Saved = $true; $document.Close(); $document = $null
    $accepted = @()
    foreach ($case in @(@($CoreOutputPath,$first,'core-native-reopened.vsdx'),
                        @($CoreSecondOutputPath,$second,'core-second-native-reopened.vsdx'))) {
        if (-not $case[0]) { continue }
        $document = $app.Documents.OpenEx((Resolve-Path -LiteralPath $case[0]).Path,128)
        $actual = Read-Rename $document
        foreach ($key in @('id','name','nameU','cachedName','formula','subAddress')) {
            if ($actual[$key] -ne $case[1][$key]) { throw "Incorrect native rename result: $key" }
        }
        $accepted += $actual
        $document.SaveAs((Join-Path $OutputDirectory $case[2])) | Out-Null
        $document.Saved = $true; $document.Close(); $document = $null
    }
    [ordered]@{ application='Microsoft Visio'; version=$app.Version; first=$first;
        second=$second; accepted=$accepted } | ConvertTo-Json -Depth 5 |
        Set-Content (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
} finally {
    if ($document) { $document.Saved = $true; $document.Close() }
    $app.Quit()
}
Write-Output $OutputDirectory
