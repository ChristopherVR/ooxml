param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [int]$PageId = 99,
    [string]$PageName = 'Inserted & Page',
    [double]$Width = 8.5,
    [double]$Height = 11,
    [double]$DrawingScale = 2,
    [double]$PageScale = 1,
    [int]$ExpectedIndex = 5,
    [int]$ExpectedPageCount = 8
)
$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $InputPath).Path
$inputHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$destination = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $destination | Out-Null
$app = $null
$document = $null
try {
    $app = New-Object -ComObject Visio.InvisibleApp
    $app.AlertResponse = 7
    $document = $app.Documents.OpenEx($source, 128)
    $page = $document.Pages.ItemFromID($PageId)
    $values = [ordered]@{
        VisioVersion = $app.Version
        InputSha256 = $inputHash
        PageId = $page.ID
        Name = $page.Name
        Index = $page.Index
        PageCount = $document.Pages.Count
        ShapeCount = $page.Shapes.Count
    }
    if ($page.Name -ne $PageName -or $page.Shapes.Count -ne 0) { throw 'Inserted page name or blank contents differ.' }
    if ($page.Index -ne $ExpectedIndex -or $document.Pages.Count -ne $ExpectedPageCount) { throw 'Inserted page order or count differs.' }
    foreach ($entry in @{ PageWidth = $Width; PageHeight = $Height; DrawingScale = $DrawingScale; PageScale = $PageScale }.GetEnumerator()) {
        $actual = $page.PageSheet.CellsU($entry.Key).ResultIU
        $values[$entry.Key] = $actual
        if ([Math]::Abs($actual - $entry.Value) -gt 1e-10) { throw "Native page value differs: $($entry.Key)" }
    }
    $document.SaveAs((Join-Path $destination 'native-reopened-insertion.vsdx'))
    $values | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $destination 'page-insertion-evidence.json')
    $values | ConvertTo-Json
} finally {
    if ($document) { $document.Saved = $true; $document.Close() }
    if ($app) { $app.Quit() }
}
