param(
    [string]$InputFile = (Join-Path $PSScriptRoot '../src/core/xlsx/__fixtures__/excel-smartart.xlsx'),
    [string]$OutputFile = (Join-Path $env:TEMP 'excel-smartart-appearance.json')
)
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $InputFile).Path, 0, $true)
    $diagrams = @()
    foreach ($sheet in $book.Worksheets) {
        foreach ($shape in $sheet.Shapes) {
            if ($shape.HasSmartArt -ne -1) { continue }
            $nodes = @()
            for ($index = 1; $index -le $shape.SmartArt.AllNodes.Count; $index++) {
                $node = $shape.SmartArt.AllNodes.Item($index)
                $font = $node.TextFrame2.TextRange.Font
                $rgb = [int]$font.Fill.ForeColor.RGB
                $hex = '{0:X2}{1:X2}{2:X2}' -f ($rgb -band 255), (($rgb -shr 8) -band 255), (($rgb -shr 16) -band 255)
                $nodes += @{ text = [string]$node.TextFrame2.TextRange.Text; fontSizePt = [double]$font.Size; fontName = [string]$font.Name; fontColor = "#$hex" }
            }
            $diagrams += @{ name = [string]$shape.Name; layout = [string]$shape.SmartArt.Layout.Name; nodes = $nodes }
        }
    }
    if ($diagrams.Count -eq 0) { throw 'No native SmartArt found' }
    @{ excelVersion = [string]$excel.Version; excelBuild = [string]$excel.Build; diagrams = $diagrams } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputFile -Encoding utf8
    Write-Output "Recorded $($diagrams.Count) native SmartArt appearance records in $OutputFile"
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}
