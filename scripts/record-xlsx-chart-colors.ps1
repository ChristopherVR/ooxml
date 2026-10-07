param(
    [string]$OutputFile = (Join-Path $env:TEMP 'excel-chart-colors.json'),
    [switch]$CustomTheme
)
$ErrorActionPreference = 'Stop'
function HexColor([int]$rgb) {
    return '#{0:X2}{1:X2}{2:X2}' -f ($rgb -band 255), (($rgb -shr 8) -band 255), (($rgb -shr 16) -band 255)
}
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    if ($CustomTheme) {
        $custom = @('234567', 'FEF4E6', '782345', 'C3DAC7', 'EC174D', '035DA8', '19C743', '8426D2', 'DA9308', '08ADAA', '2959DA', 'BC399F')
        for ($i = 0; $i -lt $custom.Count; $i++) {
            $hex = $custom[$i]
            $rgb = [Convert]::ToInt32($hex.Substring(0, 2), 16) +
                ([Convert]::ToInt32($hex.Substring(2, 2), 16) -shl 8) +
                ([Convert]::ToInt32($hex.Substring(4, 2), 16) -shl 16)
            $book.Theme.ThemeColorScheme.Colors($i + 1).RGB = [int]$rgb
        }
    }
    $scheme = @{}
    $names = @('dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink')
    for ($i = 0; $i -lt $names.Count; $i++) {
        $scheme[$names[$i]] = HexColor $book.Theme.ThemeColorScheme.Colors($i + 1).RGB
    }
    $cases = @()
    foreach ($count in @(1, 2, 4, 7, 10, 19, 55)) {
        $sheet.Cells.Clear() | Out-Null
        $sheet.Cells.Item(1, 1).Value2 = 'Month'
        $sheet.Cells.Item(2, 1).Value2 = 'Jan'
        $sheet.Cells.Item(3, 1).Value2 = 'Feb'
        for ($i = 1; $i -le $count; $i++) {
            $sheet.Cells.Item(1, $i + 1).Value2 = [string]"Series $i"
            $sheet.Cells.Item(2, $i + 1).Value2 = [double]($i * 3)
            $sheet.Cells.Item(3, $i + 1).Value2 = [double]($i * 5)
        }
        $object = $sheet.ChartObjects().Add(260, 20, 480, 300)
        $chart = $object.Chart
        $chart.ChartType = 51
        $chart.ChartStyle = 2
        $chart.SetSourceData($sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item(3, $count + 1)), 2)
        for ($palette = 10; $palette -le 26; $palette++) {
            $chart.ChartColor = $palette
            $colors = @()
            for ($i = 1; $i -le $chart.SeriesCollection().Count; $i++) {
                $colors += HexColor $chart.SeriesCollection($i).Format.Fill.ForeColor.RGB
            }
            $cases += @{ palette = $palette; seriesCount = $count; colors = $colors }
        }
        $object.Delete()
    }
    @{ excelVersion = [string]$excel.Version; excelBuild = [string]$excel.Build; scheme = $scheme; cases = $cases } |
        ConvertTo-Json -Depth 10 | Set-Content -Encoding utf8 $OutputFile
    Write-Output "Recorded $($cases.Count) native chart palettes in $OutputFile"
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}
