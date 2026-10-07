param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-chart-types'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = 'Chart data'
    $rows = @(@('Month', 'Sales', 'Cost'), @('Jan', 10, 4), @('Feb', 20, 8), @('Mar', 30, 9))
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $rows[$r].Count; $c++) {
            $value = $rows[$r][$c]
            if ($value -is [int]) {
                $sheet.Cells.Item($r + 1, $c + 1).Value2 = [double]$value
            } else {
                $sheet.Cells.Item($r + 1, $c + 1).Value2 = [string]$value
            }
        }
    }
    $chart = $sheet.ChartObjects().Add(260, 20, 480, 300).Chart
    $chart.SetSourceData($sheet.Range('A1:C4'), 2)
    $cases = @()
    foreach ($entry in @(
        @('column', 'clustered', 51), @('column', 'stacked', 52), @('column', 'percentStacked', 53),
        @('bar', 'clustered', 57), @('bar', 'stacked', 58), @('bar', 'percentStacked', 59)
    )) {
        $chart.ChartType = $entry[2]
        $chart.HasTitle = $true
        $chart.ChartTitle.Text = 'Keep me'
        $path = Join-Path $OutputFolder ($entry[0] + '-' + $entry[1] + '.xlsx')
        $book.SaveCopyAs($path)
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            $part = $zip.GetEntry('xl/charts/chart1.xml')
            $reader = [System.IO.StreamReader]::new($part.Open())
            try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        } finally { $zip.Dispose() }
        $series = @()
        for ($index = 1; $index -le $chart.SeriesCollection().Count; $index++) {
            $item = $chart.SeriesCollection($index)
            $series += @{ name = [string]$item.Name; formula = [string]$item.Formula; values = @($item.Values) }
        }
        $cases += @{ type = $entry[0]; grouping = $entry[1]; nativeType = [int]$chart.ChartType; title = [string]$chart.ChartTitle.Text; series = $series; xml = $xml }
    }
    @{ excelVersion = [string]$excel.Version; excelBuild = [string]$excel.Build; cases = $cases } |
        ConvertTo-Json -Depth 10 | Set-Content -Encoding utf8 (Join-Path $OutputFolder 'chart-types.json')
    Write-Output "Recorded $($cases.Count) native chart type/grouping transitions in $OutputFolder"
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}
