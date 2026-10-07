param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-chart-palette-edits'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
function HexColor([int]$rgb) { return '#{0:X2}{1:X2}{2:X2}' -f ($rgb -band 255), (($rgb -shr 8) -band 255), (($rgb -shr 16) -band 255) }
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = 'Chart data'
    $sheet.Cells.Item(1,1).Value2 = 'Month'
    $months = @('Jan','Feb','Mar','Apr')
    for ($r = 0; $r -lt 4; $r++) { $sheet.Cells.Item($r+2,1).Value2 = [string]$months[$r] }
    for ($s = 1; $s -le 4; $s++) {
        $sheet.Cells.Item(1,$s+1).Value2 = [string]"Series $s"
        for ($r = 2; $r -le 5; $r++) { $sheet.Cells.Item($r,$s+1).Value2 = [double]($r * $s) }
    }
    $scheme = @{}
    $names = @('dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink')
    for ($i=0; $i -lt $names.Count; $i++) { $scheme[$names[$i]] = HexColor $book.Theme.ThemeColorScheme.Colors($i+1).RGB }
    $cases = @()
    foreach ($kind in @(@('column',51),@('bar',57),@('line',4),@('area',1),@('pie',5),@('doughnut',-4120),@('scatter',-4169),@('radar',-4151))) {
        $object = $sheet.ChartObjects().Add(260,20,480,300)
        $chart = $object.Chart
        $chart.ChartType = $kind[1]
        $chart.SetSourceData($sheet.Range($(if ($kind[0] -in @('pie','doughnut')) {'A1:B5'} else {'A1:E5'})),2)
        $chart.ChartStyle = 2
        $chart.ChartColor = 10
        $chart.HasTitle = $true
        $chart.ChartTitle.Text = 'Keep me'
        foreach ($state in @('automatic','automaticChanged','manual','changed')) {
            $lineLike = $kind[0] -in @('line','scatter','radar')
            if ($state -eq 'manual') {
                $chart.ChartColor = 10
                if ($lineLike) { $chart.SeriesCollection(1).Format.Line.ForeColor.RGB = 255 }
                else { $chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = 255 }
                $chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = 65280
            }
            if ($state -in @('changed','automaticChanged')) { $chart.ChartColor = 12 }
            $path = Join-Path $OutputFolder "$($kind[0])-$state.xlsx"
            $book.SaveCopyAs($path)
            $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
            try {
                $parts = @{}
                foreach ($entry in $zip.Entries) {
                    if ($entry.FullName -notmatch '^xl/charts/') { continue }
                    $reader = [System.IO.StreamReader]::new($entry.Open())
                    try { $parts[$entry.FullName] = $reader.ReadToEnd() } finally { $reader.Dispose() }
                }
            } finally { $zip.Dispose() }
            $colors = @()
            $probeBook = $excel.Workbooks.Open($path,0,$true)
            try {
            $probeChart = $probeBook.Worksheets.Item(1).ChartObjects(1).Chart
            for ($i=1; $i -le $probeChart.SeriesCollection().Count; $i++) {
                $series = $probeChart.SeriesCollection($i)
                $primary = if ($kind[0] -eq 'scatter') { $series.MarkerForegroundColor } elseif ($lineLike) { $series.Format.Line.ForeColor.RGB } else { $series.Format.Fill.ForeColor.RGB }
                $points = @()
                if ($kind[0] -in @('pie','doughnut','column','bar','scatter')) {
                    for ($p=1; $p -le $series.Points().Count; $p++) {
                        if ($kind[0] -eq 'scatter') {
                            $rgb = $series.Points($p).MarkerBackgroundColor
                            if ($rgb -lt 0) { $rgb = $series.Points($p).MarkerForegroundColor }
                        } else { $rgb = $series.Points($p).Format.Fill.ForeColor.RGB }
                        $points += HexColor $rgb
                    }
                }
                $colors += @{ primary = HexColor $primary; points = $points }
            }
            } finally { $probeBook.Close($false) }
            $cases += @{ type = $kind[0]; state = $state; palette = [int]$chart.ChartColor; colors = $colors; parts = $parts }
        }
        $object.Delete()
    }
    @{ excelVersion = [string]$excel.Version; excelBuild = [string]$excel.Build; scheme = $scheme; cases = $cases } |
        ConvertTo-Json -Depth 10 | Set-Content -Encoding utf8 (Join-Path $OutputFolder 'palette-edits.json')
    Write-Output "Recorded $($cases.Count) native chart states in $OutputFolder"
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}
