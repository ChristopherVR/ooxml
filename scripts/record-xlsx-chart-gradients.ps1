param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-native-chart-gradients'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel = $null; $book = $null
function Hex-Color($rgb) {
    return '#{0:X2}{1:X2}{2:X2}' -f ($rgb -band 255),(($rgb -shr 8) -band 255),(($rgb -shr 16) -band 255)
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add(); $sheet = $book.Worksheets.Item(1)
    $sheet.Cells.Item(1,1).Value2 = [string]'Month'
    $sheet.Cells.Item(1,2).Value2 = [string]'Sales'
    $sheet.Cells.Item(1,3).Value2 = [string]'Cost'
    for ($row=2; $row -le 5; $row++) {
        $sheet.Cells.Item($row,1).Value2 = [string]"M$row"
        $sheet.Cells.Item($row,2).Value2 = [double]($row*2)
        $sheet.Cells.Item($row,3).Value2 = [double]$row
    }
    $scheme = @{}
    $names = @('dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink')
    for ($i=0; $i -lt $names.Count; $i++) { $scheme[$names[$i]] = Hex-Color ([int]$book.Theme.ThemeColorScheme.Colors($i+1).RGB) }
    $cases = @()
    foreach ($palette in @(10,12,14)) {
        $object = $sheet.ChartObjects().Add(260,20,480,300)
        $chart = $object.Chart; $chart.ChartType = 51
        $chart.SetSourceData($sheet.Range('A1:C5'),2)
        $chart.ChartStyle = 209; $chart.ChartColor = $palette
        $chart.HasTitle = $true; $chart.ChartTitle.Text = 'Native style'
        $path = Join-Path $OutputFolder "gradient-$palette.xlsx"
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
        $probe = $excel.Workbooks.Open($path,0,$true)
        try {
            $native = $probe.Worksheets.Item(1).ChartObjects(1).Chart
            $series = @()
            for ($s=1; $s -le $native.SeriesCollection().Count; $s++) {
                $stops = $native.SeriesCollection($s).Format.Fill.GradientStops
                $colors = @()
                for ($n=1; $n -le $stops.Count; $n++) {
                    $stop = $stops.Item($n)
                    $colors += @{ color = (Hex-Color ([int]$stop.Color.RGB)); position = [double]$stop.Position; transparency = [double]$stop.Transparency }
                }
                $series += @{ stops = $colors }
            }
            $cases += @{ palette = [int]$native.ChartColor; series = $series; parts = $parts }
            $native.Export((Join-Path $OutputFolder "gradient-$palette.png"),'PNG') | Out-Null
        } finally { $probe.Close($false) }
        $object.Delete()
    }
    @{ excelVersion = [string]$excel.Version; excelBuild = [string]$excel.Build; scheme = $scheme; cases = $cases } |
        ConvertTo-Json -Depth 10 | Set-Content -Encoding utf8 (Join-Path $OutputFolder 'gradients.json')
    Write-Output "Recorded $($cases.Count) independent native gradient charts in $OutputFolder"
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}
