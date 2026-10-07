param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-native-chart-styles'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel = $null; $book = $null
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
    for ($i=0; $i -lt $names.Count; $i++) { $rgb=[int]$book.Theme.ThemeColorScheme.Colors($i+1).RGB; $scheme[$names[$i]] = '#{0:X2}{1:X2}{2:X2}' -f ($rgb -band 255),(($rgb -shr 8) -band 255),(($rgb -shr 16) -band 255) }
    $cases = @()
    foreach ($id in 201..216) {
        $object = $sheet.ChartObjects().Add(260,20,480,300)
        $chart = $object.Chart; $chart.ChartType = 51
        $chart.SetSourceData($sheet.Range('A1:C5'),2)
        $chart.ChartStyle = $id; $chart.ChartColor = 10
        $chart.HasTitle = $true; $chart.ChartTitle.Text = 'Native style'
        $path = Join-Path $OutputFolder "column-$id.xlsx"
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
        $probeBook = $excel.Workbooks.Open($path,0,$true)
        try {
            $probeChart = $probeBook.Worksheets.Item(1).ChartObjects(1).Chart
            $cases += @{ requestedStyle = $id; style = [int]$probeChart.ChartStyle; titleFontSize = $(if ([double]$probeChart.ChartTitle.Font.Size -gt 0) { [double]$probeChart.ChartTitle.Font.Size } else { $null }); axisFontSize = [double]$probeChart.Axes(1).TickLabels.Font.Size; legendFontSize = [double]$probeChart.Legend.Font.Size; parts = $parts }
        } finally { $probeBook.Close($false) }
        $object.Delete()
    }
    @{ excelVersion = [string]$excel.Version; excelBuild = [string]$excel.Build; scheme = $scheme; cases = $cases } |
        ConvertTo-Json -Depth 8 | Set-Content -Encoding utf8 (Join-Path $OutputFolder 'styles.json')
    Write-Output "Recorded $($cases.Count) independent native style parts in $OutputFolder"
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}
